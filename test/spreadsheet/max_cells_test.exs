defmodule Spreadsheet.MaxCellsTest do
  use ExUnit.Case, async: true

  @fixtures Path.expand("../fixtures", __DIR__)

  defp fixture(name), do: Path.join(@fixtures, name)

  # Each of these is a few KB but would make calamine allocate gigabytes (or
  # terabytes) and abort the VM if it were parsed without the limit.
  describe "sparse sheets" do
    test ".xlsx: values in A1 and XFD200000" do
      message =
        ~s|sheet "Sparse" has 200000 rows × 16384 columns (3276800000 cells), | <>
          "above the :max_cells limit of 50000000"

      assert Spreadsheet.parse(fixture("sparse.xlsx")) == {:error, message}

      assert Spreadsheet.parse(fixture("sparse.xlsx"), sheet: "Sparse") ==
               {:error, message}

      content = File.read!(fixture("sparse.xlsx"))

      assert Spreadsheet.parse(content, format: :binary) == {:error, message}

      # Listing sheets does not read cells.
      assert Spreadsheet.sheet_names(fixture("sparse.xlsx")) ==
               {:ok, ["Sparse"]}
    end

    test ".xlsb: values in A1 and XFD1000001" do
      assert Spreadsheet.parse(fixture("sparse.xlsb")) ==
               {:error,
                ~s|sheet "Sparse" has 1000001 rows × 16384 columns (16384016384 cells), | <>
                  "above the :max_cells limit of 50000000"}
    end

    test ".xls: a cell at row 65536, column 65536" do
      message =
        ~s|sheet "Sheet1" has 65536 rows × 65536 columns (4294967296 cells), | <>
          "above the :max_cells limit of 50000000"

      assert Spreadsheet.parse(fixture("sparse.xls")) == {:error, message}

      # .xls workbooks are decoded when opened, so listing sheets is checked too.
      assert Spreadsheet.sheet_names(fixture("sparse.xls")) == {:error, message}

      content = File.read!(fixture("sparse.xls"))

      assert Spreadsheet.sheet_names(content, format: :binary) ==
               {:error, message}
    end

    test ".xls: a forged DIMENSIONS record, whatever :max_cells is" do
      message =
        ~s|invalid .xls file: sheet "Sheet1" declares 4294967295 rows × 255 columns | <>
          "in its DIMENSIONS record, above the .xls maximum of 65536 rows × 256 columns"

      for max_cells <- [10_000_000, :infinity] do
        assert Spreadsheet.parse(fixture("dimension.xls"), max_cells: max_cells) ==
                 {:error, message}
      end
    end

    test ".xls: declared dimensions within .xls limits do not count" do
      # A1:IV50000 declared (12.8 million cells), two values in A1 and B2.
      rows = [[1.0, nil], [nil, 7.0]]

      assert Spreadsheet.parse(fixture("template.xls")) ==
               {:ok, [{"Sheet1", rows}]}

      assert Spreadsheet.parse(fixture("template.xls"), max_cells: 4) ==
               {:ok, [{"Sheet1", rows}]}
    end

    test ".xls: a forged compound file header, whatever :max_cells is" do
      for max_cells <- [10_000_000, :infinity] do
        assert Spreadsheet.parse(fixture("fat_length.xls"),
                 max_cells: max_cells
               ) ==
                 {:error,
                  "invalid .xls file: the header declares 4294967295 FAT sectors but the file has 11"}
      end
    end

    test ".xlsx: repeated cell positions" do
      path = fixture("repeated_cells.xlsx")

      # The last of repeated cells wins.
      assert Spreadsheet.parse(path, sheet: "Once") == {:ok, [[3.0, 2.0]]}

      for max_cells <- [50_000_000, :infinity] do
        assert Spreadsheet.parse(path, sheet: "Many", max_cells: max_cells) ==
                 {:error,
                  ~s|invalid sheet "Many": 3 cell records repeat positions within | <>
                    "1 rows × 1 columns"}
      end
    end

    test ".ods: values in A1 and XFD1048576" do
      message =
        ~s|sheet "Sparse" has 1048576 rows × 16384 columns (17179869184 cells), | <>
          "above the :max_cells limit of 50000000"

      assert Spreadsheet.parse(fixture("sparse.ods")) == {:error, message}
      assert Spreadsheet.sheet_names(fixture("sparse.ods")) == {:error, message}
    end

    test ".ods: repeated columns" do
      assert Spreadsheet.parse(fixture("repeated_columns.ods")) ==
               {:error,
                ~s|sheet "Repeated" has 10000 rows × 16384 columns (163840000 cells), | <>
                  "above the :max_cells limit of 50000000"}
    end
  end

  describe ":max_cells" do
    # test_file_1 has 8 rows × 4 columns in every format.
    for file <- ~w(test_file_1.xlsx test_file_1.xls test_file_1.ods) do
      test "#{file}: limits rows × columns" do
        path = fixture(unquote(file))

        assert {:ok, [{"Sheet1", rows}]} =
                 Spreadsheet.parse(path, max_cells: 32)

        assert length(rows) == 8

        assert Spreadsheet.parse(path, max_cells: 31) ==
                 {:error,
                  ~s|sheet "Sheet1" has 8 rows × 4 columns (32 cells), | <>
                    "above the :max_cells limit of 31"}

        assert {:ok, [{"Sheet1", ^rows}]} =
                 Spreadsheet.parse(path, max_cells: :infinity)
      end
    end

    test "counts the rectangle from A1" do
      # Data in C3:E6 makes a 6 × 5 rectangle once anchored at A1.
      path = fixture("offset.xlsx")

      assert {:ok, _rows} = Spreadsheet.parse(path, max_cells: 30)
      assert {:error, message} = Spreadsheet.parse(path, max_cells: 29)
      assert message =~ "has 6 rows × 5 columns (30 cells)"
    end

    test "is a budget for the whole call" do
      path = fixture("types.xlsb")

      # Types is 9 × 2, Hidden 1 × 1, Offset 3 × 3: 28 cells in all.
      assert {:ok, [_types, _hidden, _offset]} =
               Spreadsheet.parse(path, max_cells: 28)

      assert Spreadsheet.parse(path, max_cells: 27) ==
               {:error,
                ~s|sheet "Offset" has 3 rows × 3 columns (9 cells), above the 8 cells left | <>
                  "of the :max_cells limit of 27 after 19 cells in earlier sheets"}

      assert {:error, message} = Spreadsheet.parse(path, max_cells: 17)
      assert message =~ ~s(sheet "Types" has 9 rows × 2 columns)

      assert {:ok, [["hidden"]]} =
               Spreadsheet.parse(path, sheet: "Hidden", max_cells: 1)
    end

    # sheets.* hold four 1 × 1 sheets, one of them very hidden.
    test ".xlsx: only the sheets parsed count" do
      path = fixture("sheets.xlsx")

      assert {:ok, [_, _, _, _]} = Spreadsheet.parse(path, max_cells: 4)
      assert {:error, message} = Spreadsheet.parse(path, max_cells: 3)
      assert message =~ ~s(sheet "Very Hidden" has 1 rows × 1 columns)

      assert {:ok, [_, _]} =
               Spreadsheet.parse(path, max_cells: 2, hidden: false)

      assert {:ok, [["visible"]]} =
               Spreadsheet.parse(path, sheet: "Visible", max_cells: 1)

      assert {:ok, [_, _, _, _]} = Spreadsheet.sheet_names(path, max_cells: 1)
    end

    for file <- ~w(sheets.xls sheets.ods) do
      test "#{file}: every sheet counts, since the workbook is decoded when opened" do
        path = fixture(unquote(file))

        assert {:ok, [_, _, _, _]} = Spreadsheet.sheet_names(path, max_cells: 4)

        assert {:ok, [["visible"]]} =
                 Spreadsheet.parse(path, sheet: "Visible", max_cells: 4)

        message =
          ~s|sheet "Very Hidden" has 1 rows × 1 columns (1 cells), above the 0 cells | <>
            "left of the :max_cells limit of 3 after 3 cells in earlier sheets"

        assert Spreadsheet.sheet_names(path, max_cells: 3) == {:error, message}

        assert Spreadsheet.parse(path, sheet: "Visible", max_cells: 3) ==
                 {:error, message}
      end
    end

    test "must be a positive integer or :infinity" do
      path = fixture("test_file_1.xlsx")

      for value <- [0, -1, 1.5, "10", nil] do
        assert_raise ArgumentError,
                     ~r/expected :max_cells to be a positive integer or :infinity/,
                     fn -> Spreadsheet.parse(path, max_cells: value) end

        assert_raise ArgumentError, ~r/expected :max_cells/, fn ->
          Spreadsheet.sheet_names(path, max_cells: value)
        end
      end
    end
  end
end
