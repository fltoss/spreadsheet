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
          "above the :max_cells limit of 10000000"

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
                  "above the :max_cells limit of 10000000"}
    end

    test ".xls: a cell at row 65536, column 65536" do
      message =
        ~s|sheet "Sheet1" has 65536 rows × 65536 columns (4294967296 cells), | <>
          "above the :max_cells limit of 10000000"

      assert Spreadsheet.parse(fixture("sparse.xls")) == {:error, message}

      # .xls workbooks are decoded when opened, so listing sheets is checked too.
      assert Spreadsheet.sheet_names(fixture("sparse.xls")) == {:error, message}

      content = File.read!(fixture("sparse.xls"))

      assert Spreadsheet.sheet_names(content, format: :binary) ==
               {:error, message}
    end

    test ".xls: a forged DIMENSIONS record" do
      assert {:error, message} = Spreadsheet.parse(fixture("dimension.xls"))

      assert message =~
               ~s|sheet "Sheet1" declares 4294967295 rows × 255 columns in its DIMENSIONS record|
    end

    test ".xls: a forged compound file header" do
      assert Spreadsheet.parse(fixture("fat_length.xls")) ==
               {:error,
                "invalid .xls file: the header declares 4294967295 FAT sectors but the file has 11"}
    end

    test ".ods: values in A1 and XFD1048576" do
      message =
        ~s|sheet "Sparse" has 1048576 rows × 16384 columns (17179869184 cells), | <>
          "above the :max_cells limit of 10000000"

      assert Spreadsheet.parse(fixture("sparse.ods")) == {:error, message}
      assert Spreadsheet.sheet_names(fixture("sparse.ods")) == {:error, message}
    end

    test ".ods: repeated columns" do
      assert Spreadsheet.parse(fixture("repeated_columns.ods")) ==
               {:error,
                ~s|sheet "Repeated" has 10000 rows × 16384 columns (163840000 cells), | <>
                  "above the :max_cells limit of 10000000"}
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

    test "applies to each sheet" do
      path = fixture("types.xlsb")

      assert {:error, message} = Spreadsheet.parse(path, max_cells: 8)
      assert message =~ ~s(sheet "Types")

      assert {:ok, [["hidden"]]} =
               Spreadsheet.parse(path, sheet: "Hidden", max_cells: 1)
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
