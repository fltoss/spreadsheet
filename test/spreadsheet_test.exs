defmodule SpreadsheetTest do
  use ExUnit.Case

  @base_path Path.join(__DIR__, "fixtures")

  describe "sheet_names/2" do
    test "gets the sheet_names from binary content" do
      content = File.read!(Path.join(@base_path, "test_file_1.xlsx"))

      assert Spreadsheet.sheet_names(content, format: :binary) ==
               {:ok, ["Sheet1"]}
    end

    test "gets the sheet_names from a path (default)" do
      path = Path.join(@base_path, "test_file_1.xlsx")

      assert Spreadsheet.sheet_names(path) == {:ok, ["Sheet1"]}
    end

    test "gets the sheet_names from a path (explicit)" do
      path = Path.join(@base_path, "test_file_1.xlsx")

      assert Spreadsheet.sheet_names(path, format: :filename) ==
               {:ok, ["Sheet1"]}
    end

    test "ignores hidden sheet names" do
      path = Path.join(@base_path, "test_file_1_hidden.xlsx")

      assert Spreadsheet.sheet_names(path) == {:ok, ["Sheet2", "Sheet1"]}
      assert Spreadsheet.sheet_names(path, hidden: false) == {:ok, ["Sheet1"]}

      path = Path.join(@base_path, "test_file_1_hidden_2.xlsx")

      assert Spreadsheet.sheet_names(path) == {:ok, ["Sheet1", "foobar"]}
      assert Spreadsheet.sheet_names(path, hidden: false) == {:ok, ["Sheet1"]}
    end

    test "gets the sheet_names from a path for xls" do
      path = Path.join(@base_path, "test_file_1.xls")

      assert Spreadsheet.sheet_names(path) == {:ok, ["Sheet1"]}
    end

    test "gets the sheet_names from a path for ods" do
      path = Path.join(@base_path, "test_file_1.ods")

      assert Spreadsheet.sheet_names(path) == {:ok, ["Sheet1"]}
    end

    test "reads xls files from binary content" do
      content = File.read!(Path.join(@base_path, "test_file_1.xls"))

      assert Spreadsheet.sheet_names(content, format: :binary) ==
               {:ok, ["Sheet1"]}
    end

    test "reads ods files from binary content" do
      content = File.read!(Path.join(@base_path, "test_file_1.ods"))

      assert Spreadsheet.sheet_names(content, format: :binary) ==
               {:ok, ["Sheet1"]}
    end

    test "returns error for invalid format option" do
      path = Path.join(@base_path, "test_file_1.xlsx")

      assert Spreadsheet.sheet_names(path, format: :invalid) ==
               {:error,
                "Invalid format option: :invalid. Expected :filename or :binary"}
    end
  end

  describe "parse/2 (all sheets)" do
    test "parses all sheets from a file path" do
      path = Path.join(@base_path, "test_file_1.xlsx")

      assert {:ok, sheets} = Spreadsheet.parse(path)
      assert length(sheets) == 1
      assert [{"Sheet1", [header | _rows]}] = sheets
      assert header == ["Dates", "Numbers", "Percentages", "Strings"]
    end

    test "parses all sheets from binary content" do
      content = File.read!(Path.join(@base_path, "test_file_1.xlsx"))

      assert {:ok, sheets} = Spreadsheet.parse(content, format: :binary)
      assert length(sheets) == 1
      assert [{"Sheet1", [header | _rows]}] = sheets
      assert header == ["Dates", "Numbers", "Percentages", "Strings"]
    end

    test "parses all sheets excluding hidden ones" do
      path = Path.join(@base_path, "test_file_1_hidden.xlsx")

      # With hidden sheets
      assert {:ok, all_sheets} = Spreadsheet.parse(path)
      assert length(all_sheets) == 2
      assert [{"Sheet2", _}, {"Sheet1", _}] = all_sheets

      # Without hidden sheets
      assert {:ok, visible_sheets} = Spreadsheet.parse(path, hidden: false)
      assert length(visible_sheets) == 1
      assert [{"Sheet1", _}] = visible_sheets
    end

    test "parses all sheets from xls file" do
      path = Path.join(@base_path, "test_file_1.xls")

      assert {:ok, sheets} = Spreadsheet.parse(path)
      assert length(sheets) == 1
      assert [{"Sheet1", [header | _rows]}] = sheets
      assert header == ["Dates", "Numbers", "Percentages", "Strings"]
    end

    test "parses all sheets from ods file" do
      path = Path.join(@base_path, "test_file_1.ods")

      assert {:ok, sheets} = Spreadsheet.parse(path)
      assert length(sheets) == 1
      assert [{"Sheet1", [header | _rows]}] = sheets
      assert header == ["Dates", "Numbers", "Percentages", "Strings"]
    end
  end

  describe "option validation" do
    setup do
      %{path: Path.join(@base_path, "test_file_1.xlsx")}
    end

    test "unknown options raise", %{path: path} do
      assert_raise ArgumentError, ~r/unknown keys \[:sheets\]/, fn ->
        Spreadsheet.parse(path, sheets: "Sheet1")
      end

      assert_raise ArgumentError, ~r/unknown keys \[:sheet\]/, fn ->
        Spreadsheet.sheet_names(path, sheet: "Sheet1")
      end
    end

    test "a non-string :sheet raises", %{path: path} do
      assert_raise ArgumentError,
                   ~r/expected :sheet to be a string.*got: :Sheet1/,
                   fn ->
                     Spreadsheet.parse(path, sheet: :Sheet1)
                   end
    end

    test "sheet: nil parses every sheet", %{path: path} do
      assert {:ok, [{"Sheet1", _rows}]} = Spreadsheet.parse(path, sheet: nil)
    end

    test "a non-boolean :hidden raises", %{path: path} do
      assert_raise ArgumentError,
                   ~r/expected :hidden to be a boolean, got: "no"/,
                   fn ->
                     Spreadsheet.parse(path, hidden: "no")
                   end

      assert_raise ArgumentError, ~r/expected :hidden to be a boolean/, fn ->
        Spreadsheet.sheet_names(path, hidden: nil)
      end
    end
  end

  describe "file errors" do
    test "a missing file reports the I/O error" do
      path = Path.join(@base_path, "does_not_exist.xlsx")

      assert Spreadsheet.parse(path) ==
               {:error,
                "could not open #{inspect(path)}: No such file or directory"}

      assert {:error, "could not open " <> _} = Spreadsheet.sheet_names(path)
    end

    test "a directory is not a spreadsheet" do
      assert Spreadsheet.parse(@base_path) ==
               {:error, "could not open #{inspect(@base_path)}: is a directory"}
    end

    test "file contents passed as a path return a hint instead of raising" do
      content = File.read!(Path.join(@base_path, "test_file_1.xlsx"))

      assert {:error, message} = Spreadsheet.parse(content)
      assert message =~ "did you mean format: :binary?"

      assert {:error, message} = Spreadsheet.sheet_names(<<0xFF, 0, 0xFE>>)
      assert message =~ "did you mean format: :binary?"
    end

    test "a path passed as content returns a hint" do
      path = Path.join(@base_path, "test_file_1.xlsx")

      assert {:error, message} = Spreadsheet.parse(path, format: :binary)
      assert message =~ "did you mean format: :filename?"
    end

    test "non-UTF-8 paths are accepted" do
      assert {:error, message} =
               Spreadsheet.parse(<<"missing-", 0xE9, ".xlsx">>)

      assert message =~ "No such file or directory"
    end

    test "a missing sheet returns an error naming the sheet" do
      path = Path.join(@base_path, "test_file_1.xlsx")

      assert {:error, message} = Spreadsheet.parse(path, sheet: "Nope")
      assert message =~ "'Nope'"
    end

    test "content that is not a spreadsheet cannot be detected" do
      assert Spreadsheet.parse("plain text", format: :binary) ==
               {:error,
                "Cannot detect file format: not a valid .xls, .xlsx, .xlsb or .ods file"}
    end
  end

  describe "parse/2 (specific sheet)" do
    test "parses from binary content" do
      content = File.read!(Path.join(@base_path, "test_file_1.xlsx"))
      sheet_name = "Sheet1"

      {:ok, [header | rows]} =
        Spreadsheet.parse(content, sheet: sheet_name, format: :binary)

      assert header == ["Dates", "Numbers", "Percentages", "Strings"]

      assert rows == [
               [~N[2024-12-12 00:00:00], 1234.0, 0.12, "Foobar"],
               [~N[1993-11-21 00:00:00], "00012345", 0.1212, nil],
               [~N[1987-05-08 00:00:00], 1122.0, "12", nil],
               [~N[1994-05-22 00:00:00], "12,12", 12.0, nil],
               ["2024-01-01", 11.12, "33.12%", "123"],
               [~N[1987-05-08 20:10:12], nil, nil, nil],
               [~N[1987-05-08 20:10:12], nil, nil, nil]
             ]
    end

    test "parses from path for xls" do
      path = Path.join(@base_path, "test_file_1.xls")
      sheet_name = "Sheet1"

      {:ok, [header | rows]} = Spreadsheet.parse(path, sheet: sheet_name)

      assert header == ["Dates", "Numbers", "Percentages", "Strings"]

      assert rows == [
               [~N[2024-12-12 00:00:00], 1234.0, 0.12, "Foobar"],
               [
                 ~N[1993-11-21 00:00:00],
                 "00012345",
                 0.12119999999999999,
                 nil
               ],
               [~N[1987-05-08 00:00:00], 1122.0, "12", nil],
               [~N[1994-05-22 00:00:00], "12,12", 12.0, nil],
               ["2024-01-01", 11.12, "33.12%", "123"],
               [~N[1987-05-08 20:10:12], nil, nil, nil],
               [~N[1987-05-08 20:10:12], nil, nil, nil]
             ]
    end

    test "parses from path" do
      path = Path.join(@base_path, "test_file_1.xlsx")
      sheet_name = "Sheet1"

      assert {:ok, _} = Spreadsheet.parse(path, sheet: sheet_name)
    end

    test "parses from path with explicit format" do
      path = Path.join(@base_path, "test_file_1.xlsx")
      sheet_name = "Sheet1"

      assert {:ok, _} =
               Spreadsheet.parse(path, sheet: sheet_name, format: :filename)
    end

    test "parses xls files from binary content" do
      content = File.read!(Path.join(@base_path, "test_file_1.xls"))

      assert Spreadsheet.parse(content, sheet: "Sheet1", format: :binary) == {
               :ok,
               [
                 ["Dates", "Numbers", "Percentages", "Strings"],
                 [~N[2024-12-12 00:00:00], 1234.0, 0.12, "Foobar"],
                 [
                   ~N[1993-11-21 00:00:00],
                   "00012345",
                   0.12119999999999999,
                   nil
                 ],
                 [~N[1987-05-08 00:00:00], 1122.0, "12", nil],
                 [~N[1994-05-22 00:00:00], "12,12", 12.0, nil],
                 ["2024-01-01", 11.12, "33.12%", "123"],
                 [~N[1987-05-08 20:10:12], nil, nil, nil],
                 [~N[1987-05-08 20:10:12], nil, nil, nil]
               ]
             }
    end

    test "parses ods files from binary content" do
      content = File.read!(Path.join(@base_path, "test_file_1.ods"))

      # ODS dates arrive as ISO strings but are normalised to NaiveDateTime, so
      # the result matches the xlsx/xls output for the same workbook. The
      # "2024-01-01" cell stays a string because it is a text cell, not a date.
      assert Spreadsheet.parse(content, sheet: "Sheet1", format: :binary) ==
               {
                 :ok,
                 [
                   ["Dates", "Numbers", "Percentages", "Strings"],
                   [~N[2024-12-12 00:00:00], 1234.0, 0.12, "Foobar"],
                   [~N[1993-11-21 00:00:00], "00012345", 0.1212, nil],
                   [~N[1987-05-08 00:00:00], 1122.0, "12", nil],
                   [~N[1994-05-22 00:00:00], "12,12", 12.0, nil],
                   ["2024-01-01", 11.12, "33.12%", "123"],
                   [~N[1987-05-08 20:10:12], nil, nil, nil],
                   [~N[1987-05-08 20:10:12], nil, nil, nil]
                 ]
               }
    end

    test "returns error for invalid format option" do
      path = Path.join(@base_path, "test_file_1.xlsx")

      assert Spreadsheet.parse(path, sheet: "Sheet1", format: :invalid) ==
               {:error,
                "Invalid format option: :invalid. Expected :filename or :binary"}
    end
  end
end
