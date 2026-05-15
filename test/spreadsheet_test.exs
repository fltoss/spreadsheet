defmodule SpreadsheetTest do
  use ExUnit.Case
  doctest Spreadsheet

  @base_path Path.join(__DIR__, "/files")

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

  describe "Calamine.parse_all_from_path/2 (bulk NIF — bug #4)" do
    # Bug #4: parse_all_sheets currently re-opens the workbook for each sheet.
    # A bulk NIF lets us open once and return all sheets.
    test "returns [{sheet_name, rows}] in workbook order, honouring hidden" do
      path = Path.join(@base_path, "test_file_1_hidden.xlsx")

      assert {:ok, all_sheets} =
               Spreadsheet.Calamine.parse_all_from_path(path, true)

      assert [{"Sheet2", _}, {"Sheet1", _}] = all_sheets

      assert {:ok, [{"Sheet1", _}]} =
               Spreadsheet.Calamine.parse_all_from_path(path, false)
    end

    test "binary variant returns the same shape" do
      content = File.read!(Path.join(@base_path, "test_file_1.xlsx"))

      assert {:ok, [{"Sheet1", [header | _]}]} =
               Spreadsheet.Calamine.parse_all_from_binary(content, true)

      # Raw NIF output uses Rustler's tagged form; parse_rows would unwrap it.
      assert header == [
               string: "Dates",
               string: "Numbers",
               string: "Percentages",
               string: "Strings"
             ]
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
               [~N[2024-12-12 00:00:00], 1234, 0.12, "Foobar"],
               [
                 ~N[1993-11-21 00:00:00],
                 "00012345",
                 0.12119999999999999,
                 nil
               ],
               [~N[1987-05-08 00:00:00], 1122, "12", nil],
               [~N[1994-05-22 00:00:00], "12,12", 12, nil],
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
                 [~N[2024-12-12 00:00:00], 1234, 0.12, "Foobar"],
                 [
                   ~N[1993-11-21 00:00:00],
                   "00012345",
                   0.12119999999999999,
                   nil
                 ],
                 [~N[1987-05-08 00:00:00], 1122, "12", nil],
                 [~N[1994-05-22 00:00:00], "12,12", 12, nil],
                 ["2024-01-01", 11.12, "33.12%", "123"],
                 [~N[1987-05-08 20:10:12], nil, nil, nil],
                 [~N[1987-05-08 20:10:12], nil, nil, nil]
               ]
             }
    end

    test "parses ods files from binary content" do
      content = File.read!(Path.join(@base_path, "test_file_1.ods"))

      assert Spreadsheet.parse(content, sheet: "Sheet1", format: :binary) ==
               {
                 :ok,
                 [
                   ["Dates", "Numbers", "Percentages", "Strings"],
                   ["2024-12-12", 1234.0, 0.12, "Foobar"],
                   ["1993-11-21", "00012345", 0.1212, nil],
                   ["1987-05-08", 1122.0, "12", nil],
                   ["1994-05-22", "12,12", 12.0, nil],
                   ["2024-01-01", 11.12, "33.12%", "123"],
                   ["1987-05-08T20:10:12", nil, nil, nil],
                   ["1987-05-08T20:10:12", nil, nil, nil]
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
