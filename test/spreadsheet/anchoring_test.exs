defmodule Spreadsheet.AnchoringTest do
  use ExUnit.Case, async: true

  @fixtures Path.expand("../fixtures", __DIR__)

  # Data starts at C3 and ends at E6; rows and columns before it are padded
  # with nil so that rows[r][c] is always row r + 1, column c + 1.
  @expected [
    [nil, nil, nil, nil, nil],
    [nil, nil, nil, nil, nil],
    [nil, nil, "name", "value", nil],
    [nil, nil, "a", 1.0, nil],
    [nil, nil, nil, nil, nil],
    [nil, nil, nil, nil, true]
  ]

  # The same sheet with anchor: :data starts at C3.
  @expected_from_data [
    ["name", "value", nil],
    ["a", 1.0, nil],
    [nil, nil, nil],
    [nil, nil, true]
  ]

  for file <- ~w(offset.xlsx offset.xls offset.ods) do
    test "#{file}: rows and columns are anchored at A1" do
      path = Path.join(@fixtures, unquote(file))

      assert Spreadsheet.parse(path, sheet: "Offset") == {:ok, @expected}

      assert Spreadsheet.parse(path, sheet: "Offset", anchor: :a1) ==
               {:ok, @expected}

      assert Spreadsheet.parse(path) == {:ok, [{"Offset", @expected}]}

      content = File.read!(path)

      assert Spreadsheet.parse(content, sheet: "Offset", format: :binary) ==
               {:ok, @expected}
    end

    test "#{file}: anchor: :data starts at the first non-empty row and column" do
      path = Path.join(@fixtures, unquote(file))

      assert Spreadsheet.parse(path, sheet: "Offset", anchor: :data) ==
               {:ok, @expected_from_data}

      assert Spreadsheet.parse(path, anchor: :data) ==
               {:ok, [{"Offset", @expected_from_data}]}

      content = File.read!(path)

      assert Spreadsheet.parse(content,
               sheet: "Offset",
               format: :binary,
               anchor: :data
             ) ==
               {:ok, @expected_from_data}
    end
  end

  test ".xlsb rows and columns are anchored at A1" do
    path = Path.join(@fixtures, "types.xlsb")

    assert Spreadsheet.parse(path, sheet: "Offset") ==
             {:ok, [[nil, nil, nil], [nil, nil, nil], [nil, nil, "C3"]]}
  end

  test ".xlsb anchor: :data starts at the first non-empty cell" do
    path = Path.join(@fixtures, "types.xlsb")

    assert Spreadsheet.parse(path, sheet: "Offset", anchor: :data) ==
             {:ok, [["C3"]]}
  end

  test "anchor: :data leaves a sheet that starts at A1 unchanged" do
    path = Path.join(@fixtures, "test_file_1.xlsx")

    assert {:ok, [[first_cell | _] | _] = rows} =
             Spreadsheet.parse(path, sheet: "Sheet1")

    refute is_nil(first_cell)

    assert Spreadsheet.parse(path, sheet: "Sheet1", anchor: :data) ==
             {:ok, rows}
  end

  test "anchor: :data spends only the data's rectangle from :max_cells" do
    path = Path.join(@fixtures, "offset.xlsx")

    # C3:E6 is 4 × 3 cells; from A1 it is 6 × 5.
    assert Spreadsheet.parse(path,
             sheet: "Offset",
             anchor: :data,
             max_cells: 12
           ) ==
             {:ok, @expected_from_data}

    assert {:error, "sheet \"Offset\" has 6 rows × 5 columns" <> _} =
             Spreadsheet.parse(path, sheet: "Offset", max_cells: 12)
  end
end
