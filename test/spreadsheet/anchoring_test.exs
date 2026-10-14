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

  for file <- ~w(offset.xlsx offset.xls offset.ods) do
    test "#{file}: rows and columns are anchored at A1" do
      path = Path.join(@fixtures, unquote(file))

      assert Spreadsheet.parse(path, sheet: "Offset") == {:ok, @expected}
      assert Spreadsheet.parse(path) == {:ok, [{"Offset", @expected}]}

      content = File.read!(path)

      assert Spreadsheet.parse(content, sheet: "Offset", format: :binary) ==
               {:ok, @expected}
    end
  end

  test ".xlsb rows and columns are anchored at A1" do
    path = Path.join(@fixtures, "types.xlsb")

    assert Spreadsheet.parse(path, sheet: "Offset") ==
             {:ok, [[nil, nil, nil], [nil, nil, nil], [nil, nil, "C3"]]}
  end
end
