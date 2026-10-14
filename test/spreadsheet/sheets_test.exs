defmodule Spreadsheet.SheetsTest do
  use ExUnit.Case, async: true

  @fixtures Path.expand("../fixtures", __DIR__)

  defp fixture(name), do: Path.join(@fixtures, name)

  # sheets.xls and sheets.ods are LibreOffice exports of sheets.xlsx.
  for file <- ~w(sheets.xlsx sheets.xls sheets.ods) do
    describe file do
      test "lists hidden, very hidden and non-ASCII sheets in order" do
        assert Spreadsheet.sheet_names(fixture(unquote(file))) ==
                 {:ok, ["Visible", "Hidden", "Ünïcødé 名前", "Very Hidden"]}
      end

      test "hidden: false drops hidden and very hidden sheets" do
        path = fixture(unquote(file))

        assert Spreadsheet.sheet_names(path, hidden: false) ==
                 {:ok, ["Visible", "Ünïcødé 名前"]}

        assert Spreadsheet.parse(path, hidden: false) ==
                 {:ok,
                  [{"Visible", [["visible"]]}, {"Ünïcødé 名前", [["unicode"]]}]}
      end

      test "a sheet named explicitly is parsed even when hidden" do
        path = fixture(unquote(file))

        assert Spreadsheet.parse(path, sheet: "Very Hidden", hidden: false) ==
                 {:ok, [["very hidden"]]}

        assert Spreadsheet.parse(path, sheet: "Ünïcødé 名前") ==
                 {:ok, [["unicode"]]}
      end
    end
  end

  test ".xlsb hidden sheets" do
    path = fixture("types.xlsb")

    assert Spreadsheet.sheet_names(path) == {:ok, ["Types", "Hidden", "Offset"]}

    assert Spreadsheet.sheet_names(path, hidden: false) ==
             {:ok, ["Types", "Offset"]}
  end

  test "chart sheets are listed and have no rows" do
    path = fixture("chartsheet.xlsx")

    assert Spreadsheet.sheet_names(path) == {:ok, ["Data", "Chart"]}
    assert Spreadsheet.parse(path, sheet: "Chart") == {:ok, []}
    assert Spreadsheet.parse(path) == {:ok, [{"Data", [[1.0]]}, {"Chart", []}]}
  end
end
