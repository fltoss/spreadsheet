defmodule Spreadsheet.ParserTest do
  use ExUnit.Case, async: true

  alias Spreadsheet.Parser

  describe "parse_col/1 — passthrough variants" do
    test ":empty maps to nil" do
      assert Parser.parse_col(:empty) == nil
    end

    test "scalar variants unwrap to their value" do
      assert Parser.parse_col({:int, 42}) == 42
      assert Parser.parse_col({:float, 1.5}) == 1.5
      assert Parser.parse_col({:string, "hi"}) == "hi"
      assert Parser.parse_col({:bool, true}) == true
      assert Parser.parse_col({:date_time_iso, "2024-01-15"}) == "2024-01-15"
      assert Parser.parse_col({:duration_iso, "PT1H"}) == "PT1H"
    end
  end

  describe "parse_col/1 — datetime" do
    test "parses a valid ISO datetime" do
      assert Parser.parse_col({:date_time, "2024-12-12T00:00:00"}) ==
               ~N[2024-12-12 00:00:00]
    end

    # Bug #7: sub-second precision is dropped.
    test "preserves sub-second precision when present" do
      assert Parser.parse_col({:date_time, "2024-12-12T10:30:45.123456"}) ==
               ~N[2024-12-12 10:30:45.123456]
    end

    # Bug #2: unparseable datetime currently leaks back as a raw string.
    test "returns nil when the datetime string is unparseable" do
      assert Parser.parse_col({:date_time, "Invalid DateTime"}) == nil
      assert Parser.parse_col({:date_time, "not-a-date"}) == nil
    end
  end

  describe "parse_col/1 — error cells" do
    # Bug #3: Excel error cells currently collapse to plain strings,
    # indistinguishable from real text cells that happen to read "#DIV/0!".
    test "tags Excel formula errors as {:error, reason}" do
      assert Parser.parse_col({:error, "#DIV/0!"}) == {:error, "#DIV/0!"}
      assert Parser.parse_col({:error, "#REF!"}) == {:error, "#REF!"}
      assert Parser.parse_col({:error, "#N/A"}) == {:error, "#N/A"}
    end
  end

  describe "parse_rows/1" do
    test "applies parse_col to every cell" do
      rows = [
        [{:int, 1}, {:string, "a"}, :empty],
        [{:date_time, "2024-01-15T00:00:00"}, {:error, "#REF!"}, {:bool, false}]
      ]

      assert Parser.parse_rows(rows) == [
               [1, "a", nil],
               [~N[2024-01-15 00:00:00], {:error, "#REF!"}, false]
             ]
    end
  end
end
