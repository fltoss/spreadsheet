defmodule Spreadsheet.CellTypesTest do
  use ExUnit.Case, async: true

  @fixtures Path.expand("../fixtures", __DIR__)

  # The "Types" sheet holds one label/value pair per row.
  defp values(file) do
    {:ok, [["label", "value"] | rows]} =
      Spreadsheet.parse(Path.join(@fixtures, file), sheet: "Types")

    Map.new(rows, fn [label, value] -> {label, value} end)
  end

  describe "Excel formats (.xlsx, .xls, .xlsb)" do
    for file <- ~w(types.xlsx types.xls) do
      test "#{file}: numbers are floats" do
        values = values(unquote(file))

        assert values["integer"] === 1234.0
        assert values["float"] === 0.5
        assert values["large_integer"] === 12_345_678_901_234_567.0
      end

      test "#{file}: dates and date-times are NaiveDateTimes" do
        values = values(unquote(file))

        assert values["date"] == ~N[2025-10-13 00:00:00]
        assert values["date_time"] == ~N[2025-10-13 12:00:00]
        assert values["date_time_ms"] == ~N[2025-10-13 12:00:00.123]

        assert NaiveDateTime.to_iso8601(values["date_time_ms"]) ==
                 "2025-10-13T12:00:00.123"

        assert NaiveDateTime.add(values["date"], 1, :day) ==
                 ~N[2025-10-14 00:00:00]
      end

      test "#{file}: time-of-day cells are Times" do
        values = values(unquote(file))

        assert values["time"] == ~T[12:00:00]
        assert values["time_ms"] == ~T[18:00:00.250]
        assert Time.to_string(values["time_ms"]) == "18:00:00.250"
        assert Time.diff(values["time_ms"], values["time"]) == 6 * 3600
      end

      test "#{file}: durations are Durations and keep their sign" do
        values = values(unquote(file))

        assert values["duration"] == Duration.new!(hour: 25, minute: 30)

        assert values["negative_duration"] ==
                 Duration.new!(hour: -1, minute: -30)

        assert values["duration_ms"] ==
                 Duration.new!(second: 1, microsecond: {500_000, 3})

        assert Duration.to_iso8601(values["duration"]) == "PT25H30M"
        assert Duration.to_iso8601(values["negative_duration"]) == "PT-1H-30M"
        assert Duration.to_iso8601(values["duration_ms"]) == "PT1.500S"
      end

      test "#{file}: unrepresentable dates return the serial number" do
        values = values(unquote(file))

        assert values["negative_date"] === -1.0
        assert values["out_of_range_date"] === 3_000_000.0
      end

      test "#{file}: error cells are error tuples" do
        values = values(unquote(file))

        assert values["div_by_zero"] == {:error, "#DIV/0!"}
        assert values["not_available"] == {:error, "#N/A"}
        assert values["formula_error"] == {:error, "#DIV/0!"}
      end

      test "#{file}: booleans, strings and empty cells" do
        values = values(unquote(file))

        assert values["boolean"] === true
        assert values["string"] == "Zoë"
        assert values["empty"] == nil
      end
    end

    test "types.xlsb" do
      assert values("types.xlsb") == %{
               "integer" => 1234.0,
               "float" => 0.5,
               "date_time" => ~N[2025-10-13 12:00:00],
               "time" => ~T[12:00:00],
               "negative_duration" => Duration.new!(hour: -1, minute: -30),
               "div_by_zero" => {:error, "#DIV/0!"},
               "boolean" => true,
               "string" => "Zoë"
             }

      assert values("types.xlsb")["integer"] === 1234.0
    end

    test "1904-epoch workbooks" do
      assert Spreadsheet.parse(Path.join(@fixtures, "epoch_1904.xlsx"),
               sheet: "Dates"
             ) ==
               {:ok,
                [
                  [
                    ~N[2024-01-01 00:00:00],
                    ~N[2024-01-01 18:00:00],
                    ~T[12:00:00]
                  ]
                ]}
    end
  end

  describe ".ods" do
    test "LibreOffice export of the types workbook" do
      values = values("types.ods")

      assert values["integer"] === 1234.0
      assert values["large_integer"] === 12_345_678_901_234_600.0
      assert values["date"] == ~N[2025-10-13 00:00:00]
      assert values["date_time_ms"] == ~N[2025-10-13 12:00:00.123]
      # ODS cannot tell a time of day from a duration: both are durations.
      assert values["time"] == Duration.new!(hour: 12)
      assert values["duration"] == Duration.new!(hour: 25, minute: 30)
      assert values["negative_duration"] == Duration.new!(hour: -1, minute: -30)
      assert values["boolean"] === true
      assert values["string"] == "Zoë"
      assert values["empty"] == nil
      # LibreOffice writes dates beyond year 9999 with five-digit years, which
      # are returned as the raw string.
      assert values["out_of_range_date"] == "10113-09-19"
    end

    test "ISO 8601 dates and durations" do
      assert {:ok, [dates, durations]} =
               Spreadsheet.parse(Path.join(@fixtures, "iso.ods"), sheet: "Iso")

      assert dates == [
               ~N[2024-01-15 00:00:00],
               ~N[1987-05-08 20:10:12.5],
               ~N[2024-01-15 10:00:00.123456],
               "not-a-date"
             ]

      assert Enum.map(Enum.take(dates, 3), &NaiveDateTime.to_iso8601/1) == [
               "2024-01-15T00:00:00",
               "1987-05-08T20:10:12.5",
               "2024-01-15T10:00:00.123456"
             ]

      assert durations == [
               Duration.new!(hour: 36),
               Duration.new!(hour: -1, minute: -30),
               Duration.new!(second: 1, microsecond: {500_000, 1}),
               "garbage"
             ]

      assert Enum.map(Enum.take(durations, 3), &Duration.to_iso8601/1) == [
               "PT36H",
               "PT-1H-30M",
               "PT1.5S"
             ]
    end
  end
end
