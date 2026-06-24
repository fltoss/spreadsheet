defmodule Spreadsheet.Parser do
  @moduledoc false

  def parse_rows(rows) do
    for row <- rows, do: for(col <- row, do: parse_col(col))
  end

  def parse_col(:empty), do: nil

  def parse_col({:date_time, val}) do
    case NaiveDateTime.from_iso8601(val) do
      {:ok, dt} -> dt
      _ -> nil
    end
  end

  # ODS files surface dates as ISO strings (date-only or full datetime) rather
  # than the serial :date_time form xlsx/xls use. Parse them to a NaiveDateTime
  # so callers get one consistent type regardless of source format, falling back
  # to the raw string only when the value isn't a recognisable date/datetime.
  def parse_col({:date_time_iso, val}) do
    case NaiveDateTime.from_iso8601(val) do
      {:ok, dt} ->
        dt

      {:error, _} ->
        case Date.from_iso8601(val) do
          {:ok, date} -> NaiveDateTime.new!(date, ~T[00:00:00])
          {:error, _} -> val
        end
    end
  end

  def parse_col({:error, val}), do: {:error, val}

  def parse_col({_, val}), do: val
end
