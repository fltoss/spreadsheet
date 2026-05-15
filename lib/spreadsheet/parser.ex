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

  def parse_col({:error, val}), do: {:error, val}

  def parse_col({_, val}), do: val
end
