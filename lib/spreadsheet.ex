defmodule Spreadsheet do
  @moduledoc """
  A fast, memory-efficient Elixir library for parsing spreadsheet files.

  This library provides a simple API for working with Excel (.xlsx, .xls) and
  LibreOffice (.ods) files. It's powered by Rust and Calamine for high-performance
  parsing.

  ## Features

  - Fast performance with native Rust implementation
  - Support for multiple formats: .xls, .xla, .xlsx, .xlsm, .xlam, .xlsb, and .ods
  - Memory efficient parsing from file paths or binary content
  - Sheet management with support for hidden sheets
  - Smart type handling with automatic date and number conversion
  """

  alias Spreadsheet.Calamine

  @doc """
  Returns a list of sheet names from a spreadsheet file or binary content.

  Supports Excel (.xlsx, .xls) and LibreOffice (.ods) file formats.

  ## Options

    * `:format` - Specifies the input format. Either `:filename` (default) or `:binary`.
    * `:hidden` - When `false`, excludes hidden sheets. Defaults to `true`.

  ## Examples

      # From a file path (default)
      Spreadsheet.sheet_names("workbook.xlsx")
      {:ok, ["Sheet1", "Sheet2"]}

      # From a file path (explicit)
      Spreadsheet.sheet_names("workbook.xlsx", format: :filename)
      {:ok, ["Sheet1", "Sheet2"]}

      # From binary content
      content = File.read!("workbook.xlsx")
      Spreadsheet.sheet_names(content, format: :binary)
      {:ok, ["Sheet1", "Sheet2"]}

      # Exclude hidden sheets
      Spreadsheet.sheet_names("workbook.xlsx", hidden: false)
      {:ok, ["Sheet1"]}

  """
  @spec sheet_names(binary(), keyword()) ::
          {:ok, list(String.t())} | {:error, String.t()}
  def sheet_names(path_or_content, opts \\ [])
      when is_binary(path_or_content) and is_list(opts) do
    format = Keyword.get(opts, :format, :filename)
    include_hidden = Keyword.get(opts, :hidden, true)

    case format do
      :filename ->
        Calamine.sheet_names_from_path(path_or_content, include_hidden)

      :binary ->
        Calamine.sheet_names_from_binary(path_or_content, include_hidden)

      other ->
        {:error,
         "Invalid format option: #{inspect(other)}. Expected :filename or :binary"}
    end
  end

  @doc """
  Parses a specific sheet or all sheets from a spreadsheet file or binary content.

  When called without the `:sheet` option, parses all sheets and returns
  a list of tuples containing `{sheet_name, sheet_data}`.

  When called with the `:sheet` option, parses only the specified sheet and
  returns its data as a list of lists.

  Returns the sheet data as a list of lists, where each inner list represents a row.
  The first row typically contains headers.

  Dates are automatically parsed to `NaiveDateTime` when possible, and empty cells
  are converted to `nil`.

  ## Options

    * `:sheet` - The name of the sheet to parse. If not provided, parses all sheets.
    * `:format` - Specifies the input format. Either `:filename` (default) or `:binary`.
    * `:hidden` - When `false`, excludes hidden sheets from all-sheets parsing. Defaults to `true`.
      Ignored when `:sheet` is provided — naming a sheet explicitly always parses it.

  Excel formula errors (`#REF!`, `#DIV/0!`, etc.) are returned as `{:error, reason}`
  tuples so they remain distinguishable from text cells.

  ## Examples

      # Parse a specific sheet from a file path
      Spreadsheet.parse("sales.xlsx", sheet: "Q1 Data")
      {:ok, [
        ["Product", "Sales", "Date"],
        ["Widget A", 1500.0, ~N[2024-01-15 00:00:00]]
      ]}

      # Parse all sheets from a file path
      Spreadsheet.parse("sales.xlsx")
      {:ok, [
        {"Q1 Data", [
          ["Product", "Sales", "Date"],
          ["Widget A", 1500.0, ~N[2024-01-15 00:00:00]]
        ]},
        {"Q2 Data", [
          ["Product", "Sales", "Date"],
          ["Widget B", 2300.0, ~N[2024-04-15 00:00:00]]
        ]}
      ]}

      # Parse all sheets from binary content
      content = File.read!("sales.xlsx")
      Spreadsheet.parse(content, format: :binary)
      {:ok, [
        {"Q1 Data", [...]},
        {"Q2 Data", [...]}
      ]}

      # Parse all sheets excluding hidden ones
      Spreadsheet.parse("sales.xlsx", hidden: false)
      {:ok, [{"Visible Sheet", [...]}]}

      # Parse specific sheet from binary
      Spreadsheet.parse(content, sheet: "Q1 Data", format: :binary)
      {:ok, [[...]]}

  """
  @spec parse(binary(), keyword()) ::
          {:ok, list() | list({String.t(), list()})} | {:error, binary()}
  def parse(path_or_content, opts \\ [])
      when is_binary(path_or_content) and is_list(opts) do
    format = Keyword.get(opts, :format, :filename)

    case Keyword.get(opts, :sheet) do
      nil ->
        parse_all_sheets(path_or_content, format, Keyword.get(opts, :hidden, true))

      sheet_name ->
        parse_single_sheet(path_or_content, sheet_name, format)
    end
  end

  defp parse_single_sheet(path_or_content, sheet_name, format) do
    with {:ok, rows} <- call_parser(path_or_content, sheet_name, format) do
      {:ok, Spreadsheet.Parser.parse_rows(rows)}
    end
  end

  defp call_parser(path, sheet_name, :filename),
    do: Calamine.parse_from_path(path, sheet_name)

  defp call_parser(content, sheet_name, :binary),
    do: Calamine.parse_from_binary(content, sheet_name)

  defp call_parser(_, _, other),
    do:
      {:error,
       "Invalid format option: #{inspect(other)}. Expected :filename or :binary"}

  defp parse_all_sheets(path_or_content, :filename, include_hidden) do
    with {:ok, sheets} <-
           Calamine.parse_all_from_path(path_or_content, include_hidden) do
      {:ok, Enum.map(sheets, fn {name, rows} -> {name, Spreadsheet.Parser.parse_rows(rows)} end)}
    end
  end

  defp parse_all_sheets(path_or_content, :binary, include_hidden) do
    with {:ok, sheets} <-
           Calamine.parse_all_from_binary(path_or_content, include_hidden) do
      {:ok, Enum.map(sheets, fn {name, rows} -> {name, Spreadsheet.Parser.parse_rows(rows)} end)}
    end
  end

  defp parse_all_sheets(_, other, _),
    do:
      {:error,
       "Invalid format option: #{inspect(other)}. Expected :filename or :binary"}

  @doc """
  Returns a list of sheet names from spreadsheet binary content.

  This function is deprecated. Use `sheet_names/2` with `format: :binary` instead.

  ## Options

    * `:hidden` - When `false`, excludes hidden sheets. Defaults to `true`.

  """
  @deprecated "Use sheet_names/2 with format: :binary instead"
  @spec sheet_names_from_binary(binary(), keyword()) ::
          {:ok, list(String.t())} | {:error, String.t()}
  def sheet_names_from_binary(content, opts \\ [])
      when is_binary(content) and is_list(opts) do
    sheet_names(content, Keyword.put(opts, :format, :binary))
  end

  @doc """
  Parses a specific sheet from spreadsheet binary content.

  This function is deprecated. Use `parse/2` with `sheet:` and `format: :binary` options instead.

  Returns the sheet data as a list of lists, where each inner list represents a row.

  Dates are automatically parsed to `NaiveDateTime` when possible, and empty cells
  are converted to `nil`.

  """
  @deprecated "Use parse/2 with sheet: and format: :binary options instead"
  @spec parse_from_binary(binary(), binary()) ::
          {:ok, list()} | {:error, String.t()}
  def parse_from_binary(content, sheet_name) do
    parse(content, sheet: sheet_name, format: :binary)
  end

end
