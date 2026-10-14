defmodule Spreadsheet do
  @external_resource readme = Path.join([__DIR__, "..", "README.md"])

  @moduledoc readme
             |> File.read!()
             |> String.split("<!-- MDOC !-->")
             |> Enum.fetch!(1)

  alias Spreadsheet.Calamine

  @default_max_cells 10_000_000

  # The NIF takes an unsigned 64-bit limit, or nil for no limit.
  @max_u64 18_446_744_073_709_551_615

  @typedoc """
  A cell value: `nil`, a float, a string, a boolean, a `NaiveDateTime`, a
  `Time`, a `Duration`, or `{:error, code}` for a formula error. See
  "Data Types" in the module documentation.
  """
  @type cell ::
          nil
          | float()
          | String.t()
          | boolean()
          | NaiveDateTime.t()
          | Time.t()
          | Duration.t()
          | {:error, String.t()}

  @typedoc "A sheet's rows, anchored at A1 and padded with `nil`."
  @type rows :: [[cell()]]

  @doc """
  Returns the names of the sheets in a spreadsheet, in workbook order.

  ## Options

    * `:format` - `:filename` (default) when the first argument is a path, or
      `:binary` when it is the file's content.
    * `:hidden` - when `false`, excludes hidden and very hidden sheets.
      Defaults to `true`.
    * `:max_cells` - the largest sheet accepted, as rows × columns counted from
      A1, or `:infinity`. Defaults to `#{@default_max_cells}`. Listing sheets
      only reads cells for .xls and .ods files, which are decoded in full when
      opened, so only those can exceed it here.

  Returns `{:error, reason}` when the file cannot be read or parsed, and
  raises `ArgumentError` for unknown or invalid options.

  ## Examples

      # From a file path (default)
      Spreadsheet.sheet_names("workbook.xlsx")
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
          {:ok, [String.t()]} | {:error, String.t()}
  def sheet_names(path_or_content, opts \\ [])
      when is_binary(path_or_content) and is_list(opts) do
    opts =
      Keyword.validate!(opts,
        format: :filename,
        hidden: true,
        max_cells: @default_max_cells
      )

    include_hidden = boolean_option!(opts, :hidden)
    max_cells = max_cells_option!(opts)

    with {:ok, source} <- source(path_or_content, opts[:format]) do
      Calamine.sheet_names(source, include_hidden, max_cells)
    end
  end

  @doc """
  Parses one sheet, or every sheet, of a spreadsheet.

  With the `:sheet` option, returns that sheet's rows. Without it (or with
  `sheet: nil`), returns `{sheet_name, rows}` tuples for every sheet in
  workbook order.

  Rows are lists of cells anchored at A1: row 0 is spreadsheet row 1 and
  column 0 is column A, with leading empty rows and columns padded with `nil`.
  See "Data Types" in the module documentation for the value of each cell.

  ## Options

    * `:sheet` - the name of the sheet to parse. Defaults to `nil` (all
      sheets).
    * `:format` - `:filename` (default) when the first argument is a path, or
      `:binary` when it is the file's content.
    * `:hidden` - when `false`, excludes hidden and very hidden sheets from
      all-sheets parsing. Defaults to `true`. Ignored when `:sheet` is given:
      naming a sheet always parses it.
    * `:max_cells` - the largest sheet accepted, as rows × columns counted from
      A1, or `:infinity`. Defaults to `#{@default_max_cells}`. Larger sheets
      return an error before they are expanded.

  Returns `{:error, reason}` when the file cannot be read or parsed, the sheet
  does not exist or a sheet exceeds `:max_cells`, and raises `ArgumentError`
  for unknown or invalid options.

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

      # Parse all visible sheets from binary content
      content = File.read!("sales.xlsx")
      Spreadsheet.parse(content, format: :binary, hidden: false)
      {:ok, [{"Q1 Data", [...]}, {"Q2 Data", [...]}]}

      # Allow a larger sheet
      Spreadsheet.parse("big.xlsx", sheet: "Data", max_cells: 50_000_000)
      {:ok, [[...]]}

  """
  @spec parse(binary(), keyword()) ::
          {:ok, rows() | [{String.t(), rows()}]} | {:error, String.t()}
  def parse(path_or_content, opts \\ [])
      when is_binary(path_or_content) and is_list(opts) do
    opts =
      Keyword.validate!(opts,
        format: :filename,
        hidden: true,
        sheet: nil,
        max_cells: @default_max_cells
      )

    sheet = sheet_option!(opts)
    include_hidden = boolean_option!(opts, :hidden)
    max_cells = max_cells_option!(opts)

    with {:ok, source} <- source(path_or_content, opts[:format]) do
      Calamine.parse(source, sheet, include_hidden, max_cells)
    end
  end

  @doc """
  Returns a list of sheet names from spreadsheet binary content.

  This function is deprecated. Use `sheet_names/2` with `format: :binary` instead.

  ## Options

    * `:hidden` - When `false`, excludes hidden sheets. Defaults to `true`.
    * `:max_cells` - See `sheet_names/2`.

  """
  @deprecated "Use sheet_names/2 with format: :binary instead"
  @spec sheet_names_from_binary(binary(), keyword()) ::
          {:ok, [String.t()]} | {:error, String.t()}
  def sheet_names_from_binary(content, opts \\ [])
      when is_binary(content) and is_list(opts) do
    sheet_names(content, Keyword.put(opts, :format, :binary))
  end

  @doc """
  Parses a specific sheet from spreadsheet binary content.

  This function is deprecated. Use `parse/2` with `sheet:` and `format: :binary` options instead.

  Returns the sheet's rows, as `parse/2` does.
  """
  @deprecated "Use parse/2 with sheet: and format: :binary options instead"
  @spec parse_from_binary(binary(), binary()) ::
          {:ok, rows()} | {:error, String.t()}
  def parse_from_binary(content, sheet_name) do
    parse(content, sheet: sheet_name, format: :binary)
  end

  # An invalid `:format` is reported as an error tuple rather than raised, as
  # it always has been.
  defp source(path, :filename), do: {:ok, {:path, path}}
  defp source(content, :binary), do: {:ok, {:binary, content}}

  defp source(_path_or_content, other),
    do:
      {:error,
       "Invalid format option: #{inspect(other)}. Expected :filename or :binary"}

  defp sheet_option!(opts) do
    case opts[:sheet] do
      sheet when is_binary(sheet) or is_nil(sheet) ->
        sheet

      other ->
        raise ArgumentError,
              "expected :sheet to be a string (or nil for all sheets), got: #{inspect(other)}"
    end
  end

  defp max_cells_option!(opts) do
    case opts[:max_cells] do
      :infinity ->
        nil

      max_cells when is_integer(max_cells) and max_cells > 0 ->
        min(max_cells, @max_u64)

      other ->
        raise ArgumentError,
              "expected :max_cells to be a positive integer or :infinity, got: #{inspect(other)}"
    end
  end

  defp boolean_option!(opts, key) do
    case opts[key] do
      value when is_boolean(value) ->
        value

      other ->
        raise ArgumentError,
              "expected #{inspect(key)} to be a boolean, got: #{inspect(other)}"
    end
  end
end
