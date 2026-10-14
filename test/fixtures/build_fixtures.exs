# Regenerates the hand-built fixtures in this directory:
#
#     elixir test/fixtures/build_fixtures.exs
#
# The .xlsx, .xlsb and .ods files are assembled from XML/BIFF12 records written
# here. The .xls files are converted from them with LibreOffice (`soffice` or
# `libreoffice` must be on the PATH) and some are then patched byte-by-byte to
# craft malformed input. test_file_1*.{xls,xlsx,ods} predate this script and
# are not generated.
#
# LibreOffice installed as a snap cannot read /tmp, so conversion happens in a
# scratch directory next to this script.

defmodule Fixtures.Zip do
  def write!(path, entries) do
    entries =
      Enum.map(entries, fn {name, content} ->
        {String.to_charlist(name), content}
      end)

    {:ok, {_name, binary}} = :zip.create(~c"fixture.zip", entries, [:memory])
    File.write!(path, binary)
  end
end

defmodule Fixtures.Xml do
  def escape(text) do
    text
    |> String.replace("&", "&amp;")
    |> String.replace("<", "&lt;")
    |> String.replace(">", "&gt;")
    |> String.replace("\"", "&quot;")
  end
end

defmodule Fixtures.Xlsx do
  @moduledoc false
  import Fixtures.Xml

  @main "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  @relationships "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  @package_relationships "http://schemas.openxmlformats.org/package/2006/relationships"

  # Cell styles (index into cellXfs):
  #   0 General, 1 date (14), 2 date-time (22), 3 date-time with milliseconds,
  #   4 time (21), 5 duration [h]:mm:ss (46), 6 duration with milliseconds.
  @styles """
  <?xml version="1.0" encoding="UTF-8"?>
  <styleSheet xmlns="#{@main}"><numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss.000"/><numFmt numFmtId="165" formatCode="[h]:mm:ss.000"/></numFmts><cellXfs count="7"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/><xf numFmtId="22" applyNumberFormat="1"/><xf numFmtId="164" applyNumberFormat="1"/><xf numFmtId="21" applyNumberFormat="1"/><xf numFmtId="46" applyNumberFormat="1"/><xf numFmtId="165" applyNumberFormat="1"/></cellXfs></styleSheet>
  """

  @doc """
  `sheets` is a list of `{name, options}` where options has `:rows` (a list of
  `{row_number, [cell]}`), `:state` (`"hidden"`/`"veryHidden"`) and `:chart`
  (`true` for a chartsheet). A cell is `{reference, kind, value}` or
  `{reference, kind, value, style}`.
  """
  def write!(path, sheets, opts \\ []) do
    shared_strings = Keyword.get(opts, :shared_strings, [])

    {sheet_entries, workbook_sheets, relationships, overrides} =
      sheets
      |> Enum.with_index(1)
      |> Enum.map(fn {{name, sheet_opts}, index} ->
        chart = Keyword.get(sheet_opts, :chart, false)
        folder = if chart, do: "chartsheets", else: "worksheets"
        kind = if chart, do: "chartsheet", else: "worksheet"
        state = Keyword.get(sheet_opts, :state)
        state_attribute = if state, do: ~s( state="#{state}"), else: ""

        content =
          if chart,
            do: chartsheet(),
            else: worksheet(Keyword.fetch!(sheet_opts, :rows), sheet_opts)

        {{"xl/#{folder}/sheet#{index}.xml", content},
         ~s(<sheet name="#{escape(name)}" sheetId="#{index}"#{state_attribute} r:id="rId#{index}"/>),
         ~s(<Relationship Id="rId#{index}" Type="#{@relationships}/#{kind}" Target="#{folder}/sheet#{index}.xml"/>),
         ~s(<Override PartName="/xl/#{folder}/sheet#{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.#{kind}+xml"/>)}
      end)
      |> unzip4()

    workbook_properties =
      if Keyword.get(opts, :date1904, false),
        do: ~s(<workbookPr date1904="1"/>),
        else: ""

    workbook = """
    <?xml version="1.0" encoding="UTF-8"?>
    <workbook xmlns="#{@main}" xmlns:r="#{@relationships}">#{workbook_properties}<sheets>#{Enum.join(workbook_sheets)}</sheets></workbook>
    """

    shared_strings_xml = """
    <?xml version="1.0" encoding="UTF-8"?>
    <sst xmlns="#{@main}" count="#{length(shared_strings)}" uniqueCount="#{length(shared_strings)}">#{Enum.map_join(shared_strings, &"<si><t>#{escape(&1)}</t></si>")}</sst>
    """

    count = length(sheets)

    workbook_relationships = """
    <?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="#{@package_relationships}">#{Enum.join(relationships)}<Relationship Id="rId#{count + 1}" Type="#{@relationships}/styles" Target="styles.xml"/><Relationship Id="rId#{count + 2}" Type="#{@relationships}/sharedStrings" Target="sharedStrings.xml"/></Relationships>
    """

    content_types = """
    <?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>#{Enum.join(overrides)}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>
    """

    root_relationships = """
    <?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="#{@package_relationships}"><Relationship Id="rId1" Type="#{@relationships}/officeDocument" Target="xl/workbook.xml"/></Relationships>
    """

    Fixtures.Zip.write!(
      path,
      [
        {"[Content_Types].xml", content_types},
        {"_rels/.rels", root_relationships},
        {"xl/workbook.xml", workbook},
        {"xl/_rels/workbook.xml.rels", workbook_relationships},
        {"xl/styles.xml", @styles},
        {"xl/sharedStrings.xml", shared_strings_xml}
      ] ++ sheet_entries
    )
  end

  defp unzip4(tuples) do
    {Enum.map(tuples, &elem(&1, 0)), Enum.map(tuples, &elem(&1, 1)),
     Enum.map(tuples, &elem(&1, 2)), Enum.map(tuples, &elem(&1, 3))}
  end

  defp worksheet(rows, sheet_opts) do
    dimension =
      case Keyword.get(sheet_opts, :dimension) do
        nil -> ""
        reference -> ~s(<dimension ref="#{reference}"/>)
      end

    rows_xml =
      Enum.map_join(rows, fn {row_number, cells} ->
        ~s(<row r="#{row_number}">#{Enum.map_join(cells, &cell/1)}</row>)
      end)

    """
    <?xml version="1.0" encoding="UTF-8"?>
    <worksheet xmlns="#{@main}">#{dimension}<sheetData>#{rows_xml}</sheetData></worksheet>
    """
  end

  defp chartsheet do
    """
    <?xml version="1.0" encoding="UTF-8"?>
    <chartsheet xmlns="#{@main}"><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>
    """
  end

  defp cell({reference, kind, value}), do: cell({reference, kind, value, 0})

  defp cell({reference, :number, value, style}),
    do: ~s(<c r="#{reference}" s="#{style}"><v>#{value}</v></c>)

  defp cell({reference, :shared_string, index, _style}),
    do: ~s(<c r="#{reference}" t="s"><v>#{index}</v></c>)

  defp cell({reference, :string, value, _style}),
    do:
      ~s(<c r="#{reference}" t="inlineStr"><is><t>#{escape(value)}</t></is></c>)

  defp cell({reference, :boolean, value, _style}),
    do: ~s(<c r="#{reference}" t="b"><v>#{if value, do: 1, else: 0}</v></c>)

  defp cell({reference, :error, value, _style}),
    do: ~s(<c r="#{reference}" t="e"><v>#{escape(value)}</v></c>)

  defp cell({reference, :formula_error, {formula, value}, _style}),
    do:
      ~s(<c r="#{reference}" t="e"><f>#{escape(formula)}</f><v>#{escape(value)}</v></c>)
end

defmodule Fixtures.Xlsb do
  @moduledoc false
  # A minimal BIFF12 (.xlsb) writer: just the records calamine reads.

  def write!(path, sheets, shared_strings) do
    sheet_entries =
      sheets
      |> Enum.with_index(1)
      |> Enum.map(fn {{_name, _state, cells}, index} ->
        {"xl/worksheets/sheet#{index}.bin", worksheet(cells)}
      end)

    bundle =
      sheets
      |> Enum.with_index(1)
      |> Enum.map(fn {{name, state, _cells}, index} ->
        state_number = %{visible: 0, hidden: 1, very_hidden: 2}[state]

        record(
          0x009C,
          <<state_number::32-little, index::32-little>> <>
            wide_string("rId#{index}") <> wide_string(name)
        )
      end)

    workbook =
      IO.iodata_to_binary([
        record(0x0083, <<>>),
        record(0x0099, <<0::32-little, 0::32-little>>),
        record(0x008F, <<>>),
        bundle,
        record(0x0090, <<>>),
        record(0x009D, <<0::32-little>>)
      ])

    relationships =
      sheets
      |> Enum.with_index(1)
      |> Enum.map_join(fn {_sheet, index} ->
        ~s(<Relationship Id="rId#{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet#{index}.bin"/>)
      end)

    workbook_relationships = """
    <?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">#{relationships}</Relationships>
    """

    Fixtures.Zip.write!(
      path,
      [
        {"xl/workbook.bin", workbook},
        {"xl/_rels/workbook.bin.rels", workbook_relationships},
        {"xl/styles.bin", styles()},
        {"xl/sharedStrings.bin", shared_strings_part(shared_strings)}
      ] ++ sheet_entries
    )
  end

  # Same style indexes as Fixtures.Xlsx.
  defp styles do
    formats = [{164, "yyyy-mm-dd hh:mm:ss.000"}, {165, "[h]:mm:ss.000"}]
    cell_formats = [0, 14, 22, 164, 21, 46, 165]

    IO.iodata_to_binary([
      record(0x0267, <<length(formats)::32-little>>),
      Enum.map(formats, fn {id, code} ->
        record(0x002C, <<id::16-little>> <> wide_string(code))
      end),
      record(0x0269, <<length(cell_formats)::32-little>>),
      Enum.map(cell_formats, fn id ->
        record(0x002F, <<0xFFFF::16-little, id::16-little, 0::96>>)
      end)
    ])
  end

  defp shared_strings_part(strings) do
    count = length(strings)

    IO.iodata_to_binary([
      record(0x009F, <<count::32-little, count::32-little>>),
      Enum.map(strings, &record(0x0013, <<0>> <> wide_string(&1)))
    ])
  end

  defp worksheet(cells) do
    rows = Enum.group_by(cells, &elem(&1, 0))

    {last_row, last_column} =
      Enum.reduce(cells, {0, 0}, fn cell, {row_max, column_max} ->
        {max(row_max, elem(cell, 0)), max(column_max, elem(cell, 1))}
      end)

    IO.iodata_to_binary([
      record(0x0081, <<>>),
      record(
        0x0094,
        <<0::32-little, last_row::32-little, 0::32-little,
          last_column::32-little>>
      ),
      record(0x0091, <<>>),
      rows
      |> Enum.sort()
      |> Enum.map(fn {row, row_cells} ->
        [
          record(0x0000, <<row::32-little, 0::104>>),
          Enum.map(row_cells, &cell/1)
        ]
      end),
      record(0x0092, <<>>)
    ])
  end

  defp cell({_row, column, kind, value, style}) do
    prefix = <<column::32-little, style::24-little, 0>>

    case kind do
      :number ->
        record(0x0005, prefix <> <<value::float-64-little>>)

      :integer ->
        record(
          0x0002,
          prefix <> <<Bitwise.bor(Bitwise.bsl(value, 2), 2)::32-little>>
        )

      :string ->
        record(0x0006, prefix <> wide_string(value))

      :shared_string ->
        record(0x0007, prefix <> <<value::32-little>>)

      :boolean ->
        record(0x0004, prefix <> <<if(value, do: 1, else: 0)>>)

      :error ->
        record(0x0003, prefix <> <<value>>)
    end
  end

  defp wide_string(text) do
    utf16 = :unicode.characters_to_binary(text, :utf8, {:utf16, :little})
    <<div(byte_size(utf16), 2)::32-little>> <> utf16
  end

  defp record(type, body) do
    type_bytes =
      if type < 0x80,
        do: <<type>>,
        else:
          <<Bitwise.bor(Bitwise.band(type, 0x7F), 0x80), Bitwise.bsr(type, 7)>>

    type_bytes <> size_bytes(byte_size(body)) <> body
  end

  defp size_bytes(size) when size < 0x80, do: <<size>>

  defp size_bytes(size),
    do:
      <<Bitwise.bor(Bitwise.band(size, 0x7F), 0x80)>> <>
        size_bytes(Bitwise.bsr(size, 7))
end

defmodule Fixtures.Ods do
  @moduledoc false
  import Fixtures.Xml

  @namespaces ~s(xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0")

  @doc "`tables` is a list of `{name, rows_xml}`; use `row/2` and `cell/2` to build rows."
  def write!(path, tables) do
    tables_xml =
      Enum.map_join(tables, fn {name, rows} ->
        ~s(<table:table table:name="#{escape(name)}">#{IO.iodata_to_binary(rows)}</table:table>)
      end)

    content =
      ~s(<?xml version="1.0" encoding="UTF-8"?><office:document-content #{@namespaces} office:version="1.2"><office:body><office:spreadsheet>#{tables_xml}</office:spreadsheet></office:body></office:document-content>)

    manifest =
      ~s(<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>)

    Fixtures.Zip.write!(path, [
      {"mimetype", "application/vnd.oasis.opendocument.spreadsheet"},
      {"content.xml", content},
      {"META-INF/manifest.xml", manifest}
    ])
  end

  def row(cells, repeated \\ 1) do
    repeat =
      if repeated > 1,
        do: ~s( table:number-rows-repeated="#{repeated}"),
        else: ""

    ["<table:table-row", repeat, ">", cells, "</table:table-row>"]
  end

  def cell(kind, repeated \\ 1)

  def cell(:empty, repeated), do: ~s(<table:table-cell#{columns(repeated)}/>)

  def cell({:float, value}, repeated),
    do:
      ~s(<table:table-cell office:value-type="float" office:value="#{value}"#{columns(repeated)}><text:p>#{value}</text:p></table:table-cell>)

  def cell({:string, value}, repeated),
    do:
      ~s(<table:table-cell office:value-type="string"#{columns(repeated)}><text:p>#{escape(value)}</text:p></table:table-cell>)

  def cell({:boolean, value}, repeated),
    do:
      ~s(<table:table-cell office:value-type="boolean" office:boolean-value="#{value}"#{columns(repeated)}><text:p>#{value}</text:p></table:table-cell>)

  def cell({:date, value}, repeated),
    do:
      ~s(<table:table-cell office:value-type="date" office:date-value="#{escape(value)}"#{columns(repeated)}><text:p>#{escape(value)}</text:p></table:table-cell>)

  def cell({:time, value}, repeated),
    do:
      ~s(<table:table-cell office:value-type="time" office:time-value="#{escape(value)}"#{columns(repeated)}><text:p>#{escape(value)}</text:p></table:table-cell>)

  defp columns(1), do: ""
  defp columns(repeated), do: ~s( table:number-columns-repeated="#{repeated}")
end

defmodule Fixtures.LibreOffice do
  @moduledoc false

  def convert!(source, extension, scratch) do
    executable =
      System.find_executable("soffice") || System.find_executable("libreoffice") ||
        raise "LibreOffice (soffice) is needed to build the .xls fixtures"

    File.mkdir_p!(scratch)
    copy = Path.join(scratch, Path.basename(source))
    File.cp!(source, copy)

    {output, 0} =
      System.cmd(
        executable,
        ["--headless", "--convert-to", extension, "--outdir", scratch, copy],
        stderr_to_stdout: true
      )

    converted = Path.rootname(copy) <> "." <> extension
    File.exists?(converted) || raise "conversion failed: #{output}"
    File.read!(converted)
  end
end

defmodule Fixtures.Patch do
  @moduledoc false

  @doc "Replaces the single occurrence of `pattern` in `binary`."
  def replace_once!(binary, pattern, replacement) do
    case :binary.matches(binary, pattern) do
      [{offset, length}] ->
        binary_part(binary, 0, offset) <>
          replacement <>
          binary_part(
            binary,
            offset + length,
            byte_size(binary) - offset - length
          )

      matches ->
        raise "expected one match for #{inspect(pattern)}, found #{length(matches)}"
    end
  end
end

directory = __DIR__
scratch = Path.join(directory, "libreoffice_scratch")
fixture = &Path.join(directory, &1)

# Labels go through the shared string table; values use the other cell kinds.
labels =
  ~w(label integer float date date_time date_time_ms time time_ms duration
     negative_duration duration_ms negative_date out_of_range_date div_by_zero
     not_available formula_error boolean string empty large_integer)

types = [
  {"integer", :number, 1234, 0},
  {"float", :number, 0.5, 0},
  {"date", :number, 45_943, 1},
  {"date_time", :number, 45_943.5, 2},
  # 2025-10-13 12:00:00.123
  {"date_time_ms", :number, "45943.50000142361", 3},
  {"time", :number, 0.5, 4},
  # 18:00:00.250
  {"time_ms", :number, "0.7500028935185185", 4},
  # 25:30:00
  {"duration", :number, 1.0625, 5},
  # -1:30:00
  {"negative_duration", :number, -0.0625, 5},
  # 1.5 seconds
  {"duration_ms", :number, "1.7361111111111112E-5", 6},
  {"negative_date", :number, -1, 1},
  {"out_of_range_date", :number, 3_000_000, 1},
  {"div_by_zero", :error, "#DIV/0!", 0},
  {"not_available", :error, "#N/A", 0},
  {"formula_error", :formula_error, {"1/0", "#DIV/0!"}, 0},
  {"boolean", :boolean, true, 0},
  {"string", :string, "Zoë", 0},
  {"empty", nil, nil, 0},
  {"large_integer", :number, 12_345_678_901_234_567, 0}
]

types_rows =
  [{1, [{"A1", :shared_string, 0}, {"B1", :string, "value"}]}] ++
    (types
     |> Enum.with_index(2)
     |> Enum.map(fn {{label, kind, value, style}, row} ->
       label_cell =
         {"A#{row}", :shared_string, Enum.find_index(labels, &(&1 == label))}

       value_cells = if kind, do: [{"B#{row}", kind, value, style}], else: []
       {row, [label_cell | value_cells]}
     end))

Fixtures.Xlsx.write!(fixture.("types.xlsx"), [{"Types", rows: types_rows}],
  shared_strings: labels
)

Fixtures.Xlsx.write!(
  fixture.("epoch_1904.xlsx"),
  [
    {"Dates",
     rows: [
       # 43830 days after 1904-01-01 is 2024-01-01.
       {1,
        [
          {"A1", :number, 43_830, 1},
          {"B1", :number, 43_830.75, 2},
          {"C1", :number, 0.5, 4}
        ]}
     ]}
  ],
  date1904: true
)

offset_rows = [
  {3, [{"C3", :string, "name"}, {"D3", :string, "value"}]},
  {4, [{"C4", :string, "a"}, {"D4", :number, 1}]},
  {6, [{"E6", :boolean, true}]}
]

Fixtures.Xlsx.write!(fixture.("offset.xlsx"), [{"Offset", rows: offset_rows}])

Fixtures.Xlsx.write!(fixture.("sheets.xlsx"), [
  {"Visible", rows: [{1, [{"A1", :string, "visible"}]}]},
  {"Hidden", rows: [{1, [{"A1", :string, "hidden"}]}], state: "hidden"},
  {"Ünïcødé 名前", rows: [{1, [{"A1", :string, "unicode"}]}]},
  {"Very Hidden",
   rows: [{1, [{"A1", :string, "very hidden"}]}], state: "veryHidden"}
])

Fixtures.Xlsx.write!(fixture.("chartsheet.xlsx"), [
  {"Data", rows: [{1, [{"A1", :number, 1}]}]},
  {"Chart", chart: true}
])

# Values in A1 and XFD200000: 200,000 rows × 16,384 columns (3.3 billion
# cells) once expanded into a rectangle. A forged <dimension> claims A1:A1.
Fixtures.Xlsx.write!(fixture.("sparse.xlsx"), [
  {"Sparse",
   dimension: "A1:A1",
   rows: [{1, [{"A1", :number, 1}]}, {200_000, [{"XFD200000", :number, 2}]}]}
])

# .xlsb: zero-based {row, column, kind, value, style}.
xlsb_types = [
  {0, 0, :shared_string, 0, 0},
  {0, 1, :string, "value", 0},
  {1, 0, :shared_string, 1, 0},
  {1, 1, :integer, 1234, 0},
  {2, 0, :shared_string, 2, 0},
  {2, 1, :number, 0.5, 0},
  {3, 0, :shared_string, 3, 0},
  {3, 1, :number, 45_943.5, 2},
  {4, 0, :shared_string, 4, 0},
  {4, 1, :number, 0.5, 4},
  {5, 0, :shared_string, 5, 0},
  {5, 1, :number, -0.0625, 5},
  {6, 0, :shared_string, 6, 0},
  {6, 1, :error, 0x07, 0},
  {7, 0, :shared_string, 7, 0},
  {7, 1, :boolean, true, 0},
  {8, 0, :shared_string, 8, 0},
  {8, 1, :string, "Zoë", 0}
]

xlsb_labels =
  ~w(label integer float date_time time negative_duration div_by_zero boolean string)

Fixtures.Xlsb.write!(
  fixture.("types.xlsb"),
  [
    {"Types", :visible, xlsb_types},
    {"Hidden", :hidden, [{0, 0, :string, "hidden", 0}]},
    {"Offset", :visible, [{2, 2, :string, "C3", 0}]}
  ],
  xlsb_labels
)

Fixtures.Xlsb.write!(
  fixture.("sparse.xlsb"),
  [
    {"Sparse", :visible,
     [{0, 0, :number, 1.0, 0}, {1_000_000, 16_383, :number, 2.0, 0}]}
  ],
  []
)

# .ods
alias Fixtures.Ods

Ods.write!(fixture.("offset.ods"), [
  {"Offset",
   [
     Ods.row([Ods.cell(:empty, 5)], 2),
     Ods.row([
       Ods.cell(:empty, 2),
       Ods.cell({:string, "name"}),
       Ods.cell({:string, "value"})
     ]),
     Ods.row([
       Ods.cell(:empty, 2),
       Ods.cell({:string, "a"}),
       Ods.cell({:float, 1})
     ]),
     Ods.row([Ods.cell(:empty, 5)]),
     Ods.row([Ods.cell(:empty, 4), Ods.cell({:boolean, true})])
   ]}
])

Ods.write!(fixture.("iso.ods"), [
  {"Iso",
   [
     Ods.row([
       Ods.cell({:date, "2024-01-15"}),
       Ods.cell({:date, "1987-05-08T20:10:12.5"}),
       Ods.cell({:date, "2024-01-15T10:00:00.123456789"}),
       Ods.cell({:date, "not-a-date"})
     ]),
     Ods.row([
       Ods.cell({:time, "PT36H00M00S"}),
       Ods.cell({:time, "-PT01H30M00S"}),
       Ods.cell({:time, "PT0H0M1.5S"}),
       Ods.cell({:time, "garbage"})
     ])
   ]}
])

# A1 and the last cell of the largest sheet ODS allows: 1,048,576 rows ×
# 16,384 columns once expanded.
Ods.write!(fixture.("sparse.ods"), [
  {"Sparse",
   [
     Ods.row([Ods.cell({:float, 1})]),
     Ods.row([Ods.cell(:empty)], 1_048_574),
     Ods.row([Ods.cell(:empty, 16_383), Ods.cell({:float, 2})])
   ]}
])

# 10,000 rows, each a single cell element repeated across all 16,384 columns:
# a few KB compressed, 163 million cells once calamine expands the repeats.
Ods.write!(fixture.("repeated_columns.ods"), [
  {"Repeated", List.duplicate(Ods.row([Ods.cell({:float, 1}, 16_384)]), 10_000)}
])

# .xls (via LibreOffice)
for name <- ~w(types offset sheets) do
  File.write!(
    fixture.("#{name}.xls"),
    Fixtures.LibreOffice.convert!(fixture.("#{name}.xlsx"), "xls", scratch)
  )
end

for name <- ~w(types sheets) do
  File.write!(
    fixture.("#{name}.ods"),
    Fixtures.LibreOffice.convert!(fixture.("#{name}.xlsx"), "ods", scratch)
  )
end

# Crafted .xls files, patched from a LibreOffice export with 1 in A1 and 7 in B2.
Fixtures.Xlsx.write!(Path.join(scratch, "base.xlsx"), [
  {"Sheet1", rows: [{1, [{"A1", :number, 1}]}, {2, [{"B2", :number, 7}]}]}
])

base_xls =
  Fixtures.LibreOffice.convert!(Path.join(scratch, "base.xlsx"), "xls", scratch)

# B2 is an RK record (type 0x027E, 10 bytes) at row 1, column 1. Moving it to
# row 65535, column 65535 makes calamine allocate a 65,536 × 65,536 rectangle
# (about 137 GB).
File.write!(
  fixture.("sparse.xls"),
  Fixtures.Patch.replace_once!(
    base_xls,
    <<0x7E, 0x02, 10, 0, 1, 0, 1, 0>>,
    <<0x7E, 0x02, 10, 0, 0xFF, 0xFF, 0xFF, 0xFF>>
  )
)

# The DIMENSIONS record (type 0x0200, 14 bytes) declares rows 0..2 and columns
# 0..2. Claiming 4,294,967,295 rows × 255 columns makes calamine reserve about
# 44 TB before reading any cell.
File.write!(
  fixture.("dimension.xls"),
  Fixtures.Patch.replace_once!(
    base_xls,
    <<0x00, 0x02, 14, 0, 0::32-little, 2::32-little, 0::16-little, 2::16-little,
      0::16>>,
    <<0x00, 0x02, 14, 0, 0::32-little, 0xFFFFFFFF::32-little, 0::16-little,
      255::16-little, 0::16>>
  )
)

# A formatting-only template: the DIMENSIONS record declares A1:IV50000
# (50,000 rows × 256 columns, within .xls limits) but the sheet holds two
# values. It parses whatever :max_cells is.
File.write!(
  fixture.("template.xls"),
  Fixtures.Patch.replace_once!(
    base_xls,
    <<0x00, 0x02, 14, 0, 0::32-little, 2::32-little, 0::16-little, 2::16-little,
      0::16>>,
    <<0x00, 0x02, 14, 0, 0::32-little, 50_000::32-little, 0::16-little,
      256::16-little, 0::16>>
  )
)

# The compound file header's FAT sector count (offset 44) is used as a Vec
# capacity: 0xFFFFFFFF entries is a 16 GB reservation.
<<header::binary-size(44), _fat_sectors::32-little, rest::binary>> = base_xls

File.write!(
  fixture.("fat_length.xls"),
  header <> <<0xFFFFFFFF::32-little>> <> rest
)

File.rm_rf!(scratch)
