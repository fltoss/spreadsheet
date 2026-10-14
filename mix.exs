defmodule Spreadsheet.MixProject do
  use Mix.Project

  @github_url "https://github.com/fltoss/spreadsheet"
  @version "0.6.4"

  def project do
    [
      app: :spreadsheet,
      version: @version,
      elixir: "~> 1.18",
      start_permanent: Mix.env() == :prod,
      description: description(),
      package: package(),
      docs: docs(),
      deps: deps(),
      # The fixture generator is a script, not a test file.
      test_ignore_filters: [&String.starts_with?(&1, "test/fixtures/")]
    ]
  end

  def application do
    [
      extra_applications: [:logger]
    ]
  end

  defp deps do
    [
      {:rustler_precompiled, "~> 0.8"},
      {:rustler, "~> 0.38.0", runtime: false, optional: true},
      {:ex_doc, ">= 0.0.0", only: :dev, runtime: false},
      {:credo, "~> 1.4", only: [:dev, :test], runtime: false}
    ]
  end

  defp package do
    [
      files: [
        "lib",
        "mix.exs",
        "README*",
        "CHANGELOG*",
        "LICENSE*",
        "native/spreadsheet/.cargo",
        "native/spreadsheet/src",
        "native/spreadsheet/Cargo*",
        "checksum-*.exs"
      ],
      maintainers: ["Wilhelm H Kirschbaum"],
      licenses: ["MIT"],
      links: %{
        "GitHub" => @github_url,
        "Changelog" => "#{@github_url}/blob/main/CHANGELOG.md"
      }
    ]
  end

  defp docs do
    [
      main: "Spreadsheet",
      source_url: @github_url,
      source_ref: "v#{@version}",
      extras: ["CHANGELOG.md"]
    ]
  end

  defp description do
    "Fast, memory-efficient Elixir library for parsing spreadsheet files (.xls, .xlsx, .xlsm, .xlsb, .ods) powered by Rust and Calamine. Features precompiled NIFs, smart type handling, and support for hidden sheets."
  end
end
