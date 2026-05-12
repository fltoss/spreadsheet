defmodule Spreadsheet.MixProject do
  use Mix.Project

  @github_url "https://github.com/fltoss/spreadsheet"
  @version "0.4.10"

  def project do
    [
      app: :spreadsheet,
      version: @version,
      elixir: "~> 1.18",
      start_permanent: Mix.env() == :prod,
      description: description(),
      package: package(),
      deps: deps()
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
      {:rustler, "~> 0.37.1", runtime: false, optional: true},
      {:decimal, "~> 2.0 or ~> 3.0"},
      {:ex_doc, ">= 0.0.0", only: :dev, runtime: false},
      {:igniter, "~> 0.5", only: [:dev, :test]},
      {:credo, "~> 1.4", only: [:dev, :test], runtime: false},
      {:dialyxir, "~> 1.3", only: [:dev, :test], runtime: false}
    ]
  end

  defp package do
    [
      files: [
        "lib",
        "mix.exs",
        "README*",
        "LICENSE*",
        "native/spreadsheet/.cargo",
        "native/spreadsheet/src",
        "native/spreadsheet/Cargo*",
        "checksum-*.exs"
      ],
      maintainers: ["Wilhelm H Kirschbaum"],
      licenses: ["MIT"],
      links: %{"GitHub" => @github_url}
    ]
  end

  defp description do
    "Fast, memory-efficient Elixir library for parsing spreadsheet files (.xls, .xlsx, .xlsm, .xlsb, .ods) powered by Rust and Calamine. Features precompiled NIFs, smart type handling, and support for hidden sheets."
  end
end
