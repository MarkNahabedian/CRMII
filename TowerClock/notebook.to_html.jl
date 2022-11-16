using Pkg

Pkg.add(url="https://github.com/rikhuijzer/PlutoStaticHTML.jl")

using PlutoStaticHTML

build_notebooks(
    BuildOptions(@__DIR__),
    ["tower_clock_notebook.jl"],
    OutputOptions()
)
