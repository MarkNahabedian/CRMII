
using CSV
using DataFrames
using Tables


INPUT_CSV_FILES = [
    "MIT_1.csv",
    "MIT_2.csv"
]

# Headings of the relevalt columns of the input files:
INPUT_COLUMN_PHOTO_CODE = "photo code"
INPUT_COLUMN_ITEM = "item"

DONATION_RECORD_NUMBER = "2024.021"

df = let
    df = DataFrame()
    for f in INPUT_CSV_FILES
        temp_df = CSV.read(f, DataFrame; header=2)
        df = vcat(df, temp_df)
    end
    df
end


# The third component of the catalog number has been edited into the
# INPUT_COLUMN_PHOTO_CODE column:
MODIFIED_PHOTO_CODE_REGEXP =
    r"^(?P<photo_code>[A-Z]+) *(?P<catnum>[0-9]+)$"

function parse_photo_code_column(str)
    m = match(MODIFIED_PHOTO_CODE_REGEXP, str)
    if m isa Nothing
        return nothing
    end
    m.captures
end

catalog_numbers = let 
    re = MODIFIED_PHOTO_CODE_REGEXP
    catnums = []
    for pc in getproperty(df, INPUT_COLUMN_PHOTO_CODE)
        if pc isa Missing
            continue
        end
        m = parse_photo_code_column(pc)
        if m isa Nothing
            push!(catnums, pc)
            continue
        end
        (photo_code, catnum) = m
        push!(catnums, join([photo_code,
                            join([DONATION_RECORD_NUMBER, catnum], ".")],
              " "))
    end
    catnums
end

PROFICIO_IMPORT_HEADERS = [
    "Catalog #" => function(dfrow)
        m = parse_photo_code_column(
            getproperty(dfrow, INPUT_COLUMN_PHOTO_CODE))
        if m isa Nothing
            return ""
        end
        join([DONATION_RECORD_NUMBER,
              m[2]
              ], ".")
    end,
    "Accession #",
    "Object Name",
    "Title" => function(dfrow)
        dfrow."item"
    end,
    "Description" => function(dfrow)
        join(filter(s -> s isa String,
                    map(i -> dfrow[i], 3:5)),
             " | ")
    end,
    "Component Parts",
    "Artist/Maker",
    "Material",
    "Date Made",
    "Record Status" => "DRAFT RECORD",
    "Place Made:Place",
    "Place Made:State-Prov",
    "Place Made:Country",
    "Place Made:Notes",
    "Location:Building" => "CRMII MAIN BUILDING",
    "Location:Room",
    "Location:Shelf",
    "Location:Box" => "2024.021 Edythe Humphries",
    "Location:Notes" => "Protected: Yes.",
    "Object Status" => "STORAGE",
    "Web Ready?" => "N",
    "Count" => "1",
    "Count Unit" => "EA",
    "Condition" => "GOOD",
    "Cond Desc" => "Operational: Yes.",
    "# of Pieces",
    "Other Nbrs" => function(dfrow)
        captures = parse_photo_code_column(
            getproperty(dfrow, INPUT_COLUMN_PHOTO_CODE))
        if captures isa Vector
            "donation record photo code $(captures[1])"
        else
            ""
        end
    end,
    "Notes",
    "Dim-Eng:D",
    "Dim-Eng:H",
    "Dim-Eng:L",
    "Dim-Eng:Diam",
    "Dims-Other:Weight",
    "Dim-Eng:Remarks",
    "Dims-Other:Date Measured",
    "Cataloger" => "Mark Nahabedian",
    "Catalog Date" => "10-03-2024"
]

proficio_headings() = map(PROFICIO_IMPORT_HEADERS) do h
    if h isa Pair
        h.first
    else
        h
    end
end

function process_input_record(dfrow)
    if (let
            v = getproperty(dfrow, INPUT_COLUMN_PHOTO_CODE)
            !(v isa String && length(v) > 0)
        end)
        return missing
    end
    map(PROFICIO_IMPORT_HEADERS) do field
        if field isa String
            ""
        elseif field isa Pair
            if field.second isa String
                field.second
            elseif field.second isa Function
                field.second(dfrow)
            else
                error("unsupported second of field pair $field")
            end
        end
    end
end

let
    results = []
    for row in eachrow(df)
        r = process_input_record(row)
        if r isa Missing
            continue
        elseif r[1] isa AbstractString
            if length(r[1]) > 0
                push!(results, r)
            end
        elseif typeof(r[1]) in [Missing, Nothing]
        else 
            error("unsupported column 1 value $(r[1]), $(typeof(r[1]))")
        end
    end
    tbl = Tables.table(stack(results; dims=1))
    println(length(results), "\t", length(Tables.rows(tbl)))
    CSV.write(joinpath(@__DIR__, "proficio_import.csv"),
              tbl;
              header = proficio_headings())
end
        
