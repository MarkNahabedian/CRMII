
using XLSX: readxlsx, sheetnames

# Directory containing the input files for the Nichols Collection:
NICHOLS_DIR = abspath(joinpath(@__FILE__, "../../data/Nichols"))


NICHOLS_SPREADSHEET = joinpath(NICHOLS_DIR,
                               "Nichols Collection Contents Copy.xlsx")

# List sheets:
workbook = readxlsx(NICHOLS_SPREADSHEET)
sheetnames(workbook)


############################################################
# CSV file escaping:

function rfc4180escape(str::AbstractString)
    buf = IOBuffer()
    s = split(str, '"')
    for s1 in s
        write(buf, '"')
        write(buf, s1)
        write(buf, '"')
    end
    String(take!(buf))
end

############################################################
### EXTRACTING THE DOCUMENT DESCRIPTIONS:

# These files contain descriptive text for each document in the collection:
TEXT_DESCRIPTION_FILE_NAME_REGEXP = r"^Nichols Part [0-9]+.txt$"

TEXT_DESCRIPTION_FILES =
    map(filter(readdir(NICHOLS_DIR)) do filename
            match(TEXT_DESCRIPTION_FILE_NAME_REGEXP, filename) != nothing
        end) do filename
            joinpath(NICHOLS_DIR, filename)
        end

# We extract those descriptions into this Dict, keyed by document ID:
DOCUMENT_DESCRIPTIONS = Dict{String, String}()

function note_description!(id::AbstractString, text::AbstractString)
    if haskey(DOCUMENT_DESCRIPTIONS, id)
        if length(text) > 0
            previous = DOCUMENT_DESCRIPTIONS[id]
            DOCUMENT_DESCRIPTIONS[id] = previous * " " * strip(text)
        end
    else
        DOCUMENT_DESCRIPTIONS[id] = strip(text)
    end
end

# Regular expression for the input description lines:
FILE_DESCRIPTION_REGEXP = r"^(?<id>[0-9]+-[0-9]+-[0-9]+)(?<desc>[^0-9].*)$"

# Descrioption records that aren't recognized are written to a file
# whose name is based on that of the input file:
function ignore_file_path(description_file_path)
    d, f = splitdir(description_file_path)
    b, ext = splitext(f)
    joinpath(d, b * "-ignored" * ext)
end

function extracted_descriptions_path()
    joinpath(NICHOLS_DIR, "descriptions.tsv")
end


"""
    read_text_descriptions(filepath::AbstractString)
Add extracted document descriptions to `DOCUMENT_DESCRIPTIONS`.
Write unrecognized records to an "-ignored" file.
"""
function read_text_descriptions(filepath::AbstractString)
    line_number = 0
    open(ignore_file_path(filepath), "w") do ignored
        previous_id = nothing
        for line in eachline(filepath)
            line_number += 1
            m = match(FILE_DESCRIPTION_REGEXP, line)
            if m == nothing
                if previous_id == nothing
                    println(ignored, "$line_number\t$(rfc4180escape(line))")
                else
                    note_description!(previous_id, line)
                end
            else
                note_description!(m["id"], m["desc"])
                previous_id = m["id"]
            end
        end
    end
end
    
# Extract the descriptions:
map(read_text_descriptions, TEXT_DESCRIPTION_FILES)

function write_descriptions()
    open(extracted_descriptions_path(), "w") do io
        for (id, text) in DOCUMENT_DESCRIPTIONS
            println(io, "$id\t$(rfc4180escape(text))")
        end
    end
end

write_descriptions()

