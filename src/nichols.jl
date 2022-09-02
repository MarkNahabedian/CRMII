
using XLSX: readxlsx, sheetnames

# Directory containing the input files for the Nichols Collection:
NICHOLS_DIR = abspath(joinpath(@__FILE__, "../../data/Nichols"))


NICHOLS_SPREADSHEET = joinpath(NICHOLS_DIR,
                               "Nichols Collection Contents Copy.xlsx")

# List sheets:
workbook = readxlsx(NICHOLS_SPREADSHEET)
sheetnames(workbook)


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

# Regular expression for the input description lines:
FILE_DESCRIPTION_REGEXP = r"^(?<id>[0-9]+-[0-9]+-[0-9]+)(?<desc>[^0-9].*)$"

# Descrioption records that aren't recognized are written to a file
# whose name is based on that of the input file:
function ignore_file_path(description_file_path)
    d, f = splitdir(description_file_path)
    b, ext = splitext(f)
    joinpath(d, b * "-ignored" * ext)
end


"""
    read_text_descriptions(filepath::AbstractString)
Add extracted document descriptions to `DOCUMENT_DESCRIPTIONS`.
Write unrecognized records to an "-ignored" file.
"""
function read_text_descriptions(filepath::AbstractString)
    line_number = 0
    open(ignore_file_path(filepath), "w") do ignored
        for line in eachline(filepath)
            line_number += 1
            m = match(FILE_DESCRIPTION_REGEXP, line)
            if m == nothing
                println(ignored, "$line_number\t$line")
            else
                DOCUMENT_DESCRIPTIONS[m["id"]] = m["desc"]
            end
        end
    end
end
    
# Extract the descriptions:
map(read_text_descriptions, TEXT_DESCRIPTION_FILES)

