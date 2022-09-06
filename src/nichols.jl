
using XLSX
using XLSX: readxlsx, sheetnames, row_number, eachrow
using Parameters
using Logging

############################################################
# Input and Output Data Files

# Directory containing the input files for the Nichols Collection:
NICHOLS_DIR = abspath(joinpath(@__FILE__, "../../data/Nichols"))

NICHOLS_SPREADSHEET = joinpath(NICHOLS_DIR,
                               "Nichols Collection Contents Copy.xlsx")

# These files contain descriptive text for each document in the collection:
TEXT_DESCRIPTION_FILE_NAME_REGEXP = r"^Nichols Part [0-9]+.txt$"

TEXT_DESCRIPTION_FILES =
    map(filter(readdir(NICHOLS_DIR)) do filename
            match(TEXT_DESCRIPTION_FILE_NAME_REGEXP, filename) != nothing
        end) do filename
            joinpath(NICHOLS_DIR, filename)
        end

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

injestion_log_file(dir) =
    abspath(joinpath(dir, "injestion_log.txt"))


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


############################################################
# Working with the Spreadsheet

# List sheets:
workbook = readxlsx(NICHOLS_SPREADSHEET)
sheetnames(workbook)

# "Nichols Box 2"
DATA_SHEET_NAME_REGEXP = r"Nichols Box (?<box_number>[0-9]+)"

# What's in a sheet"
# workbook["Nichols Box 1"]["A1:D10"]


@with_kw mutable struct NicholsCollection
    directory = NICHOLS_DIR
    boxes = Vector{NicholsBox}()
end

# Each sheet that matches DATA_SHEET_NAME_REGEXP describes one box
@with_kw mutable struct NicholsBox
    box_number       # from sheet name
    # first_doc_id     # infer from first row?
    # last_doc_id      # infer from first row
    description
    folders = Vector{NicholsFolder}()
end

@with_kw mutable struct NicholsFolder
    folder_number
    description
    documents = Vector{NicholsDocument}()
end

struct NicholsDocument
    id
    title
    description
end

function injest(collection::NicholsCollection)
    workbook = readxlsx(joinpath(collection.directory,
                                 "Nichols Collection Contents Copy.xlsx"))
    for sheetname in sheetnames(workbook)
        m = match(DATA_SHEET_NAME_REGEXP, sheetname)
        box = nothing
        folder = nothing
        if m != nothing    
            box_number = m["box_number"]
            sheet = workbook[sheetname]
            for row in eachrow(sheet)
                try
                    if row.row == 1
                        box = NicholsBox(box_number = box_number,
                                         description = row[1])
                        push!(collection.boxes,box)
                    elseif box != nothing &&
                        (m = match(r"^Folder (?<number>[0-9]+):(?<description>.*)$",
                                   row[1])) != nothing
                        folder = NicholsFolder(folder_number = m["number"],
                                               description = m["description"])
                        push!(box.folders, folder)
                    elseif folder != nothing &&
                        (m = match(r"^(?<docnum>[0-9]{2}-[0-9]{2}-[0-9]{3})(?<title>.*)$",
                                   row[1])) != nothing
                        id = m["docnum"]
                        push!(folder.documents,
                              NicholsDocument(id, m["title"],
                                              get(DOCUMENT_DESCRIPTIONS, id, nothing)))
                    else
                        @warn("Unrecognized", sheet=sheetname, row=row.row, text=row[1])
                    end
                catch e
                    @error(e, sheet=sheetname, row=row.row)
                end
            end
        end
    end
    return collection
end

function show_counts(collection::NicholsCollection)
    println("Boxes: $(length(collection.boxes))")
    for box in collection.boxes
        println("  Box $(box.box_number)")
        for folder in box.folders
            println("    Folder $(folder.folder_number) $(length(folder.documents)) documents, $(folder.description)")
        end
    end
end

NICHOLS = NicholsCollection()

open(injestion_log_file(NICHOLS.directory), "w") do logio
    with_logger(SimpleLogger(logio)) do
        injest(NICHOLS)
    end
end

show_counts(NICHOLS)

