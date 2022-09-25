
using XLSX
using XLSX: readxlsx, sheetnames, row_number, eachrow
using DataStructures
using Parameters
using Logging
using OrderedCollections: OrderedDict, Forward

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

id_index_path(dir) =
    abspath(joinpath(dir, "id_index.tsv"))


############################################################
# CSV file escaping:

rfc4180escape(::Nothing) = ""

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
    id_index = SortedDict(Forward)  # document id -> NicholsDocument
end

# Each sheet that matches DATA_SHEET_NAME_REGEXP describes one box
@with_kw mutable struct NicholsBox
    box_number       # from sheet name
    # first_doc_id     # infer from first row?
    # last_doc_id      # infer from first row
    description
    folders = Vector{NicholsFolder}()
    unfoldered_documents = Vector{NicholsDocument}()
end

@with_kw mutable struct NicholsFolder
    folder_number
    row_number
    description
    extra_rows = Vector{AbstractString}()
    documents = Vector{NicholsDocument}()
end

@with_kw mutable struct NicholsDocument
    box_number = nothing
    folder_number = nothing
    row_number    # row in sheet
    id
    title
    extra_rows = Vector{AbstractString}()
    description
end

function NicholsDocument(collection::NicholsCollection,
                         box::NicholsBox,
                         folder::Union{Nothing, NicholsFolder},
                         row_number, 
                         id::AbstractString, title::AbstractString)
    d = NicholsDocument(box_number = box.box_number,
                        folder_number = (folder == nothing) ?
                            nothing :
                            folder.folder_number,
                        row_number = row_number,
                        id = id,
                        title = title,
                        description = get(DOCUMENT_DESCRIPTIONS, id, nothing))
    if haskey(collection.id_index, d.id)
        @warn "Duplicate key" id=d.id
    else
        collection.id_index[d.id] = d
    end
    d
end

FOLDER_ROW_REGEXP = r"^Folder (?<number>[0-9]+):(?<description>.*)$"
DOCUMENT_ROW_REGEXP = r"^(?<docnum>[0-9]{2}-[0-9]{2}-[0-9]{3})(?<title>.*)$"

function injest(collection::NicholsCollection)
    workbook = readxlsx(joinpath(collection.directory,
                                 "Nichols Collection Contents Copy.xlsx"))
    for sheetname in sheetnames(workbook)
        m = match(DATA_SHEET_NAME_REGEXP, sheetname)
        box = nothing
        folder = nothing
        previous_document = nothing
        if m != nothing    
            box_number = m["box_number"]
            sheet = workbook[sheetname]
            for row in eachrow(sheet)
                try
                    if row[1] isa Missing
                        continue
                    elseif row.row == 1
                        box = NicholsBox(box_number = box_number,
                                         description = row[1])
                        previous_document = nothing
                        push!(collection.boxes,box)
                    elseif box != nothing &&
                        (m = match(FOLDER_ROW_REGEXP, row[1])) != nothing
                        folder = NicholsFolder(folder_number = m["number"],
                                               row_number = row.row,
                                               description = m["description"])
                        previous_document = nothing
                        push!(box.folders, folder)
                    elseif (folder != nothing || box != nothing) &&
                        (m = match(DOCUMENT_ROW_REGEXP, row[1])) != nothing
                        id = m["docnum"]
                        previous_document = NicholsDocument(collection, box, folder, row.row, id, m["title"])
                        if folder != nothing
                            push!(folder.documents, previous_document)
                        else
                            push!(box.unfoldered_documents, previous_document)
                        end
                    else
                        if previous_document != nothing
                            push!(previous_document.extra_rows, row[1])
                        elseif folder != nothing
                            push!(folder.extra_rows, row[1])
                        else
                            @warn("Unrecognized", sheet=sheetname, row=row.row, text=row[1])
                        end
                    end
                catch e
                    if e isa MethodError
                        rethrow()
                    else
                        @error(e, sheet=sheetname, row=row.row)
                    end
                end
            end
        end
    end
    return collection
end

function post_injestion_cleanup(collection::NicholsCollection)
    # Some folder rows were not followed by document rows but the
    # folder's "title" looks like a document row.  For such cases,
    # create the document and add it to the follder.
    for box in collection.boxes
        for folder in box.folders
            if length(folder.documents) == 0
                m = match(DOCUMENT_ROW_REGEXP, strip(folder.description))
                if m != nothing
                    push!(folder.documents,
                          NicholsDocument(collection, box, folder, folder.row_number,
                                          m["docnum"], m["title"]))
                end
            end
        end
    end
end

function add_description_only_documents(collection::NicholsCollection)
    # For each element of DOCUMENT_DESCRIPTIONS for which we don't
    # have a document, add that description to the index.
    for (id, description) in DOCUMENT_DESCRIPTIONS
        if !haskey(collection.id_index, id)
            collection.id_index[id] =
                NicholsDocument(;
                                id = id,
                                row_number = -1,
                                title = nothing,
                                description = description)
        end
    end
end

function write_unmatched_docuemnts_and_descriptions(collection::NicholsCollection)
    collected_documents = OrderedSet{String}()
    collect_doc_ids(docs) =
        for doc in docs
            push!(collected_documents, doc.id)
        end
    for box in collection.boxes
        collect_doc_ids(box.unfoldered_documents)
        for folder in box.folders
            collect_doc_ids(folder.documents)
        end
    end
    function write_set(filename, set)
        open(joinpath(collection.directory, filename), "w") do io
            write(io, join(set, "\n"))
        end
    end
    write_set("docs_without_descs.txt",
              setdiff(collected_documents,
                      keys(DOCUMENT_DESCRIPTIONS)))
    write_set("descs_without_docs.txt",
              setdiff(keys(DOCUMENT_DESCRIPTIONS),
                      collected_documents))
end

function show_counts(collection::NicholsCollection)
    println("Boxes: $(length(collection.boxes))")
    for box in collection.boxes
        println("  Box $(box.box_number)")
        println("    $(length(box.unfoldered_documents)) unfoldered documents")
        for folder in box.folders
            println("    Folder $(folder.folder_number) $(length(folder.documents)) documents, $(folder.description)")
        end
    end
end

DOCUMENTS_TSV_HEADINGS = ["box #", "folder #", "row #", "id", "title", "description"]

function write_document_list(collection::NicholsCollection)
    open(joinpath(collection.directory, "documents2.tsv"), "w") do io
        println(io, join(DOCUMENTS_TSV_HEADINGS, "\t"))
        function writedoc(doc::NicholsDocument)
            println(io, join([
                "$(doc.box_number)",
                "$(doc.folder_number)",
                "$(doc.row_number)",
                "$(doc.id)",
                "$(rfc4180escape(doc.title))",
                "$(rfc4180escape(doc.description))"
                ], "\t"))
        end
        for box in collection.boxes
            map(writedoc, box.unfoldered_documents)
            for folder in box.folders
                map(writedoc, folder.documents)
            end
        end
    end
end


ITEM_NUMBER_INDEX_HEADERS = [ "catalog #", "has title", "has description" ]

function write_item_number_index(collection::NicholsCollection)
    open(id_index_path(collection.directory), "w") do io
        println(io, join(ITEM_NUMBER_INDEX_HEADERS, "\t"))
        for (id, doc) in collection.id_index
            println(io, join([
                id,
                (doc.title != nothing) ? "✓" : "",
                (doc.description != nothing) ? "✓" : ""
            ], "\t"))
        end
    end
end


NICHOLS = NicholsCollection()

# Extract data from the spreadsheet:
open(injestion_log_file(NICHOLS.directory), "w") do logio
    with_logger(SimpleLogger(logio)) do
        injest(NICHOLS)
    end
end

post_injestion_cleanup(NICHOLS)

add_description_only_documents(NICHOLS)

write_unmatched_docuemnts_and_descriptions(NICHOLS)

show_counts(NICHOLS)

write_document_list(NICHOLS)

write_item_number_index(NICHOLS)

function explore(sheet, line)
    workbook[sheet]["A$(line - 1):A$(line + 1)"]
end

