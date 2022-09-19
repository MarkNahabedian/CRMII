These data files decribe documents in the Charles River Museum of
Industry and Innovation's Nichols Collection.

<tt>Nichols Collection Contents Copy.xlsx</tt> is an Excel spreadsheet
whose sheets list the documents in the collection organized by box and
folder.  This data was gathered by interns during physical examination
of the collection.

The files <tt>Nichols Part _.txt</tt> and <tt>Nichols Part _.pdf</tt>
provide descriptions of the documents that were scanned and
OCRed from a printout of the "finding aid."  They were extracted from
<tt>Nicols Finding Aid TXT and PDF.zip</tt>.

The program `../../src/nichols.jl` reads those files and joins the
documents as described in the spreadsheet with the descriptrions in
the finding aid.  It produces these files

* <tt>descriptions.tsv</tt> is a two column tab separated file that
  associates a document ID (column 1) with its description from the
  finding aid (column 2).

* During processing of the <tt>Nichols Part _.txt</tt> files, any
  input lines that could not be processed are written to
  <tt>injestion_log.txt</tt>.

* <tt>documents2.tsv</tt> (formerly <tt>documents.tsv</tt> until a
  row number column was added) has one row per document. Each row
  provides the box number, folder number, row number in the
  spreadsheet, title (from the spreadsheet) and description (from the
  finding aid).

* The program generated some warnings during the processing of the
  spreadsheet.  These warnings are written to <tt>injestion_log.txt</tt>.

* Typically due to data entry errors, there are some document IDs in
  the spreadsheet that don't match document IDs in the finding aid,
  and vice versa.  The document ids that are present in one source
  but not the other are listed in <tt>descs_without_docs.txt</tt> and
  <tt>docs_without_descs.txt</tt>.

