import abc
import csv
from pymarc import Record, Field, Subfield


BARTON_CSV = (
    "Title",
    "Creator",
    "Is Part Of",
    "Course Information",
    "Dissertation",
    "Thesis Supervisor",
    "Subject",
    "MESH subjects",
    "Genre",
    "Description",
    "Contents",
    "Other title",
    "Related titles",
    "Series",
    "Publisher",
    "Creation Date",
    "Edition",
    "Format",
    "Frequency",
    "Source",
    "Bound with",
    "Rare collection",
    "Donor note",
    "Local note",
    "Chomsky collection note",
    "Notes",
    "Cast and Production",
    "Language",
    "Reproduction",
    "Identifier",
    "Lin")


def read_barton_csv_file(path):
    with open(path, "r") as input:
        return [ row for row in csv.DictReader(input) ]
    pass

# rec = read_barton_csv_file("c:/Users/Mark Nahabedian/.julia/dev/CRMII/MARC/samples/Excel_20230517_204618.csv")[0]

class Book (object):

    @classmethod
    def from_barton_csv(cls, csv_row):
        book = Book()
        book.title = csv_row["Title"]
        book.creator = csv_row["Creator"]
        book.subject = csv_row["Subject"]
        book.description = csv_row["Description"]
        book.publisher = csv_row["Publisher"]
        return book

    pass

# b = Book.from_barton_csv(rec)

