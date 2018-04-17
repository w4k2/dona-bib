#!/usr/bin/env python
import sys
import pandas as pd
from datetime import datetime
import pprint
from unidecode import unidecode
from math import isnan
from bibtexparser.bwriter import BibTexWriter
from bibtexparser.bibdatabase import BibDatabase

pp = pprint.PrettyPrinter(indent=4)

# Check and load command line arguments
argv = sys.argv
if len(argv) != 2:
    print("W argumencie proszę podać ścieżkę do pliku xlsx.")
    exit()
xlsx_path = argv[1]

# Get generation timestamp
timestamp = datetime.strptime(
    ''.join(xlsx_path.split('.')[0].split('_')[-2:]),
    '%Y%m%d%H%M%S%f')
print("# Podsumowanie wygenerowane %s" % timestamp)

# Load xlsx file
data = pd.read_excel(xlsx_path)
keys_to_drop = ['Wydział główny', "Status pracy", "Numer kierunku badań", "Numer ewidencyjny pracownika PWroc.", "Numer ewidencyjny pracownika PWroc..1", "Jednostka organizacyjna główna", "Afiliacja", "Afiliacja.1", 'Jednostka organizacyjna powiązana', 'Wydział powiązany', "Uwagi dotyczące dysertacji", "Numer ewidencyjny promotora", "Wskaźnik open access", "Uwagi dotyczące źródeł finansowania pracy"]
data = data.drop(keys_to_drop, axis=1)

key_translator = {
    'Nr syst.': "system_id",
    'Numer archiwalny': "archival_id",
    'Rodzaj pracy': "type",
    "Rok zaliczenia": "year",
    "Data zdokumentowania": "archival_date",
    'Nazwisko i imię pierwszego autora': "first_author",
    'Nazwisko i imię kolejnego autora': "other_authors",
    'Tytuł pracy': "title",
    'Miejsce wydania': "place",
    'Nazwa wydawcy': "publisher",
    'Data wydania': "date",
    'Liczba stron': "pages",
    'W: ': "in",
    'Tytuł serii': "series_title",
    'Oznaczenie części': "part_description",
    'Tytuł podserii': "subseries_title",
    'ISSN': "issn",
    'Numeracja': "numeration",
    'Nazwisko i imię promotora': "advisor",
    'Kod języka': "language",
    'ISBN': "isbn",
    'ISSN.1': "issn2",
    'Punktacja czasopisma na liście MNiSW': "mnisw",
    'Rok ukazania się listy/rok obowiązywania listy': "isi_year",
    'Lista Filadelfijska': "isi",
    'Impact Factor': "if",
    'Rok którego dotyczy IF': "if_year",
    'Punkty Web of Science (tylko dla referatów)': "wos",
    'Data dodania punktów WOS': "wos_year",
    'Numer zlecenia': "order",
    'Numer grantu': "grant",
    'Numer projektu Działu Zarządzania Projektami': "project"
}
data = data.rename(index=str, columns=key_translator)

number_of_documents = data.shape[0]
print("# Wczytano %i dokumentów" % number_of_documents)
column_names = list(data)
print(column_names)
document_types = data['type'].unique()
category_translator = {
    "Redakcja monografii i prac zbiorowych": "proceedings",
    "Doktorat": "phdthesis",
    "Referat konferencyjny": "inproceedings",
    "Artykuł": "article",
    "Rozdział w książce": "incollection",
    "Książka": "book",
    "Podręcznik": "manual",
    "Raport serii SPR": "techreport",
    "Raport serii PRE": "techreport",
    "Habilitacja": "book",
    "Referat lub komunikat niepublikowany": "unpublished",
    "Skrypt": "manual",
    "Komunikat konferencyjny": "inproceedings",
    "Rozdział w monografii": "incollection",
    "Monografia": "book",
    "Redakcja mat. konferencyjnych": "proceedings",
    "Streszczenie": "misc",
    "Recenzja": "misc"
}

# Initialize parser
db = BibDatabase()

# Name joiner
def joined_authors(document):
    if 'other_authors' in document:
        first = document['first_author']
        other = document['other_authors']
        joined = '\n'.join([first, other])
        joined = joined.replace('\n', ' and ')
        return joined
    else:
        return document['first_author']

vc = 1000
def explode_in(value):
    #print("\n---")
    #print(value)
    elements = value.split(',')
    elements = [e.strip() for e in elements]

    acc = []
    pages = ""
    for e in elements:
        a = e
        if a.find(' rys.') != -1 or a.find('bibliogr. ') != -1 or a.find(' tab.') != -1:
            #print('--- %s' % a)
            continue
        if a.find('.s.') != -1:
            loc = a.find('.s.')
            #print("--- %s [%s]" % (a[:loc+1], a[loc+3:]))
            pages = a[loc+4:]
            a = a[:loc+1]
        acc.append(a)
        #print(a)
    acc = ', '.join(acc)
    return acc, pages


def explode_in_a(value):
    #print("\n---")
    #print(value)
    elements = value.split(',')
    elements = [e.strip() for e in elements]

    acc = []
    volume = None
    number = None
    for e in elements:
        a = e
        if a.find(' rys.') != -1 or a.find('bibligr.') != -1 or a.find('bibliogr.') != -1 or a.find('bibliogr ') != -1 or a.find('Bibliogr. ') != -1 or a.find(' tab.') != -1 or a[:3] == 's. ' or a[:5] == "vol. " or a[:4] == "vol " or a[:3] == "T. " or a[:5] == "art. " or a[:5] == "Summ." or a[:3] == "nr " or a[:2] == "nr" or a[:3] == "R. " or a[:6] == "part. " or a == "5" or a == "6" or a == "7" or a[:4] == "pt. " or a[:3] == "z. " or a[:9]=="Streszcz." or a[:7] =="suppl. ":
            #print('--- %s' % a)
            if a[:5] == "vol. ":
                volume = a[5:]
                #print("VOLUME %s" % volume)
            if a[:3] == "T. ":
                volume = a[3:]
                #print("VOLUME %s" % volume)
            if a[:4] == "vol ":
                volume = a[4:]
                #print("VOLUME %s" % volume)
            if a[:3] == "nr ":
                number = a[3:]
                #print("NUMBER %s" % number)
            elif a[:2] == "nr":
                number = a[2:]
                #print("NUMBER %s" % number)
            continue
        acc.append(a)
        #print(a)
    acc = ', '.join(acc)[:-5]
    #print("ACC = %s" % acc)
    #print(volume)
    #print(number)
    return acc, volume, number

k_id_acc = []
letters = 'abcdefghijklmnopqrstuvwxyz'
def key_id(document):
    a = "%s%i" % (
        document['first_author'].split(',')[0].lower(),
        document['year']
    )
    a = unidecode(a)
    k_id_acc.append(a)
    count = k_id_acc.count(a)
    a += letters[count-1]
    a = a.replace('-', '')
    return a

# Iterate groups
i = 1
for document_type in document_types:
    i -= 1
    documents_in_type = data.loc[data['type'] == document_type]
    category = category_translator[document_type]
    print("| %s - %s (%i)" % (document_type, category,
                              documents_in_type.shape[0]))
    for index, row in documents_in_type.iterrows():
        document = row.to_dict()
        document = {k: document[k] for k in document if type(document[k]) is str or (not type(document[k]) is str and not isnan(document[k]))}
        # pp.pprint(document)
        authors = unidecode(joined_authors(document))
        identifier = key_id(document)

        # Base fields
        dict = {
            "ID": identifier,
            "ENTRYTYPE": category,
            "title": document['title'][:-2],
            "year": str(document['year']),
            "meta": "%s,%s,%s,%s,%s,%s,%s,%s,%s,%.3f,%.0f,%s,%s" % (
                document["language"],
                document["system_id"],
                document["archival_id"],
                document["archival_date"],
                document["project"] if 'project' in document else '',
                document["grant"] if 'grant' in document else '',
                document['mnisw'].split("\n")[0] if 'mnisw' in document else '',
                'true' if document['isi'] == 'Tak' else 'false',
                document['isi_year'].split('\n')[-1] if 'isi_year' in document else '',
                document['if'] if 'if' in document else 0,
                document['wos'] if 'wos' in document else 0,
                document['wos_year'] if 'wos_year' in document else '',
                document_type
            ),
        }
        # General field
        if 'pages' in document:
            pages = document['pages']
            loc = pages.find('s.')
            pages = pages[:loc-1]
            dict.update({"pages": pages})
        if 'issn' in document:
            dict.update({"ISSN": document['issn']})
        if 'issn2' in document:
            dict.update({"ISSN": document['issn2']})
        if 'series_title' in document:
            dict.update({"series": document["series_title"]})

        # General authorship
        if category in ['misc', 'book', 'incollection', 'inproceedings', 'manual', 'unpublished', 'techreport', 'article', 'phdthesis']:
            dict.update({
                "author": authors
            })

        # Category fields
        if category == 'book':
            dict.update({
                "publisher": "%s %s" % (
                    document["publisher"],
                    document["place"][:-2]
                )
            })
        if category == 'proceedings':
            dict.update({
                "editor": authors,
                "publisher": "%s %s" % (
                    document["publisher"],
                    document["place"][:-2]
                )
            })
        if category == 'incollection':
            dict.update({
                "booktitle": document["in"]
            })
            if 'numeration' in document:
                dict.update({'volume': document['numeration']})
        if category == 'inproceedings':
            booktitle, pages = explode_in(document["in"])
            dict.update({
                "booktitle": booktitle,
                "pages": pages
            })
        if category == 'manual':
            dict.update({
                'organization': document["publisher"],
                'address': document["place"]
            })
        if category == 'misc':
            dict.update({
                "howpublished": document['in']
            })
        if category == 'techreport':
            dict.update({
                'number': document['numeration']
            })
        if category == 'article':
            journal, volume, number = explode_in_a(document["in"])
            dict.update({
                'journal': journal
            })
            if not volume is None:
                dict.update({'volume': volume})
            if not number is None:
                dict.update({'number': number})
        if category == 'phdthesis':
            dict.update({
                'institution': 'Wroclaw University of Science and Technology, Department of Electronics'
            })

        for key in dict:
            dict[key] = unidecode(dict[key])
            dict[key] = dict[key].replace('&', '\\&')
            dict[key] = dict[key].replace('_', '\\_')

        db.entries.append(dict)
        # pp.pprint(dict)
        # break

# Write to file
writer = BibTexWriter()
writer.indent = '    '     # indent entries with 4 spaces instead of one
writer.comma_first = True  # place the comma at the beginning of the line
with open('bibtex.bib', 'w') as bibfile:
    bibfile.write(writer.write(db))
