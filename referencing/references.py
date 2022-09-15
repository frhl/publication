import requests
import xml.etree.ElementTree as ET
import sys
import calendar
​
# Parse PubMed IDs from the command line.
# To get this to work, download the PMID numbers from your bibliography in pubmed, and 
# then pass the file to this code, and send the result to a .bib file.
​
# Usage:
# python3 pubmed_to_bib.py PMID-export-example.txt > PMID-export-bibtex-example.bib
​
pmid_file = sys.argv[1]
with open(pmid_file) as f:
    pmids = [line.rstrip('\n') for line in f]
​
## Fetch XML data from Entrez.
efetch = 'https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi'
r = requests.get(
    '{}?db=pubmed&id={}&rettype=abstract'.format(efetch, ','.join(pmids)))
​
## Loop over the PubMed IDs and parse the XML.
root = ET.fromstring(r.text)
for PubmedArticle in root.iter('PubmedArticle'):
    PMID = PubmedArticle.find('./MedlineCitation/PMID')
    ISSN = PubmedArticle.find('./MedlineCitation/Article/Journal/ISSN')
    Volume = PubmedArticle.find('./MedlineCitation/Article/Journal/JournalIssue/Volume')
    Issue = PubmedArticle.find('./MedlineCitation/Article/Journal/JournalIssue/Issue')
    Year = PubmedArticle.find('./MedlineCitation/Article/Journal/JournalIssue/PubDate/Year')
    Month = PubmedArticle.find('./MedlineCitation/Article/Journal/JournalIssue/PubDate/Month')
    Title = PubmedArticle.find('./MedlineCitation/Article/Journal/Title')
    ArticleTitle = PubmedArticle.find('./MedlineCitation/Article/ArticleTitle')
    MedlinePgn = PubmedArticle.find('./MedlineCitation/Article/Pagination/MedlinePgn')
    Abstract = PubmedArticle.find('./MedlineCitation/Article/Abstract/AbstractText')
    authors = []
    for Author in PubmedArticle.iter('Author'):
        try:
            LastName = Author.find('LastName').text
            ForeName = Author.find('ForeName').text
        except AttributeError:  # e.g. CollectiveName
            continue
        authors.append('{}, {}'.format(LastName, ForeName))
    ## Use InvestigatorList instead of AuthorList
    if len(authors) == 0:
        for Investigator in PubmedArticle.iter('Investigator'):
            try:
                LastName = Investigator.find('LastName').text
                ForeName = Investigator.find('ForeName').text
            except AttributeError:  # e.g. CollectiveName
                continue
            authors.append('{}, {}'.format(LastName, ForeName))
    if Year is None:
        _ = PubmedArticle.find('./MedlineCitation/Article/Journal/JournalIssue/PubDate/MedlineDate')
        Year = _.text[:4]
        Month = '{:02d}'.format(list(calendar.month_abbr).index(_.text[5:8]))
    else:
        Year = Year.text
        if Month is not None:
            Month = Month.text
    try:
        for _ in (PMID.text, Volume.text, Title.text, ArticleTitle.text, MedlinePgn.text, Abstract.text, ''.join(authors)):
            if _ is None:
                continue
            assert '{' not in _, _
            assert '}' not in _, _
    except AttributeError:
        pass
    ## Print the bibtex formatted output.
    try:
        if len(authors)>0:
            print('@Article{{{}{}pmid{},'.format(
                authors[0].split(',')[0].replace(" ", ""), Year, PMID.text))
        else:
            print('@Article{{{}pmid{},'.format(Year, PMID.text))
    except IndexError:
        print('IndexError', pmids, file=sys.stderr, flush=True)
    except AttributeError:
        print('AttributeError', pmids, file=sys.stderr, flush=True)
    print(' Author="{}",'.format(' AND '.join(authors)))
    print(' Title={{{}}},'.format(ArticleTitle.text))
    print(' Journal={{{}}},'.format(Title.text))
    print(' Year={{{}}},'.format(Year))
    if Volume is not None:
        print(' Volume={{{}}},'.format(Volume.text))
    if Issue is not None:
        print(' Number={{{}}},'.format(Issue.text))
    if MedlinePgn is not None:
        print(' Pages={{{}}},'.format(MedlinePgn.text))
    if Month is not None:
        print(' Month={{{}}},'.format(Month))
    print(' ISSN={{{}}},'.format(ISSN.text))
    print('}')



