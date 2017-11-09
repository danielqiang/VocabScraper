import bs4 as bs
import urllib.request
import os
import google_search as g
import docx2txt
import regex as re
from docx import Document
from docx.shared import Pt


# TODO: 1. Implement abbreviation replacements, ie. replace SRAS with "Short Run Aggregate Supply" then re-scrape.
# TODO: 2. Implement more generalized HTML parser for scraping.

def scrape(terms):
    """
    Uses Google API to search for definitions of terms.

    :param terms: List of terms to search using Google API.
    :return: Dict containing terms as keys and definitions as values.
    """

    not_found = []
    definitions = {}
    for term in terms:
        try:
            results = g.google_search(term + " economics investopedia", num=1)
            # Step through API search results dict to obtain Google's featured snippet.
            # For Investopedia, this is stored under the tag 'twitter:description'.
            if 'investopedia' in results['formattedUrl']:
                definitions[term] = results['pagemap']['metatags'][0]['twitter:description']
                continue
            # Now try Wikipedia.
            results = g.google_search(term + " economics wikipedia", num=1)
            if 'wikipedia' in results['formattedUrl']:
                # Wikipedia's snippet is automatically shortened by Google's API.
                # By cleaning the snippet and going to the page URL, we can find
                # the paragraph containing the shortened snippet and return the
                # entire paragraph as the definition.
                definitions[term] = parse(results['formattedUrl'], results['snippet'])
                continue
        # If there's an error, print the error and continue.
        except Exception as e:
            print(e)
            not_found.append(term)

    return definitions, not_found


def parse(term_url, snippet):
    """
    :param term_url: Page URL from Google API.
    :param snippet: Shortened snippet from Google API.
    :return: Full-length snippet grabbed from URL.
    """

    # Remove newlines, ellipses, and extra whitespace from snippet.
    clean_snippet = re.sub("[\n...]", "", snippet).strip()
    # Create a fake user agent to bypass bot detection.
    url_headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                   '(KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36'}
    # Grab and parse website html with BeautifulSoup.
    req = urllib.request.Request(term_url, headers=url_headers)
    sauce = urllib.request.urlopen(req)
    soup = bs.BeautifulSoup(sauce, "lxml")
    for paragraph in soup.body.find_all('p'):
        # If the snippet is part of the paragraph, return the whole paragraph.
        if clean_snippet in paragraph.text:
            return paragraph.text


def read_doc(docname, filename):
    """
    Finds and reads a docx document containing list of terms to be searched.

    :param docname: Name of docx document to be read.
    :param filename: Name of directory containing document.
    :return: List of terms within document.
    """

    os.chdir(os.path.join('..', 'Workbook', filename))
    text = docx2txt.process(docname)
    # Terms are delimited by newlines. filter() removes extra whitespace and newlines.
    terms = list(filter(None, text.splitlines()))
    # Remove terms containing the word "vocabulary".
    for i, term in enumerate(terms):
        if 'vocabulary' in term.lower():
            del terms[:(i + 1)]
            break
    return terms


def add_par(docname, text, fname='Times New Roman', fsize=12, fbold=False, funderline=False):
    """
    Helper function for writing to docx documents.

    :param docname: Name of docx document to write to.
    :param text: Text to write, passed as a string.
    :param fname: Name of font to use.
    :param fsize: Size of font to use.
    :param fbold: Use bold text if true.
    :param funderline: Underline text if true.
    """

    run = docname.add_paragraph().add_run(text)
    font = run.font

    font.name = fname
    font.size = Pt(fsize)
    font.bold = fbold
    font.underline = funderline


def add_mla_header(doc, *args):
    """
    Adds an MLA header to a docx document.
    """
    for arg in args:
        add_par(doc, arg)
    # Add an empty line after writing MLA header.
    add_par(doc, "")


def make_doc(docname, definitions, not_found, *args):
    """
    Creates a Word document with an MLA header, terms, definitions, and words we didn't find.

    :param docname: Name of document.
    :param definitions: Dict with terms as keys and definitions as values.
    :param not_found: List of terms we didn't find.
    :param args: Parts for MLA header.
    """

    doc = Document()
    add_mla_header(doc, *args)
    add_par(doc, "Vocabulary List:", fbold=True)
    # Write terms and definitions.
    for i, term in enumerate(definitions.keys()):
        add_par(doc, str(i+1) + ".   " + term)
        add_par(doc, definitions[term])
    # Write terms we didn't find definitions for.
    if not_found:
        add_par(doc, "Words not found:", funderline=True)
        for term in not_found:
            add_par(doc, term)
    doc.save(docname)


def main():
    """
    Read a Word document containing terms without definitions.
    Google the terms, then create a new Word document with terms and definitions.
    """
    # Grab a list of terms from Word Doc (Econ Vocab List).
    terms = read_doc("Chapter 12 Aggreagate demand and aggregate supply vocabulary.docx",
                     "Ch 12 Aggregate demand and aggregate supply")
    # Clean up terms (remove all parentheses and text between parentheses).
    terms = [re.sub(r'\([^)]*\)', '', term) for term in terms]
    # Scrape Google for term definitions.
    definitions, not_found = scrape(terms)
    # Write words/definitions to new Word Doc.
    make_doc("Chapter 12 Vocabulary.docx", definitions, not_found, "Daniel Qiang",
             "11/6/2017", "Sherman Per. 6", "Chapter 12: Aggregate Demand and Aggregate Supply")

if __name__ == '__main__':
    main()




