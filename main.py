import os
import string as s
import bs4 as bs
import urllib.request
import docx2txt
import tkinter
import googleapiclient.errors
import googleapiclient.discovery
import regex as re
from tkinter.filedialog import askopenfilename
from docx import Document
from docx.shared import Pt


def google_search(query, api_key, cse_id, **kwargs):
    """
    Uses Google API to query a search and returns the search results. Requires valid API key and CSE ID.

    :param query: Query, passed as string.
    :param api_key: Google API Key, passed as string.
    :param cse_id: Google CSE ID, passed as string.
    :param kwargs: Arguments for Google API list() function.
    """

    try:
        service = googleapiclient.discovery.build("customsearch", "v1", developerKey=api_key)
        res = service.cse().list(q=query, cx=cse_id, **kwargs).execute()
        return res['items'][0]
    except googleapiclient.errors.HttpError:
        exit("Missing Google API Key and/or CSE ID.")


def scrape(terms, api_key, cse_id):
    """
    Google a definition for each term provided and return them.

    :param terms: List of terms to search using Google API.
    :param api_key: Google API Key, passed as string.
    :param cse_id: Google CSE ID, passed as string.
    :rtype: dict
    """

    # Clean up terms for Google Search. Remove all parentheses and text between parentheses then map the
    # old terms to the new ones.
    clean_terms = [re.sub(r'\([^)]*\)', '', term) for term in terms]
    not_found = []
    definitions = {}
    for i, clean_term in enumerate(clean_terms):
        print("Googling " + clean_term + "...")
        # Try Investopedia.
        results = google_search(clean_term + " definition economics investopedia", api_key, cse_id, num=1)
        if 'investopedia' in results['formattedUrl'] and \
                'twitter:description' in results['pagemap']['metatags'][0]['twitter:description']:
            definitions[terms[i]] = results['pagemap']['metatags'][0]['twitter:description']
        # Now try Wikipedia.
        else:
            results = google_search(clean_term + " definition economics wikipedia", api_key, cse_id, num=1)
            # Google shortens their website descriptions with ellipses if they are too long.
            # Wikipedia's descriptions are almost always shortened, so we visit Wikipedia
            # and grab the rest of the text snippet directly from the site html.
            definitions[terms[i]] = parse(results['formattedUrl'], results['snippet'])
    return definitions, not_found


def parse(term_url, snippet):
    """
    Parses HTML text to grab the full paragraph containing the shortened snippet.
    Helper function for scrape().

    :param term_url: Page URL from Google API.
    :param snippet: Shortened snippet from Google API.
    :return: Full-length snippet grabbed from URL.
    """

    # Remove ellipses from snippet.
    clean_snippet = re.sub("[...]", "", snippet)
    # Create a fake user agent to bypass bot detection.
    url_headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                   '(KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36'}
    # Grab and parse website HTML with BeautifulSoup.
    r = urllib.request.Request(term_url, headers=url_headers)
    soup = bs.BeautifulSoup(urllib.request.urlopen(r), "lxml")
    for paragraph in soup.body.find_all('p'):
        # If the snippet is part of the paragraph, return the whole paragraph.
        if clean_snippet in paragraph.text:
            return paragraph.text


def read_doc(filepath):
    """
    Parses a Word file containing vocabulary terms and returns a list of terms..

    :param filepath: Path for docx file containing search terms.
    """
    # Two lists - Chars allowed in terms and words not allowed in terms
    allowed_chars = s.ascii_letters + s.digits + '-() '
    not_allowed_words = ["Vocab", "Chapter"]

    text = docx2txt.process(filepath)
    # Split text with newlines as delimiters.
    text = list(filter(None, text.split('\n')))

    terms = []
    for string in text:
        # If every char in string is allowed and string contains no not-allowed words, append to terms
        if all(char in allowed_chars for char in string) and not any(word in string for word in not_allowed_words):
            terms.append(string)
    return terms


def add_par(docname, text, font='Times New Roman', font_size=12, bold_text=False, underline_text=False):
    """
    Helper function for writing to docx documents.

    :param docname: Name of docx document to write to.
    :param text: Text to write, passed as a string.
    :param font: Type of font to use (Times New Roman, Calibri, etc.)
    :param font_size: Size of font to use.
    :param bold_text: Use bold text if true.
    :param underline_text: Underline text if true.
    """

    run = docname.add_paragraph().add_run(text)
    paragraph = run.font

    paragraph.name = font
    paragraph.size = Pt(font_size)
    paragraph.bold = bold_text
    paragraph.underline = underline_text


def add_mla_header(doc, **kwargs):
    """
    Adds an MLA header to a docx document. Helper function for makedoc().

    :param doc: Docx document object.
    :param args: Parts for MLA header.
    """
    for kwarg in kwargs:
        add_par(doc, kwarg)
    # Add an empty line after writing MLA header.
    add_par(doc, "")


def make_doc(docname, filepath, definitions, not_found, **kwargs):
    """
    Creates a Word document with economics terms and their definitions.

    :param docname: Name of document.
    :param filepath: Path (location) to save document.
    :param definitions: Dict with terms as keys and definitions as values.
    :param not_found: List of terms we didn't find.
    :param args: Parts for MLA header.
    """

    os.chdir(filepath)
    doc = Document()
    # My preferred MLA formatting. Can be modified as needed.
    add_mla_header(doc, **kwargs)
    add_par(doc, "Vocabulary List:", bold_text=True)
    # Write terms and definitions.
    for i, term in enumerate(definitions.keys()):
        add_par(doc, str(i+1) + ".   " + term)
        add_par(doc, definitions[term])
    # Write terms we didn't find definitions for.
    if not_found:
        add_par(doc, "Words not found:", underline_text=True)
        for term in not_found:
            add_par(doc, term)
    doc.save(docname)


def main(makefile, **kwargs):
    """
    Runs the script.
    """
    # Key and ID for Google's customsearch API.
    key = ""
    cse_id = ""
    # Opens an instance of tkinter for user file selection.
    tkinter.Tk().withdraw()  # Close the root window
    vocab_fpath = askopenfilename()
    # Make sure the user selected a file.
    if vocab_fpath == '':
        exit("No input file selected.")
    print("File received.")
    terms = read_doc(vocab_fpath)
    definitions, not_found = scrape(terms, api_key=key, cse_id=cse_id)
    make_doc(makefile, os.path.dirname(vocab_fpath), definitions, not_found, **kwargs)
    os.startfile(makefile)


if __name__ == '__main__':
    main("Chapter 1 Vocabulary.docx", name="Daniel Qiang", date="2/4/18",
         period="Sherman Per. 6", subject="Macroeconomic Objectives")
