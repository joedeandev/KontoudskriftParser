# Kontoudskrift Parser

Generates a CSV file from Danske Bank account reports, using PyMuPDF and 
BeautifulSoup for parsing.

To download account reports in bulk, select multiple messages in e-Boks, and
select "Gem lokal kopi" from the menu at the top.

Highly dependent on the formatting of the HTML-converted PDF; any unexpected
changes to the document format or HTML parser will likely break the parser.
PDF parsing is ugly; this probably isn't the best code you've ever read.

Run as `python KontoudfskriftParser.py <path to directory with reports> 
<report filename regex>`, and the output will be saved in the same directory as
`output.csv`.