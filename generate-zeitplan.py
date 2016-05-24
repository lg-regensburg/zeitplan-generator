#!/usr/bin/python3
# import pydocx
from pydocx.export import PyDocXHTMLExporter
from bs4 import BeautifulSoup

# must be docX
file = "2016_gala16_zeitplan-Stand_2016-05-10.docx"

exporter = PyDocXHTMLExporter(file)
html = exporter.export()

soup = BeautifulSoup(html, "html5lib")

for paragraph in soup.find_all('p'):
    paragraph['style'] = 'text-align: center;'
    if ("markierte" in paragraph.prettify()):
        paragraph['style'] = "background-color: #C0C0FF;"

for span in soup.find_all('span'):
    span.unwrap()

for table in soup.find_all('table'):
    del(table['border'])
    table['class']='zeitplan'
    table['style']="width:650px"

for td in soup.find_all('td'):
    # if for child in td.children:
    if ('Gala-Programm' in td.prettify()):
        td['height']='80px'
        td['style'] = 'vertical-align:middle; text-align:center;'
    for stong in td.find_all('strong'):
        if not ('Gala-Programm' in td.strong.string):
            td.strong.unwrap()
            td['style'] = "background-color: #C0C0FF"

time_table = str(soup.body.prettify())
print(time_table)
time_table = time_table.replace('<body>','').replace('</body>','').replace('\\n','')

outfile = open("zeitplan.php", 'w', encoding='utf-8')
outfile.write("<?php $title = \"Zeitplan\";\n $menu['current'] = \"zeitplan\"; ?>\n<?php require(\"_header.inc.php\"); ?>\n\n<h2>Zeitplan</h2>\n<br>\n")
outfile.write(time_table)
outfile.write('\n<?php require("_footer.inc.php"); ?>')

outfile.close()

print('generate pdf and upload zeitplan.php and pdf')
input('press enter to close.')
