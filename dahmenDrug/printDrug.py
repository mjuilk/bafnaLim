#Importing Libraries#
from datetime import datetime
from zoneinfo import ZoneInfo
import os
import htmlark
import xmltodict
from pathlib import Path
from jinja2 import Environment, FileSystemLoader

# Templates
root_path = Path(__file__).parent
environment = Environment(loader=FileSystemLoader(f"{root_path}/templates/"))
templates = {
    "drugs": "drugsDah1.html",
}

#XML parsing
with open('context.xml', 'r') as xml:
    doc = xmltodict.parse(xml.read())
def changeParams(table):
    title = table.pop('title')
    print("[{}]".format(title))
    for key in table:
        if key == "ageW":
            keyString = "age (Weeks)"
        elif key == "ageM":
            keyString = "age (Months)"
        elif key == "bw":
            keyString = "Body Weight (grams)"
        else:
            keyString = key
        newParam = input("Please enter new {} :\n(Press enter to leave this field blank) ".format(keyString))
        table[key] = newParam

##Date and time code##
###Now###
reportDateNowString = datetime.now(ZoneInfo("Europe/Berlin")).strftime("%d.%m.%Y (%H:%M:%S)")

class LabelTemplate:
    ''' Class to write QR code images and related data into html templates and generate pdf file.
    Attrs:
        result_template: html template using jinja
        context (dict): variables passed to template
    Methods:
        write_template: Save query data to html and pdf files
        embed_resources: Embed external dependencies into the template, i.e. images and CSS styles
        html2pdf: Convert html to pdf
    '''
    def __init__(self, table, subjects, multi_table):
        cwd = os.getcwd()
        self.result_template = environment.get_template(templates[table])
        self.context = {
            "root_path": root_path,
            "subjects": subjects,
            "multi_tables": multi_table
        }

    def write_template(self):
        ''' Save query data to html and pdf files
        '''
        html_file, pdf_file = "results.html", "results.pdf"
        with open(html_file, mode='w') as f:
            f.write(self.result_template.render(self.context))
            print("Saving results")
        self.embed_resources(html_file)
        print("Saving pdf file")
        self.html2pdf(html_file, pdf_file)

    @staticmethod
    def embed_resources(outfile):
        ''' Embed external dependencies into the template, i.e. images and CSS styles
        '''
        packed_html = htmlark.convert_page(outfile, ignore_errors=False)
        with open(outfile, mode='w') as f:
            f.write(packed_html)

    @staticmethod
    def html2pdf(infile, outfile):
        ''' Convert html to pdf
        '''
        command = f"wkhtmltopdf --enable-local-file-access {infile} {outfile}"
        os.system(command)

tables = doc['sheet']['table']
def main():
    data = []
    for i in range(len(tables)):
        changeParams(tables[i])
        data.append(tables[i])
    multi_tables = False
    LT = LabelTemplate("drugs", data, multi_tables)
    LT.write_template()

if __name__=="__main__":
    main()