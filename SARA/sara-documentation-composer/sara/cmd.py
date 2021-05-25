import argparse
import sys
import os
from jinja2 import Environment, FileSystemLoader
from jinja2.exceptions import TemplateNotFound

import pandas as pd
import related

from models import Project, Document

class Cmd:

    def __init__(self):
        self.domain_loaded_from_samples=False
        self._setup_parser()
        self._env = None
        self._template = None

    def _setup_parser(self):
        parser = argparse.ArgumentParser(description="Cmd")

        template_group = parser.add_argument_group("templates")
        template_group.add_argument("--location", action='append', help="Template folders locations")

        project_group = parser.add_argument_group("definitions")
        project_group.add_argument("--project",help="Project definition file in yaml format",default=".")
        project_group.add_argument("--document",help="Documentation definition file in yaml format",default=".")

        parser.add_argument("action", choices=['render','sample'])
        parser.add_argument("--template", required=True)

        self._parser = parser


    def run(self, args):
        arguments = self.configure(args)

        action_method = getattr(self, arguments.action)
        return action_method()

    def configure(self, args):
        """
        Configure command from arguments
        :param args:
        :return:
        """
        arguments = self._parser.parse_args(args=args)

        self.configure_template_runtime(arguments)
        self.configure_domain_objects(arguments)

        return arguments


    def configure_template_runtime(self, arguments):
        self._env = Environment(loader=FileSystemLoader(arguments.location))
        try:
            self._template = self._env.get_template(arguments.template)
        except TemplateNotFound as tnf:
            print(tnf,arguments.location,arguments.template)
            raise tnf


    def configure_domain_objects(self, arguments):
        self.domain_loaded_from_samples = ( arguments.action == 'sample' )
        if self.domain_loaded_from_samples :
            self._create_samples()
        else:
            with open(arguments.project,'r') as file:
                self.project = related.from_yaml(file,Project)
                print(self.project)

            with open(arguments.document, 'r') as file:
                self.document =related.from_yaml(file,Document)

        self.document.configure_from_project(self.project)

    def render(self):
        args = {
            'doc': self.document,
            'project': self.project
        }
        result = self._template.render(args)

        # Creation du fichier Asciidoc contenant le rapport
        w = open("temporaryrender.adoc", "w")
        w.write(result)
        w.close()

        # Conversion du fichier Asciidoc en PDF et ouverture
        os.system("asciidoctor-pdf temporaryrender.adoc")

        # Conversion du fichier Asciidoc en HTML et ouverture
        os.system("asciidoctor temporaryrender.adoc")

        # Conversion du fichier HTML au format WORD et Mise en page
        os.system("pandoc --reference-doc custom-reference.docx temporaryrender.html -o temporaryrender.docx")

        # Conversion du fichier HTML au format EXCEL, recuperation des tableaux uniquement
        file_path = 'temporaryrender.html'
        with open(file_path, 'r') as f:
            table = pd.read_html(f.read())
            len_tab = len(table)
            df = pd.concat(table[0:len_tab])
        df.to_excel("temporaryrender.xlsx")

    def sample(self):
        self.render()

    def _create_samples(self):
        self.project = Project.sample()
        print(self.project)

        self.document = Document.sample()
        print(self.document)
        print(' - as yaml - ')
        print(related.to_yaml(self.document))

if __name__ == '__main__':
    cmd = Cmd()
    cmd.run(sys.argv[1:])

