# Importation des packages
import argparse
import sys
import os
from jinja2 import Environment, FileSystemLoader
from jinja2.exceptions import TemplateNotFound

import pandas as pd
import related

# Importation des classes Project et Document du fichier models.py
from models import Project, Document

class Cmd:

    def __init__(self):
        self.domain_loaded_from_samples=False
        self._setup_parser()
        self._env = None
        self._template = None

    # Méthode relative à la syntaxe de la ligne de commande d'appel du programme
    def _setup_parser(self):
        parser = argparse.ArgumentParser(description="Cmd")

        template_group = parser.add_argument_group("templates")

        # argument de l'emplacement du répertoire des templates
        template_group.add_argument("--location", action='append', help="Template folders locations")

        # argument de l'emplacement des fichiers project et document
        project_group = parser.add_argument_group("definitions")
        project_group.add_argument("--project",help="Project definition file in yaml format",default=".")
        project_group.add_argument("--document",help="Documentation definition file in yaml format",default=".")

        # argument de l'action render ou sample
        parser.add_argument("action", choices=['render','sample'])

        # argument du nom du template principal
        parser.add_argument("--template", required=True)

        self._parser = parser


    def run(self, args):
        arguments = self.configure(args)

        action_method = getattr(self, arguments.action)
        return action_method()

    # Configuration des arguments de la ligne de commande en variables exploitables
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

        # Si l'action est sample, les templates sont créés à travers la méthode create_samples
        if self.domain_loaded_from_samples :
            self._create_samples()

        # Sinon, l'action est render et donc on utilise les fichiers .yaml project et document
        else:
            with open(arguments.project,'r') as file:
                self.project = related.from_yaml(file,Project)
                print(self.project)

            with open(arguments.document, 'r') as file:
                self.document =related.from_yaml(file,Document)

        self.document.configure_from_project(self.project)

    # Méthode relative à l'action render
    def render(self):
        args = {
            'doc': self.document,
            'project': self.project
        }
        result = self._template.render(args)

        # Création du fichier Asciidoc contenant le rapport
        w = open("temporaryrender.adoc", "w")
        w.write(result)
        w.close()

        # Conversion du fichier Asciidoc en PDF
        os.system("asciidoctor-pdf temporaryrender.adoc")

        # Conversion du fichier Asciidoc en HTML
        os.system("asciidoctor temporaryrender.adoc")

        # Conversion du fichier HTML au format WORD et mise en page
        os.system("pandoc --reference-doc custom-reference.docx temporaryrender.html -o temporaryrender.docx")

        # Conversion du fichier HTML au format EXCEL, récupération des tableaux uniquement
        file_path = 'temporaryrender.html'
        with open(file_path, 'r') as f:
            table = pd.read_html(f.read())
            len_tab = len(table)
            df = pd.concat(table[0:len_tab])
        df.to_excel("temporaryrender.xlsx")

    # Méthode relative à l'action sample
    def sample(self):
        self.render()

    def _create_samples(self):
        self.project = Project.sample()
        print(self.project)

        self.document = Document.sample()
        print(self.document)
        print(' - as yaml - ')
        print(related.to_yaml(self.document))

# Exécution de la fonction principale
if __name__ == '__main__':
    cmd = Cmd()

    # Utilisation de la méthode run de tous les arguments de la ligne de commande sauf le premier
    cmd.run(sys.argv[1:])

