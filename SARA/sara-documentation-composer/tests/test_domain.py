from sara.models import Document
from related import  to_yaml, from_yaml
from datetime import date


def test_document_instantiation():
    # --Given--
    # --When--
    doc = Document(title='My Title')

    # --Then--
    assert doc.title == 'My Title'
    print(to_yaml(doc))


def test_load_document_from_yaml():
    # --Given--
    txt = '''
title: My Title
version: 0.0.1
date: "2019-02-20"
reference: RCF-250-ER
custodian: ["foo","bar"]
preparations: []
title_template: Document about {name}'''

    # --When--
    obj = from_yaml(txt,Document)

    # --Then--
    assert isinstance(obj, Document)
    assert obj.date == date(2019,2,20)
    print(obj)