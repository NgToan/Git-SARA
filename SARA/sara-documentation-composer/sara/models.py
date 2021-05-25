from related import StringField, SetField, SequenceField, DateField, ChildField, MappingField
from related import immutable, mutable

@immutable
class Agent:
    name = StringField()
    role = StringField(required=False)

missingAgent = Agent(name="Missing information",role="Missing information")

@mutable
class Project:
    name = StringField()
    long_name = StringField()
    leaders = SetField(Agent, default=[missingAgent])
    implementation_leaders = SetField(Agent, default=[missingAgent])
    validation_leaders = SetField(Agent, default=[missingAgent])
    version = StringField(default="0.0.1")

    @staticmethod
    def sample():
        return Project(name='My Project',
                       long_name='My very precious project about stars',
                       leaders=[Agent(name='John Doe')],
                       implementation_leaders=[Agent(name='Foo Bar'),Agent(name='Stan Getz')],
                       validation_leaders=[Agent(name='Billy Holliday'),Agent(name='Chet Baker')])

@immutable
class Action:
    name = StringField()
    description = StringField()
    role = StringField(required=False)
    date = DateField(required=False)

@immutable
class Issue:
    id = StringField()
    description = StringField()
    date = DateField()
    comment = StringField(required=False)
    page = SetField(str,required=False)


@immutable
class LinkedDocument:
    id = StringField()
    title = StringField()
    document_reference =  StringField()
    issue = StringField()
    date = DateField()

@mutable
class Document:
    title = StringField(required=True)
    version = StringField(required=False)
    date = DateField(required=False)
    reference = StringField(required=False)
    custodian = ChildField(Agent, required=False)
    preparations = SequenceField(Action, required=False)
    contributions = SequenceField(Action, required=False)
    approbations = SequenceField(Action, required=False)
    issues = SequenceField(Issue, required=False)
    applicable_documents = SequenceField(LinkedDocument, required=False)
    reference_documents = SequenceField(LinkedDocument, required=False)
    title_template = StringField(default="Document about {name}", required=False)


    def configure_from_project(self, project):
        self.title = self.title_template.format(name=project.name)

    @staticmethod
    def sample():
        d = Document("document title")
        d.version = "1.0"
        d.reference = "REC-FOO-A"
        d.custodian = "mister.custodian"
        
        d.preparations.append(Action(name='John Doe',role='Leader', description='Preparation of review',date='2019-01-01'))
        
        d.contributions= [
            Action(name='John Doe', role='Leader', description='Preparation of review', date='2019-01-01'),
            Action(name='John Doe', role='Leader', description='Preparation of review', date='2019-01-01'),
            Action(name='John Doe', role='Leader', description='Preparation of review', date='2019-01-01')]

        d.approbations.append(Action(name='John Doe',role='Leader', description='Preparation of review',date='2019-01-01'))

        d.issues = [
            Issue(id='0.1',date='2018-11-01',page=["2"],description='Typo',comment='file to view the source code until proper documentation is generated.'),
            Issue(id='0.2', date='2018-11-01', page=["2"], description='Typo very long', comment='balbalbalbalbalb'),
            Issue(id='0.3', date='2018-11-01', page=["2","3"], description='Typo', comment='datetime field formatted using formatter.'),
            Issue(id='0.4', date='2018-11-01', page=["2"], description='Typo', comment='balbalbalbalbalb'),
            Issue(id='0.5', date='2018-11-01', page=["2"], description='Typo', comment='Adding your own field types is fairly straightforward '),
        ]

        return d