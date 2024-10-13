
import random
import genanki
import docx

class AnkiPackage:
    def __init__(self,deckName,packageNumber):
        self.model_id = random.randrange(1 << 30, 1 << 31)
        self.deck_id = random.randrange(1 << 30, 1 << 31)
        self.model = genanki.Model(self.model_id , deckName,
          fields=[
            {'name': 'Question'},
            {'name': 'Answer'},
            {'name': 'Example'},
          ],
          templates=[
            {
              'name': 'Card type 1',
              'qfmt': '<style> div {text-align: center;font-size:30px;}</style><div>{{Question}}</div>',
              'afmt': '{{FrontSide}}{{Answer}}<br>{{Example}}',
            },
          ])
        self.deck = genanki.Deck(self.deck_id, deckName)
        self.package = genanki.Package(self.deck)
        self.package.media_files = []
        self.file_name = f"C:/Users/f3412/Desktop/out/{deckName}_{packageNumber}.apkg"
    def add(self,question,ankiAnswer,ankiExample):
        note = genanki.Note(model=self.model,fields=[f'{question}', f'{ankiAnswer}',f'{ankiExample}'])
        self.deck.add_note(note)
    def save(self):
        #date = datetime.now().strftime("%Y%m%d%H%M%S")
        self.package.write_to_file(self.file_name)



path = 'C:/Users/f3412/Desktop/C-393.docx'
doc = docx.Document(path)

# Assuming the table is the first element in the document
table = doc.tables[0]

# Iterate over rows and cells to extract text
myAnki = AnkiPackage('myEnglishWords',1)
for row in table.rows:
    q = row.cells[0].text.strip()
    a = row.cells[1].text.strip()
    e = row.cells[2].text.strip()
    myAnki.add(q,a,e)
myAnki.save()