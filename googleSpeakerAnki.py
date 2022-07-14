from google_speech import Speech
import pandas as pd
from pandas import read_excel
import time
import random
import genanki
from datetime import datetime

class AnkiPackage:
    def __init__(self,deckName):
        self.model_id = random.randrange(1 << 30, 1 << 31)
        self.deck_id = random.randrange(1 << 30, 1 << 31)
        self.model = genanki.Model(self.model_id , deckName,
          fields=[
            {'name': 'Question'},
            {'name': 'QuestionSound'},
            {'name': 'Answer'},
            {'name': 'AnswerSound'},
          ],
          templates=[
            {
              'name': 'Card type 1',
              'qfmt': '<style> div {text-align: center;font-size:30px;}</style><div>{{Question}}<br>{{QuestionSound}}</div>',
              'afmt': '{{FrontSide}}{{Answer}}<br>{{AnswerSound}}',
            },
          ])
        self.deck = genanki.Deck(self.deck_id, deckName)
        self.package = genanki.Package(self.deck)
        self.package.media_files = []
    def add(self,question,questionSnd,ankiAnswer,answerSound):
        self.package.media_files.append(questionSnd)
        self.package.media_files.append(answerSound)
        note = genanki.Note(model=self.model,fields=[f'{question}', f'[sound:{questionSnd}]',f'{ankiAnswer}',f'[sound:{answerSound}]'])
        self.deck.add_note(note)
    def save(self):
        date = datetime.now().strftime("%Y%m%d%H%M%S")
        anki_questions_file = f"anki_question{date}.apkg"
        self.package.write_to_file(anki_questions_file)

myAnki = AnkiPackage('test')

my_sheet = 'Saved translations' # change it to your sheet name, you can find your sheet name at the bottom left of your excel file
file_name = my_sheet+'.xlsx' # change it to the name of your excel file
df = read_excel(file_name, sheet_name = my_sheet)
#print(df.head()) # shows headers with top 5 rows
df = df.sample(frac=0.01)
print(len(df))
for i in range(len(df)):
    czTxt=''
    eTxt=''
    czSnd=None
    oSnd=None
    oLang=''
    for qa in range(2): #loop over Q&A  
        txtType = df.iloc[i, qa]
        txt = df.iloc[i, 2+qa] 
        if(txtType == '捷克语' or txtType == 'Czech'):
            czTxt=txt

        if(txtType == '英语' or txtType == 'English'):
            oTxt=txt
            oLang='en'
        if(txtType == '中文（简体）' or txtType == 'Chinese (Simplified)' ):
            oTxt=txt
            oLang='zh'
    print(czTxt)        
    czSnd = Speech(czTxt, 'cs')
    czSndFile = f"{i}snd.mp3"
    czSnd.save(czSndFile)
    czSnd.play()

    oSnd = Speech(oTxt, oLang)
    oSndFile = f"{i}asnd.mp3"
    oSnd.save(oSndFile)
    oSnd.play()
    
    myAnki.add(czTxt,czSndFile,oTxt,oSndFile)

myAnki.save()
