from google_speech import Speech
import pandas as pd
from pandas import read_excel
import time
import random
import genanki
from datetime import datetime
class AnkiPackage:
    def __init__(self,pkgPath,deckName):
        self.anki_folder = pkgPath;
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
        anki_questions_file = self.anki_folder + f"anki_question{date}.apkg"
        self.package.write_to_file(anki_questions_file)
anki_folder = 'C:/Users/f3412/Desktop/anki/'
myAnki = AnkiPackage(anki_folder, 'test')

my_sheet = 'Saved translations' # change it to your sheet name, you can find your sheet name at the bottom left of your excel file
file_name = 'C:\\Users\\f3412\\Desktop\\'+my_sheet+'.xlsx' # change it to the name of your excel file
df = read_excel(file_name, sheet_name = my_sheet)
#print(df.head()) # shows headers with top 5 rows
df = df.sample(frac=0.02)
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
    czSnd.save(myAnki.anki_folder + f"{i}snd.mp3")
    czSnd.play()
    
    # time.sleep(3)
    oSnd = Speech(oTxt, oLang)
    oSnd.save(myAnki.anki_folder + f"{i}asnd.mp3")
    oSnd.play()
    # time.sleep(3)
    # czSnd.play()#for review
    # time.sleep(2)
    
    question = czTxt
    questionSnd = myAnki.anki_folder + f"{i}snd.mp3"
    ankiAnswer = oTxt
    answerSound = myAnki.anki_folder + f"{i}asnd.mp3"
    myAnki.add(question,questionSnd,ankiAnswer,answerSound)
    

    
#print(df.iat[0,1])

# say "Hello World"
text = "练习结束，白白了。"
lang = "zh"
speech = Speech(text, lang)
speech.play()


myAnki.save()

# you can also apply audio effects while playing (using SoX)
# see http://sox.sourceforge.net/sox.html#EFFECTS for full effect documentation
#sox_effects = ("speed", "0.5")
#speech.play(sox_effects)

        