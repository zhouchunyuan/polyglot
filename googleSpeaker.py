from google_speech import Speech
import pandas as pd
from pandas import read_excel
import time

my_sheet = '已保存的译文' # change it to your sheet name, you can find your sheet name at the bottom left of your excel file
file_name = 'C:\\Users\\f3412\\Desktop\\已保存的译文.xlsx' # change it to the name of your excel file
df = read_excel(file_name, sheet_name = my_sheet)
#print(df.head()) # shows headers with top 5 rows
df = df.sample(frac=1)
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
    czSnd = Speech(czTxt+ '.a', 'cs')#'a' is a psudo sound for bug fixing
    czSnd.play()
    time.sleep(3)
    oSnd = Speech(oTxt+ '.a', oLang)#'a' is a psudo sound for bug fixing
    oSnd.play()
    time.sleep(3)
    czSnd.play()#for review
    time.sleep(2)
    

    
#print(df.iat[0,1])

# say "Hello World"
text = "练习结束，白白了。"
lang = "zh"
speech = Speech(text, lang)
speech.play()

# you can also apply audio effects while playing (using SoX)
# see http://sox.sourceforge.net/sox.html#EFFECTS for full effect documentation
#sox_effects = ("speed", "0.5")
#speech.play(sox_effects)