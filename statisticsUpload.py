# The key of the spreadsheet, e.g. 0BmgG6nO_6dprdS1MN3d3MkdPa142WFRrdnRRUWl1UFE
spreadsheetKey = '16edX0XpVyzneq0vVrHjyaPmqYmrv4zAcm76sMW4sjm0' 

# When you open the statistics for a deck for which no worksheet has been created in your spreadsheet, create one?
# This will only be done, when the worksheet has the exact name of a deck, 
# e.g. "All::Japanese::Vocabulary", if your deck is All -> Japanese -> Vocabulary in the deckbrowser
createNewWorkSheets = True 

# Defines the FIRST FIELD
# Each day results in a new row. 
startRow = 1
startColumn = 1

# The order of the fields. The first field will get placed into the first column, and so on. 
# There are more possible fields.
# To know which are available, scroll down and look for self.data[XXX]. The XXX can be copied into here:
order = [
    "totalCardsStudied",
    "timeSpentStudying",
    "matureCardsStudied",
    "percentCorrectAllCards",
    "percentCorrectMatureCards",
    "totalMatureCards",
    "totalYoungLearnCards",
    "totalUnseenCards",
    "totalSuspendedCards",
    "totalDueCards"
]

# If dayInFirstRow is set True:
    # >Necessary< if you want to have maximal one row per day
    # The current DAY is calculated as follows
    #       from aqt import mw
    #       import datetime
    #       currentDay = (datetime.datetime.today() - datetime.datetime.fromtimestamp(mw.col.crt)).days
    # If you want to know the DAY beforehand, open the Anki console with ctrl+shift+: (double-colon) and enter these 3 lines and press ctrl+enter
    # If you have a table already with entries for 100 days, then you *must* calculate DAY-100 and write that value into the FIRST FIELD
# If dayInFirstRow is set False:
    # append one row to the document and use this for new values
dayInFirstRow = True

pathToPython = 'C:\Python27\Lib'

### FINISHED !!! ###














import os,sys,inspect
currentdir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
sys.path.insert(0,os.path.join(currentdir,'statisticsUpload')) 
sys.path.insert(0,pathToPython) 

import gflags
import requests
import gspread
import json
import os.path

client_id = "944084996494-5t61b1d6sftt6ue21o73s9kqaa0imte6.apps.googleusercontent.com"
client_secret = "x4tG2GVWIjg7DYfibyqAng3o"
scope = ['https://spreadsheets.google.com/feeds']

def setup():
    from requests_oauthlib import OAuth2Session
    if not os.path.isfile('creds.data'):    
        oauth = OAuth2Session(client_id,scope=scope,
                redirect_uri='urn:ietf:wg:oauth:2.0:oob')
        authorization_url, state = oauth.authorization_url(
            'https://accounts.google.com/o/oauth2/auth',
            access_type="offline",
            approval_prompt="force")
        import webbrowser
        webbrowser.open(authorization_url)
        from PyQt4 import QtGui
        from aqt import mw
        authorization_response, ok = QtGui.QInputDialog.getText(mw, 'Google Spreadsheet Authorization', 'Enter the full Key')
        if not authorization_response or not ok:
            raise Exception('Need Authorization!')
        token = oauth.fetch_token(
            token_url='https://accounts.google.com/o/oauth2/token',
            code=authorization_response,
            client_id=client_id,
            client_secret=client_secret)
        with open('creds.data','w') as f:
            json.dump(token,f)
        return oauth,token
    else:
        with open('creds.data','r') as f:
            token = json.load(f)
        oauth = OAuth2Session(client_id,token=token)
        token = oauth.refresh_token(
            client_id=client_id,
            client_secret=client_secret,
            token_url='https://accounts.google.com/o/oauth2/token')
        with open('creds.data','w') as f:
            json.dump(token,f)
        return oauth,token
    

def authenticate_google_docs():
    oauth,cdata = setup()
    data = {
        'refresh_token' : cdata['refresh_token'],
        'client_id' : client_id,
        'client_secret' : client_secret,
        'grant_type' : 'refresh_token',
    }
    
    gc = gspread.Client(auth=cdata)
    gc.login()
    return gc

    
def upload(title):
    gc = authenticate_google_docs()
    sh = gc.open_by_key(spreadsheetKey)
    if title not in [ws.title for ws in sh.worksheets()]:
        if not createNewWorkSheets:
            return None
        else:
            return sh.add_worksheet(title=title,rows="365",cols=len(order))
    else:
        return sh.worksheet(title)

from anki.stats import CollectionStats
class ChangedCollectionStats(CollectionStats):
    def __init__(self, col):
        CollectionStats.__init__(self,col)
        
    def report(self, type=0):
        self.type = type
        self.data = {}
        try:
            title = self.col.decks.get(self.col.decks.active()[0])[u'name']
            ws = upload(title)
            if ws is not None:
                lim = self._revlogLimit()
                if lim:
                    lim = " and " + lim
                a,b,c,d,e,f,g = self.col.db.first("""
                    select count(), sum(time)/1000,
                    sum(case when ease = 1 then 1 else 0 end), /* failed */
                    sum(case when type = 0 then 1 else 0 end), /* learning */
                    sum(case when type = 1 then 1 else 0 end), /* review */
                    sum(case when type = 2 then 1 else 0 end), /* relearn */
                    sum(case when type = 3 then 1 else 0 end) /* filter */
                    from revlog where id > ? """+lim, (self.col.sched.dayCutoff-86400)*1000)
                self.data["totalCardsStudied"] = a or 0
                self.data["timeSpentStudying"] = b or 0
                self.data["againCount"] = c or 0
                self.data["learn"] = d or 0
                self.data["review"] = e or 0
                self.data["relearn"] = f or 0
                self.data["filtered"] = g or 0
                if self.data["totalCardsStudied"] > 0:
                    self.data["percentCorrectAllCards"] = "{0:.2f}".format((1-self.data["againCount"]/float(self.data["totalCardsStudied"]))*100)
                else:
                    self.data["percentCorrectAllCards"] = 0
                    
                mcnt, msum = self.col.db.first("""
                    select count(), sum(case when ease = 1 then 0 else 1 end) from revlog
                    where lastIvl >= 21 and id > ?"""+lim, (self.col.sched.dayCutoff-86400)*1000)
                self.data["matureCardsStudied"] = msum or 0
                if mcnt > 0:
                    self.data["percentCorrectMatureCards"] = "{0:.2f}".format(self.data["matureCardsStudied"] / float(mcnt) * 100)
                else:
                    self.data["percentCorrectMatureCards"] = 0
                
                a,b,c,d = self.col.db.first("""
                    select
                    sum(case when queue=2 and ivl >= 21 then 1 else 0 end), -- mtr
                    sum(case when queue in (1,3) or (queue=2 and ivl < 21) then 1 else 0 end), -- yng/lrn
                    sum(case when queue=0 then 1 else 0 end), -- new
                    sum(case when queue<0 then 1 else 0 end) -- susp
                    from cards where did in %s""" % self._limit())
                self.data["totalMatureCards"] = a or 0
                self.data["totalYoungLearnCards"] = b or 0
                self.data["totalUnseenCards"] = c or 0
                self.data["totalSuspendedCards"] = d or 0
                for deck in self.col.sched.deckDueList():
                    if deck[0] == title:
                        self.data["totalDueCards"] = deck[2]
                        break
                s = ''
                for key,value in self.data.items():
                    s += '{0}:{1}\n'.format(key,value)
                from aqt import mw
                import datetime
                currentDay = (datetime.datetime.today() - datetime.datetime.fromtimestamp(mw.col.crt)).days
                from PyQt4 import QtGui
                QtGui.QMessageBox.information(mw,'Statistics Upload', s)
                
                
                col = startColumn
                if not dayInFirstRow:
                    row = ws.row_count + 1
                else:
                    col = col + 1
                    if ws.cell(startRow,startColumn).value == "":
                        row = startRow
                    else:
                        row = currentDay - int(ws.cell(startRow,startColumn).value) + 1
                    ws.update_cell(row,startColumn,str(currentDay))
                if row > ws.row_count:
                    ws.resize(rows = row)
                    
                cell_list = []
                for field in order:
                    cell = self.cell(row, col)
                    cell.value = self.data[field]
                    cell_list.append(cell)
                    col = col + 1
                ws.update_cells(cell_list)
        except:
            pass
        return CollectionStats.report(self,type)


def stats(self):
    return ChangedCollectionStats(self)
    
from anki.collection import _Collection
_Collection.stats = stats