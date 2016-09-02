#
# Reads player information out of an excel document, 
# calculates the sizes of the various populations, and then writes these into the excel document.
#

import openpyxl

wb = openpyxl.load_workbook('Analysis.xlsx')
ws = wb.get_sheet_by_name("rawdata")

playerlist=[]

for i in range (1, 27):  
    templist=[]
    current_row = 3
    while(True):
        val = ws.cell(column=i, row = current_row).value
        if val != None:
            templist.append(val)
            current_row+=1
        else:
            playerlist.append(templist)
            break

newplayercount=[]
activeplayercount=[]
returningplayercount=[]
#inactiveplayers=[]
masterlist=[]

for i in range (0, len(playerlist)):
    newplayers=0
    activeplayers=0
    returningplayers=0
    for j in range(0, len(playerlist[i])):
        #new player check
        if playerlist[i][j] not in masterlist:
            newplayers+=1
            masterlist.append(playerlist[i][j])
        else:
            #check for returning players (last game was > 6 months ago)
            if i>2 and playerlist[i][j] not in playerlist[i-1] and playerlist[i][j] not in playerlist[i-2]:
                returningplayers+=1
            else:
                activeplayers+=1
                
    newplayercount.append(newplayers)
    activeplayercount.append(activeplayers)
    returningplayercount.append(returningplayers)
    

ws = wb.get_sheet_by_name("playergroups")
#Set headers
_=ws.cell(column=1, row = 2, value = "New")
_=ws.cell(column=1, row = 3, value = "Active")
_=ws.cell(column=1, row = 4, value = "Returning")
current_column = 2
for i in range (2010, 2017):  
    for j in  range (1,5):
        _=ws.cell(column=current_column, row = 1, value = str(i) + "Q" + str(j))
        current_column+=1
        
#populate!
current_column = 2
for number in newplayercount:
     _=ws.cell(column=current_column, row = 2, value = number)
     current_column+=1

current_column = 2
for number in activeplayercount:
     _=ws.cell(column=current_column, row = 3, value = number)
     current_column+=1
     
current_column = 2
for number in returningplayercount:
     _=ws.cell(column=current_column, row = 4, value = number)
     current_column+=1
     
wb.save(filename="Analysis.xlsx")        
