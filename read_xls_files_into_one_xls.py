# -*- coding: utf-8 -*-
"""
Created on Mon Aug  9 16:29:10 2021

@author: Jan
"""

import xlrd
import os
import xlsxwriter
from tkinter import filedialog
from tkinter import Tk


                        #'Tag'      , 'Header', ValueOffset , Column(-offset)

variables2BextractedbyOffsetToTag =  [
                        {'Tag':'Leeftijd'       ,'Header':'Leeftijd'    , 'Offset': 1,'Column':0, 'Condition':''         },
                        {'Tag':'Geslacht:'      ,'Header':'Geslacht'    , 'Offset': 1,'Column':1, 'Condition':''},
                        {'Tag':'Lengte:'        ,'Header':'Lengte'      , 'Offset': 1,'Column':2, 'Condition':''},
                        {'Tag':'Gewicht:'       ,'Header':'Gewicht'     , 'Offset': 1,'Column':3, 'Condition':''},
                        {'Tag':'Stappen'        ,'Header':'Stappen'     , 'Offset': 1,'Column':4, 'Condition':'IsFietser'},
                        {'Tag':'Stappen'        ,'Header':'StepDuration', 'Offset': 2,'Column':5, 'Condition':'IsFietser'},
                        {'Tag':'Beginbelasting' ,'Header':'BegBelast'   , 'Offset': 1,'Column':6, 'Condition':'IsFietser'},
                        {'Tag':'Stappen'        ,'Header':'Stappen'     , 'Offset': 1,'Column':7, 'Condition':'IsLoper'},
                        {'Tag':'Stappen'        ,'Header':'StepDuration', 'Offset': 2,'Column':8, 'Condition':'IsLoper'},                        
                        {'Tag':'Beginbelasting' ,'Header':'BegSnelheid' , 'Offset': 1,'Column':9, 'Condition':'IsLoper'},
                        {'Tag':'WattMAX'        ,'Header':'WattMax'     , 'Offset': 1,'Column':10, 'Condition':'IsFietser'},
                        {'Tag':'Watt(/kg)'      ,'Header':'Watt/kg'     , 'Offset':-1,'Column':11, 'Condition':'IsFietser'},
                        {'Tag':'2-mmol/l'       ,'Header':'Vel 2-mmol'  , 'Offset': 1,'Column':12, 'Condition':'IsLoper'},
                        {'Tag':'3-mmol/l'       ,'Header':'Vel 3-mmol'  , 'Offset': 1,'Column':13,'Condition':'IsLoper'},
                        {'Tag':'4-mmol/l'       ,'Header':'Vel 4-mmol'  , 'Offset': 1,'Column':14,'Condition':'IsLoper'},
                        {'Tag':'2-mmol/l'       ,'Header':'Watt 2-mmol' , 'Offset': 1,'Column':15,'Condition':'IsFietser'},
                        {'Tag':'3-mmol/l'       ,'Header':'Watt 3-mmol' , 'Offset': 1,'Column':16,'Condition':'IsFietser'},
                        {'Tag':'4-mmol/l'       ,'Header':'Watt 4-mmol' , 'Offset': 1,'Column':17,'Condition':'IsFietser'},
                        {'Tag':'2-mmol/l'       ,'Header':'HF 2-mmol'   , 'Offset': 2,'Column':18,'Condition':''},
                        {'Tag':'3-mmol/l'       ,'Header':'HF 3-mmol'   , 'Offset': 2,'Column':19,'Condition':''},
                        {'Tag':'4-mmol/l'       ,'Header':'HF 4-mmol'   , 'Offset': 2,'Column':20,'Condition':''},
                        {'Tag':'2-mmol/l'       ,'Header':'VO2 2-mmol'  , 'Offset': 3,'Column':21,'Condition':'V02recorded'},
                        {'Tag':'3-mmol/l'       ,'Header':'VO2 3-mmol'  , 'Offset': 3,'Column':22,'Condition':'V02recorded'},
                        {'Tag':'4-mmol/l'       ,'Header':'VO2 4-mmol'  , 'Offset': 3,'Column':23,'Condition':'V02recorded'}
                                      ]
#headers=['Naam','Leeftijd','Geslacht','Lengte','Gewicht','WattMax','Watt/kg','Int2','Int3','Int4','HF2','HF3','HF4','VO2-2','VO2-3','VO2-4']



dataOutDir='../DataOut'
outXls=dataOutDir+'/SPORTS_MET_DB.xls'

keepItShortToTest=0


# some helper functions 
def findCell(sh, searchedValue):
    for row in range(sh.nrows):
        for col in range(sh.ncols):
            myCell = sh.cell(row, col)
            if myCell.value == searchedValue:
                return row,col
    return -1,-1    


def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass
    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass
    return False


# main program
print('creating xls file ',outXls, 'and mark for all capilars which polygons it touches') 
outWorkbook = xlsxwriter.Workbook(outXls)
outSheet = outWorkbook.add_worksheet()
bold          = outWorkbook.add_format({'bold': True})
number_format = outWorkbook.add_format({'num_format': '#,##0.00'})

#write headers for the variables retrieved from the filename
outSheet.write(0,0, 'Name'    ,bold)
outSheet.write(0,1, 'PersonID',bold)
outSheet.write(0,2, 'Protocol',bold)
#write headers for the +- special variables
outSheet.write(0,3, 'Max MZ value')
outSheet.write(0,4, 'VO2-MAX')
offset=5 # first 'x' columns are occupied already
# write headers for the list of variables retrieved via an offset to a tag
i=0
for searchedVariable in variables2BextractedbyOffsetToTag:
    outSheet.write(0,offset+i,searchedVariable['Header'])
    i=i+1

useHardCodedPath=1
if useHardCodedPath :
  path = '../DataIn'
else:
  root = Tk()
  path = filedialog.askdirectory(initialdir = "../DataIn",title = "Select directory of wghich the files are to be loaded into the db")
  root.destroy()

# retrieve all filenames from the chosen directory
files = os.listdir(path)

# process each file (one after the other)
fileNr=0;
for fileIn in files:
    fileNr=fileNr+1
    if keepItShortToTest:
       if fileNr >= 50:
         break
    print('file nr:',fileNr,':',fileIn)
    
    # retrieve the information embedded in the filename
    variablesFromFilename=fileIn.split('_')
    
    nameFromFilename    =variablesFromFilename[0]
    personIDFromFileName=variablesFromFilename[1]
    protocolFromFileName=variablesFromFilename[3]
    
    outSheet.write(fileNr, 0, nameFromFilename    ,bold)
    outSheet.write(fileNr, 1, personIDFromFileName,bold)
    outSheet.write(fileNr, 2, protocolFromFileName,bold)
    
    conditionsFullfilled=['']
    
    if 'Fiets' in protocolFromFileName:
        sheetsToBeLookedAt=['INVOER-FIETS','Afdruk-FIETS Absoluut','Afdruk-FIETS']
        WattMax_Recorded=True
        conditionsFullfilled.append('IsFietser')
    elif 'Lpb' in protocolFromFileName:
        sheetsToBeLookedAt=['INVOER-LPB','Afdruk-LPB Absoluut','Afdruk-LPB']
        conditionsFullfilled.append('IsLoper')
        WattMax_Recorded=False
    else:
        sheetsToBeLookedAt=[]
        print ('protocol not recognised from filename')
    
    if 'VO2' in protocolFromFileName:
        conditionsFullfilled.append('V02recorded')

        
    if 'MZ' in protocolFromFileName:
        conditionsFullfilled.append('MZrecorded')
        
    # retrieve the info from the xls sheet     
    pathToFile=path+'/'+fileIn
    wb = xlrd.open_workbook(pathToFile)

    for searchedVariable in variables2BextractedbyOffsetToTag:
       for sh in wb.sheets():  
          if sh.name in sheetsToBeLookedAt:  
             if (searchedVariable['Condition'] in conditionsFullfilled):  
                 myRow,myCol=findCell(sh, searchedVariable['Tag'])
                 if myRow!=-1:
                    column=myCol+searchedVariable['Offset']
                    value= sh.cell_value(myRow,column)
                    print('found cell:',searchedVariable['Tag'],' with value:',value,' in sheet:',sh.name)
                    if is_number(value):
                        outSheet.write(fileNr, searchedVariable['Column']+offset,value,number_format )
                    else:    
                        outSheet.write(fileNr, searchedVariable['Column']+offset,value )
             else:
                    outSheet.write(fileNr, searchedVariable['Column']+offset,'na' )
    
    if ('MZrecorded' in conditionsFullfilled) :
        # fetch the max lactaat (max value underneath MZ for 'Melkzuur') from the appropiate sheets
        for sh in wb.sheets():  
           if sh.name in sheetsToBeLookedAt:  
              myRow,myCol=findCell(sh, 'MZ')
              if myRow!=-1:
                 maxMz=0 
                 for line in range(1,15):
                   value= sh.cell_value(myRow+line,myCol)
                   if is_number(value):
                      if value > maxMz:
                         maxMz=value
                 print('found cell:','MZ',' with max value:',maxMz,' in sheet:',sh.name)
                 columnForThisVariable=3
                 outSheet.write(fileNr, columnForThisVariable,maxMz)
    
              myRow,myCol=findCell(sh, 'La')
              if myRow!=-1:
                 maxMz=0 
                 for line in range(1,15):
                   value= sh.cell_value(myRow+line,myCol)
                   if is_number(value):
                      if value > maxMz:
                         maxMz=value
                 print('found cell:','La',' with max value:',maxMz,' in sheet:',sh.name)
                 columnForThisVariable=3
                 outSheet.write(fileNr, columnForThisVariable,maxMz)
    else:
       columnForThisVariable=3
       outSheet.write(fileNr, columnForThisVariable,'na')
       

    if ('V02recorded' in conditionsFullfilled):
        # do something simil
        print('VO2 recorded')
        for sh in wb.sheets():  
          if sh.name in sheetsToBeLookedAt:  
             myRow,myCol=findCell(sh, 'VO2-MAX')
             if myRow!=-1:
                value= sh.cell_value(myRow,myCol+1)
                print('found cell:','VO2-MAX',' with value:',value,' in sheet:',sh.name)
                columnForThisVariable=4
                outSheet.write(fileNr, columnForThisVariable,value,number_format )
    else:
       columnForThisVariable=4
       outSheet.write(fileNr, columnForThisVariable,'na')
       
    print('.....')
outWorkbook.close()  