import csv, MySQLdb, traceback
import MySQLdb
from collections import defaultdict
import xlrd as xl
from datetime import datetime
import sys

gInputFileNameIncludingPath = "C:\\Users\\AKDANI\\Desktop\\pytn\\Subaru\\Subaru Monthly Channel H1_H2 Report December 2012.xls"
gCursor = ""
gMydb = ""

#=======================================================================================================

def sqldbConnection():
    print "SqlDBConnection..."
    global gCursor, gMydb
    try:
        gMydb = MySQLdb.connect(host='127.0.0.1', user='root', passwd='proloy', db='proloy')
        gCursor = gMydb.cursor()
        print "SqlDBConnection Successfully done.."
        
    except Exception , e: 
        print "Database is not connected",  e[0], e.args[1]
        exit();
#=======================================================================================================

def main():
    global gInputFileNameIncludingPath, gCursor, gMydb ,datecol
    
    sqldbConnection()
    workbook = xl.open_workbook(gInputFileNameIncludingPath)
    gSheetNameList = workbook.sheet_names()
        
    if( gSheetNameList[15]=="Paid Search" ):
        sheet=gSheetNameList[15]
        print sheet 
        worksheet = workbook.sheet_by_name(sheet)
        lNumRows = worksheet.nrows - 2
        lNumCells = 13
        
        lCurrRow = 9
        
        while lCurrRow <= lNumRows:
            lCurrcell = 1
            
            
              
            if len(str(worksheet.cell_value(lCurrRow,1))) == 0 and len(str(worksheet.cell_value(lCurrRow+2,1))) == 0 :
                lCurrRow +=4
            
            print lCurrRow,lCurrcell
            
            if lCurrRow == 64:
                lCurrRow = 69
                
            if lCurrRow == 71:
                lCurrRow = 74
                

            while lCurrcell < lNumCells :
                
                
               if len(str(worksheet.cell_value(lCurrRow-3,1))) > 0 and len(str(worksheet.cell_value(lCurrRow-2,1))) ==0 :
                   lDate = datetime(*xl.xldate_as_tuple(worksheet.cell_value(lCurrRow-3,1), workbook.datemode))
                   lDay_string = lDate.strftime('%Y-%m-%d')
               
               
               if len(str(worksheet.cell_value(lCurrRow ,1))) == 0  and len(str(worksheet.cell_value(lCurrRow ,lCurrcell+4))) == 0 :
                   lCurrRow +=3
                   break
               
               lCampaign = str(worksheet.cell_value(8,lCurrcell+1))
                
               
               
               
               
               
               if str(worksheet.cell_value(lCurrRow,1)).find("Total")>=0:
                     break
                   
               rowstring = worksheet.cell_value(lCurrRow,1)
                   
                 
               lCurrcell += 1
               
               if worksheet.cell_value(7,lCurrcell) :
                   target=str(worksheet.cell_value(7,lCurrcell))
               ltarget= target
               print ltarget
               
                
               lCellValue = worksheet.cell_value(lCurrRow,lCurrcell)
               if lCellValue == "" or lCellValue == "NA":
                     lCellValue = 0.0
                     
               
               lUniqueId = sheet.capitalize().strip().replace(" ","_") +"_" +rowstring.strip().replace(" ","_")+\
                    "_"+ ltarget.strip().replace(" " , "_") + "_" + lCampaign.strip().replace(" ","_") 
               
               lNodeName = rowstring.strip().replace(" ","_")+ "_" +ltarget.strip().replace(" " , "_")+ "_" + lCampaign.strip().replace(" ","_")
               
               lCategory = "Client Data:Sales:" +sheet.capitalize() + ":" + rowstring.strip()+ ":" + ltarget + ":" + lCampaign
               
               lComment = ""
               lSql = 'INSERT INTO paid(unique_id, date, value, category, node_name, comments) VALUE ("'\
                      +lUniqueId+'","'+ lDay_string +'", "'+str(lCellValue)+'","'+lCategory+'","'+lNodeName +'","'+lComment+'")'
               gCursor.execute(lSql)
               
               
               
               
            lCurrRow += 1
        
        print "....................data inserted successfully.................."
        gMydb.commit() 
                
                
        
        
main()  
        
