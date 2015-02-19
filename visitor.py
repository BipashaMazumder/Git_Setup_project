import csv, MySQLdb, traceback
from collections import defaultdict
import xlrd as xl
from datetime import datetime
import sys

gInputFileNameIncludingPath = "C:\Users\AKDANI\Desktop\\pytn\Retailer Enquiry vs Website Visitors vs WOI vs Hitwise Share.xls"
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
    
    if( gSheetNameList[0]=="Sheet1" ):
        sheet=gSheetNameList[0]
        print sheet
        worksheet = workbook.sheet_by_name(sheet)
        lNumRows = worksheet.nrows - 26
        lNumCells = worksheet.ncols - 1
        
        print lNumCells
        lCurrRow = 5
        daterow=4
        while lCurrRow <= lNumRows:
            lCurrCell = 2
            datecol=2
            
            while lCurrCell <= lNumCells:
                daterow1=4
                datecol1=2
                if lCurrRow < 11  :
                    lDate = datetime(*xl.xldate_as_tuple(worksheet.cell_value(daterow1, datecol1), workbook.datemode))
                    lDay_string = lDate.strftime('%Y-%m-%d')
                if lCurrRow == 11  :
                    lCurrRow = 14
                    daterow=13
                    datecol=2
                if lCurrRow == 20 :
                    lCurrRow = 23
                    daterow=22
                    datecol=2
                if (worksheet.cell_value(daterow, datecol)):
                    lDate = datetime(*xl.xldate_as_tuple(worksheet.cell_value(daterow, datecol), workbook.datemode))
                    lDay_string = lDate.strftime('%Y-%m-%d')
                else:
                    break
                if(worksheet.cell_value(lCurrRow,1)):
                    f=str(worksheet.cell_value(lCurrRow,1))
                    Id = f.strip().replace(" - "," ")
                    var = Id.split()
                    var2= Id.replace(var[0] , "")
                    a1=var2.replace("  ", " ")
                    var3 =a1.strip().replace(" ","_")
                   
                    luni= var3
                    
                lUniqueId = luni
                
                lNodeName = lUniqueId
                
                lCategory = "Client Data:Sales:" + lUniqueId
                
                lCellValue = worksheet.cell_value(int(lCurrRow), int(lCurrCell))
                if f == "2012 Web - FaR":
                    print lCellValue, lCurrRow, lCurrCell, lDay_string, daterow
                
                if lCellValue == "":
                    lCellValue = 0.0
            
                lComment = ""
                sql = 'INSERT INTO visitor(unique_id, date, value, category, node_name, comments) VALUE ("'\
                +lUniqueId+'","'+ lDay_string+'", "'+str(lCellValue)+'","'+lCategory+'","'+lNodeName +'","'+lComment+'")'
           
                gCursor.execute(sql)
                datecol +=1
                lCurrCell +=1 
                print lCellValue,lDay_string,lCurrRow
            lCurrRow += 1
            
        
        gMydb.commit()    
main()             
                 