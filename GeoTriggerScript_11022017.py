# importing the requests library
import requests
import json
import lxml
from xml.dom import minidom
from xml.etree import ElementTree
import xml.dom.minidom
import csv
import sqlite3
import xlrd
import xml.etree.ElementTree as ET
import os
from fnmatch import fnmatch
import datetime
from os.path import basename

##  Initial file would be here:
##  G:\Inform\Analytics\Prospect_research\Geotriggers\data
##  after adding zillow data, files will be saved here:
##  G:\Research\Geotriggers\


def find_file(folder):
    print "find_file function"    
    # <><><><><><><><><><>  Get Today's Date  <><><><><><><><><><>
    today = datetime.date.today()    
    today = today.strftime('%Y%m%d')
    
  
    print "<><><><><><><><><><>  Get Today's Files  <><><><><><><><><><>"
    folderContent = os.listdir(folder)
    eligibleFiles = []
    for i, file in enumerate(folderContent):
        
        if file.startswith(today) and file.endswith(".xlsx"):  # or -> if today in file
            
            myfile = file
            
            print myfile
            #eligibleFiles.append(file)
    #print(eligibleFiles)
    return myfile


def csv_from_excel(myxlsfile,input_xls, csvfilename,data):
    print "within the csv to excel function"
    print myxlsfile
    print input_xls
    print csvfilename
    print data
    
    
    wb = xlrd.open_workbook(myxlsfile)
    sh = wb.sheet_by_name('Geotriggers Report - Movers')
    
    my_csv_file = open(data, 'wb')
    wr = csv.writer(my_csv_file, quoting=csv.QUOTE_ALL)

    for rownum in xrange(sh.nrows):        
        wr.writerow(sh.row_values(rownum))

    my_csv_file.close()


def callZilliowAPI(fp, myPara1, myPara2):        
    # api-endpoint
    URL = "http://www.zillow.com/webservice/GetSearchResults.htm"         
    # defining a params dict for the parameters to be sent to the API
##    PARAMS = {'zws-id':"X1-ZWz1g1g4isdix7_4g4th",
##                'address': "214 West Bailey Road",
##                'citystatezip': "Naperville IL"
##                  }

     
    PARAMS = {'zws-id':"X1-ZWz1g1g4isdix7_4g4th",
              #'zws-id':"X1-ZWz1a4wwh7dnuz_aqadi",
              'address': myPara1,
              'citystatezip': myPara2 ,
              'rentzestimate': 'true'
            }


    #http://www.zillow.com/webservice/GetSearchResults.htm?zws-id=X1-ZWz1a4wwh7dnuz_aqadi&address=175 West 87th Street, Apartment 2A&citystatezip=New York%2CNY%2C10024&rentzestimate=true    
        # sending get request and saving the response as response object
    r = requests.get(url = URL, params = PARAMS)
    page = requests.request(method="get", url = URL, params = PARAMS)      
    print ">>>>>>>>>> content   >>>>>>>>>>>>>>>>>>>>>>>>"
    print r.content
    print type(r.content)
    root = ET.fromstring(r.content)
    print root          
    print "above is the root element"
    print root[0][0].text  #214 West Bailey Road
    print root[0][1].text  #Naperville IL
    print root[1][0].text #messages
    if "Error" in root[1][0].text:
        print "Error"
        myresult1 = 0
        myresult2 = " "
        myresult3 = 0
        myresult4 = 0
    else:
        print root[2][0][0][3][0].text  ## Amount
        myresult1= root[2][0][0][3][0].text
        print root[2][0][0][3][1].text  ## Last Upate
        myresult2= root[2][0][0][3][1].text 
        print root[2][0][0][3][3].text  ## Value Change       
        print root[2][0][0][3][4][0].text  ## Low Amount
        myresult3 = root[2][0][0][3][4][0].text    
        print root[2][0][0][3][4][1].text  ## High Amount
        myresult4 = root[2][0][0][3][4][1].text
    return (myresult1, myresult2, myresult3, myresult4)

def open_csvfile(fp, mycsvfilename, mycsvfile,path):
    print "my output file:" + str(fp)
    print "mycsvfilename:" + mycsvfilename
    print "mycsvfile: " + mycsvfile
    print "mypath: " + path
    print "within the open csv system"
    ## added file name 09292016  
    mycsvreader =csv.reader(open(mycsvfile,"rb"))
##    mycsvreader = open(mycsvfile, "r")
##    mycsvreader.seek(0)
    row1 = next(mycsvreader)
    row1.append("Zestimate")
    row1.append("Date Updated")
    row1.append("Zestimate Low")
    row1.append("Zestimate High")
    a = csv.writer(fp, delimiter =',')
    print row1
    a.writerows([row1])    
    for row in mycsvreader:
        print row
        myaddress = row[4]
        print myaddress
        myPara1 = myaddress        
        mycity = row[5]
        print mycity
        mystate = row[6]        
        print mystate
        myzip = row[7]
        print myzip        
        myPara2 = mycity + " " + mystate + " " + myzip
        print myPara2        
        myvalue, myupdatedate, mylowamnt, myhighamnt = callZilliowAPI(fp, myPara1, myPara2)
        print " >>>>>>>>>>>>>> I am now back to the open csvfile function >>>>>>>>>>>>>>>>>>>> "
        print myvalue
        print myupdatedate
        print mylowamnt
        print myhighamnt        
        row.append(myvalue)
        row.append(myupdatedate)
        row.append(mylowamnt)
        row.append(myhighamnt)
        print row               
        a.writerows([row])
    return 

def main():

    #root = r"G:\SHARED\Analytics\Communications\email_data\All Email Data"  ## this is data from sphere data from Nov.2014 - March2016
    #root = r"G:\Inform\Analytics\ZillowAPI"
    root_raw = r"G:\Inform\Analytics\Prospect_research\Geotriggers\data"
    #input_xls = "20171020_Geotriggers.xlsx"
    root_ouput = r"G:\Research\Geotriggers"
    #root_ouput = r"G:\Inform\Analytics\Xiaohong\GeoTrigger"
    myinput_xls = find_file(root_raw)    
    myxlsfile = os.path.join(root_raw, myinput_xls)
    print myxlsfile

    base=os.path.splitext(myinput_xls)[0]
    print " print out the base "
    print base
    csvfilename = base + ".csv"
    print " ------------------------------   remove extension and add csv  ------------------------------ "
    print csvfilename
    
    #csvfilename  = "20171020_Geotriggers.csv"
    pattern = "*.csv"
    n = 0
    
    csvile = os.path.join(root_raw, csvfilename)
    print csvile
    
    csv_from_excel(myxlsfile,myinput_xls, csvfilename, csvile)
    ouput_csv = os.path.join(root_ouput, csvfilename)


    


    
    #data = r"G:\Inform\Analytics\ZillowAPI\20171012_GeotriggersMoreValues_Update.csv"
    
    print " -----------------------------  DONE with Converting ------------------------------ "
    csvfile_name = []
    try:
        with open(ouput_csv, 'ab') as fp:
            for path, subdirs, files in os.walk(root_raw):
                for name in files:                                        
                    ###### Csv Files #################
                    if fnmatch(name, pattern):
                        mycsvfile = os.path.join(path, name)
                        mycsvfilename = name               
                        
                        if mycsvfilename == csvfilename:
                            print ">>>>> got the file name >>>>>>>>>>>>>>"
                            print mycsvfilename
                            open_csvfile(fp, mycsvfilename, mycsvfile,path)  ## added file name 09292016                                                                                             
                        else:                            
                            print "bad files" + mycsvfilename 
        
##        mycsvreader=csv.reader(open(mycsvfile,"rb"))    
##        row1 = next(mycsvreader)
##        if row1 == ['First Name', 'Last Name', 'Current Email Address', 'Supporter ID']:        
##            for row in mycsvreader:
##                row.append(mycsvfile)
##                row.append(mycsvfilename)
##                row.pop(0)
##                row.pop(0)
##                row.insert(0, u' ')
##                row.insert(3,u' ')
##                row.insert(4,u' ')
##                row.insert(5,u'E-mail Sent Successfully')                
##                mymodrow = []       
##                mylist =[row[0],row[2], row[1], row[3], row[4], row[5], row[6], row[7]]     ## added file name        
##                a = csv.writer(fp, delimiter =',')
##                if mycount > 0:
##                    a.writerows([mylist])
##                if mycount == 1:
##                    print "SPE_CSV"+"|"+ mycsvfile+"|"+ str(mylist)                                    
##                mycount += 1    

    except ZeroDivisionError as e:
        print e        
        sys.exit(0)
    except IOError as e:
        print e

if __name__ == "__main__":
    main()
