#! python3

#Decoration: Line Spacing
print()

##Hi Message Section##

#Import 'username' for Hi Message
import getpass
username=getpass.getuser()

def getHiMsg(name):
    return 'Hi '+name+', this applications helps you to download production volumes of OCT main lines, just by entring the from-to dates.'

#Hi Message Loop
Kusern = ['barghoma','sabriamr','elzog','eletradh','sakroma']
KUserN = ['Barghoma','Sabriamr','Elzog','Eletradh','Sakroma']
KName = ['Barghash','Sabri','Zyad','Adham','Sakr']
foundUser = False
for usern,UserN,Name in zip(Kusern,KUserN,KName):
    if username == usern or username == UserN:
        print(getHiMsg(Name))
        foundUser = True
        break
        
if foundUser != True:
    Name = 'OCTian'
    print(getHiMsg(Name))

#Start of Repetition Loop   
repeat = True
while (repeat):
    
    #Decoration: Line Spacing
    print()
    
##Inputs Section##
    
    print('Please enter "from Date" in form of "d/m/yyyy":')
    Date1=input()

    #Decoration: Line Spacing
    print()

    print('Please enter "to Date" in form of "d/m/yyyy":')
    Date2=input()

##Excecution Section##

    #Open Browser
    from selenium import webdriver
    import os
    driver=webdriver.Chrome('chromedriver.exe')

    #Create Target Folder
    import os
    import datetime
    from selenium.webdriver.chrome.options import Options
    currentDT = datetime.datetime.now()
    today=currentDT.strftime("%d-%m-%Y %I-%M %p")
    DirName = str(today)
    os.makedirs('PPM Export '+DirName)
    CWD = str(os.getcwd())
    download_dir = CWD+'\PPM Export '+DirName
    chrome_options = webdriver.ChromeOptions()
    preferences = {"download.default_directory": download_dir ,
                   "directory_upgrade": True,
                   "safebrowsing.enabled": True }
    chrome_options.add_experimental_option("prefs", preferences)
    driver = webdriver.Chrome(options=chrome_options,executable_path='chromedriver.exe')

    #Sign Into PPM
    driver.get('http://appsoct/PPM')
    UsernameElem=driver.find_element_by_css_selector('#txtUserName')
    UsernameElem.send_keys('Barghoma')
    PasswordElem=driver.find_element_by_css_selector('#txtPassword')
    PasswordElem.send_keys('123')
    OKButtonElem=driver.find_element_by_css_selector('#btnLogin')
    OKButtonElem.click()

    #Imports
    import xlwings as xw
    import os
    os.chdir('PPM Export '+DirName)
        
    #Create Consolidate.xlsx File
    con = xw.Book()
    conws = con.sheets('Sheet1')
    con.save('Consolidated '+str(today)+'.xlsx')
    
    #Browser Loop
    LineID = ['0b91d766-ce63-4a46-8160-fd43aaf32484','0d80d162-49c9-441b-83bd-5bdcc5b4008a','92f449f6-b103-4cdd-919d-5bddd10a46cf','5698ef75-ff27-4e09-9da1-1ada2c2b6826','372f9a95-1c38-4adf-b4ae-988da45ac73a','65f35e66-d467-4af8-a9d7-a86475adbd5f']
    TabNr = range(1,7)
    for ID,tab in zip(LineID,TabNr):

        #Open Empty Tabs
        driver.execute_script("window.open('');")

        #Get Report
        driver.switch_to.window(driver.window_handles[tab])
        driver.get('http://appsoct/PPM/Reports/ppmreport.aspx?sdate='+Date1+'&edate='+Date2+'&selteam=&selshift=00000000-0000-0000-0000-000000000000&shiftname=00000000-0000-0000-0000-000000000000&speriod=P3-Week1-Day1  ('+Date1+')&eperiod=P4-Week1-Day3  ('+Date2+')&rpt=ProductionDetailReport&type=Period&filter=TRS&lineguid='+str(ID))

        #Download Report
        driver.switch_to.window(driver.window_handles[tab])
        ExcelElem=driver.find_element_by_css_selector('#ctl00_ReportContent_ReportViewer1_ctl01_ctl05_ctl00 > option:nth-child(6)')
        ExcelElem.click()
        ExportElem=driver.find_element_by_css_selector('#ctl00_ReportContent_ReportViewer1_ctl01_ctl05_ctl01')
        ExportElem.click()

    #Excel Loop
    Reports = ['ProductionReportDetail_OCT','ProductionReportDetail_OCT (1)','ProductionReportDetail_OCT (2)','ProductionReportDetail_OCT (3)','ProductionReportDetail_OCT (4)','ProductionReportDetail_OCT (5)']
    LinesNames = ['Flutes 1','Flutes 2','LML','Twix','ML2','Jewels']
    ConCol = range(2,8)
    for report,ln,cc in zip(Reports,LinesNames,ConCol):    

        #Open Production Report
        bookName = report+'.xls'
        sheetName = 'ProductionReportDetail_OCT'
        wb = xw.Book(bookName)
        ws = wb.sheets[sheetName]

        #Specify "Total Volume" Column
        myCell = ws.api.UsedRange.Find('Total Volume')
        C = myCell.Column

        #Insert Sum Column
        xw.Range((1,C),(1250,C)).insert()
        SumTitle = ws.range((1,C))
        SumTitle.value = str(ln)

        #Write Sum Formula in First Cell
        SumCell = ws.range((2,C))
        B4SumCell = ws.range(2,C-1)
        SumCell.value = '=Sum(J2:'+B4SumCell.get_address(False,False)+')'

        #Autofill
        from xlwings.constants import AutoFillType
        LastRow = ws.range(1,1).end('down').row
        SumCell.api.AutoFill(ws.range((2,C),(LastRow,C)).api,0)

        #Copy & Paste Sum Column
        ws.range((1,C),(LastRow,C)).copy()
        conws.range(1,cc).paste('values')

    # Time Stamp Column
    ws.range((1,9),(LastRow,9)).copy()
    conws.range(1,1).paste('values')

##Quitting Program Section##

    #Quit Excel Reports
    for report in Reports:
        bookName = report+'.xls'
        sheetName = 'ProductionReportDetail_OCT'
        wb = xw.Book(bookName)
        ws = wb.sheets[sheetName]
        wb.save()
        wb.close()
        
    #Quit Browser    
    driver.quit()

    #Decoration: Line Spacing
    print()

    #Repeat Message 
    print('Press "r" to repeat, or press any key to exit')
    userinput=input()
    if userinput == 'r':
        repeat = True
    else:
        repeat = False

#End!
