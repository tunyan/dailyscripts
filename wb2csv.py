from __future__ import print_function, division
import csv
import sys
from xlrd import open_workbook
import os, shutil
from __builtin__ import file
from shutil import copy
from vobject.icalendar import PRODID

"""
Script to extract named sheets from a workbook into csv

*** salesfile is passed from command line passed as -w salesfile  
*** Sheets are converted to csv in data/ folder
"""
# Command line Argument Handling
try:
    import argparse
    parser = argparse.ArgumentParser(description='Script for creating csv files from xls file')
    parser.add_argument('-m','--dictfile', help='e.g -m fido_dict.csv', required=True)
    parser.add_argument('-w','--salesfile', help='e.g -w salesfile.csv', required=True)
    parser.add_argument('-d','--date', help='e.g -d 2017-03-09', required=True)
    args = vars(parser.parse_args())
except ImportError:
    parser = None

if not os.path.exists('./data'):
    os.makedirs('./data')

if not os.path.exists('./OUT'):
    os.makedirs('./OUT')

    
DICTFILE = open(args['dictfile'], 'rt')
WBFILE = args['salesfile']
DATEINVOICE = args['date']
ERRORFILE = 'OUT/errfile' + DATEINVOICE + '.csv'


errfile = open(ERRORFILE, 'w')
DATAFOLDER = 'data'
salesid = {}
custid = {}

# delete all files in folder
def delfiles(folder):
    for the_file in os.listdir(folder):
        file_path = os.path.join(folder, the_file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as e:
            print(e)

def customerqc(name,type):
    """
     DO quality control on customer name to reduce rejections
    """
    custname = name.strip()
    if type == 'CUSTOMER':
        
        
        custname = custname.replace('Jesus Love','Jesus-Love')
        custname = custname.replace('Roland','Rowland')
        custname = custname.replace('Christain','Christian')
        custname = custname.replace('Pishoh Gole','Pishon Gole')
        custname = custname.replace('Stella Amaran','Stella Amara')
        custname = custname.replace('Omorome','Omoreme')    
        custname = custname.replace('Egerekumo','Ederekumo')
        custname = custname.replace('Mathew','Matthew')
        custname = custname.replace('Mercy Ndubuisi','Mercy Ndubusi')        
        custname = custname.replace('Adebayo Fumi','Adebayo Funmi')
        
      
        custname = custname.replace('Okafor Priscilla','Priscilla Okafor')
        custname = custname.replace('Nigerian Neavy','Nig Navy')
        custname = custname.replace('Nigerian Navy','Nig Navy')
        custname = custname.replace('Doris Ogede','Ogede Doris')
        custname = custname.replace('Emeka Okolo','Okolo Emeka')
        custname = custname.replace('Sunday David','Sunny David')
        custname = custname.replace('Olayode Ujro','Olayode Ujiro')
        custname = custname.replace('Ganiyu Ayo','Ganiyu Motunrayo')
        custname = custname.replace('Chima Customer','Chima-Customer')
        custname = custname.replace('Ayodele Franca','Franca Ayodele')
        custname = custname.replace('Godspower Customer','Godspower-Customer')
        custname = custname.replace('New Integrated Service','New Integrated Services')
        custname = custname.replace('Florence','Flourence')
        custname = custname.replace('Seriyai Sam','Serieya Sam')
        custname = custname.replace('Obos Fransica','Obos Francisca')
        custname = custname.replace('Matthew Onyedikachi','Onyedikachi Matthew')
        custname = custname.replace('Godwin O','Godwin Orieneme')
        custname = custname.replace('Magdalene','Magdaline')
        custname = custname.replace('Toma','Kekemeke Toma')
        custname = custname.replace('Joan','Joan-Customer')
        custname = custname.replace('Jenakumo Jenifer','Jennifer')
        custname = custname.replace('Ayerigha','Ayengha')
        custname = custname.replace('Azode Chidebere','Azode Chidiebere')
        custname = custname.replace('Ukutta Ella','Ella Ukuta')
        custname = custname.replace('Ogechi','Ogechukwu')
        
        
        custname = custname.replace('Powiede Alokpa','Poweide Alokpa')
        custname = custname.replace('Pricilla Okafor','Priscilla Okafor')
        custname = custname.replace('Bennedicta','Benedicta')
        custname = custname.replace('Ifeiyenwa Nweke','Nweke Ifeanyiwa')
        custname = custname.replace('Samsom Freedom','Samson Freedom')
        custname = custname.replace('Omokuro','Omukoro')
        custname = custname.replace('Rhoda Charles','Charles Rhoda')
        custname = custname.replace('Oyeintaro Timipa','Timipa Oyintari')
        custname = custname.replace('Luck Water','Lucky Waters')
        custname = custname.replace('Benjamin Karibi','Benjamin Kariebi')
        custname = custname.replace('Akinola Damilola','Akinola Dami')
        custname = custname.replace('Akan Augustin','Akan')
        custname = custname.replace('Grace Douglas','Douglas Grace')
        custname = custname.replace('GodsGift','GodGift')
        custname = custname.replace('Jesus love','Jesus-Love')
        custname = custname.replace('Oguru John','Oguru')
        custname = custname.replace('Selepre','seleipre Ero')
        custname = custname.replace('preye mathew','Preye Matthew')
        custname = custname.replace('Uche Pick up','Uche Pick-up')
        custname = custname.replace('Omadudu Faith','Faith Omadudu')
        custname = custname.replace('Suleman','Suleiman')
        custname = custname.replace('Sam Seriyai','Serieya Sam')
        custname = custname.replace('Oghenefejiro Monday','Oghenefein Monday')
        custname = custname.replace('Eduno Benjamin','Benjamin Eduno')
        custname = custname.replace('Ikechukwu Ossai','IK Ossai')
        #custname = custname.replace('Oluchukwu','Oluchi')
        custname = custname.replace('Capt Timi','Captain Timi')
        custname = custname.replace('Amax Benita','Benita Max')
        
        
        custname = custname.replace('Azibapu Success','Success Azibapu')
        custname = custname.replace('Odi Ada','Ochi Ada')
        custname = custname.replace('sunday david','Sunny David')
        custname = custname.replace('EMMANUEL CHUKWUKA','Emmanuel C')
        custname = custname.replace('Ariye Suware','Ariye')
        custname = custname.replace('Alex Oleya','Alex Oliya Customer')
        custname = custname.replace('Innocent Oji','Innocent Orji')
        custname = custname.replace('Onoro Ogheneovo','Ogheneovo Onoro')
        custname = custname.replace('Praise Kiani','Praise Kelani')
        custname = custname.replace('Nozi','Ngozi')
        custname = custname.replace('Solomon Macaulay Ebi','Solomon Macauley Ebi')
        custname = custname.replace('Gbeinz Ebinimi','Gbeinzi Ebinimi')
        custname = custname.replace('Serieyai Sam','Serieya Sam')
        custname = custname.replace('Kessrich','Kess Rich')
        custname = custname.replace('Samuel Angalabiri','Samuel Agalabiri')
        custname = custname.replace('Victory Ikeinga','Victoria Ikenga')
        custname = custname.replace('Gloria Esseh','Gloria Eseh')
        custname = custname.replace('Ogrega Roseline','Rose Ogriga')
        custname = custname.replace('Priscillia Okafor','Priscilla Okafor')
        custname = custname.replace('Godswill Achibong','Godswill Archibong')
        custname = custname.replace('Faith Hillary','Faith Hilary')
        custname = custname.replace('Essiet Elizabeth','Elizabeth Essiet')
        custname = custname.replace('Stanley Ukodhe','Stanley Ukedhe')
        custname = custname.replace('Olatunbosun Yemi','Olatubosun Yemi') 
        custname = custname.replace('Emem Owie Tongu','Emem Owei Tongu')
        custname = custname.replace('Charles Ehimare','Ehimare Charles')  
        
        
        custname = custname.replace('Ifeanyiwa Uchendu','Ifeyinwa Uchendu')                  
        custname = custname.replace('Oghenefejiro Monday','Oghenefein Monday')
        custname = custname.replace('Clara Wongi','Clara Wongo')
        custname = custname.replace('Ukpong James','Ubong James')  
        custname = custname.replace('Bridget Adebayo','Adebayo Bridget') 
        custname = custname.replace('Victory Eghra','Victory Eghre')
        custname = custname.replace('Philllip','Philip')
        custname = custname.replace('Sulieman','Suleiman')
        custname = custname.replace('Rita Mebine','Mebine Rita')
        custname = custname.replace('Newyork Samuel','Newyork Sambo')
        custname = custname.replace('Ayi Phillip','Ayi Philip')
        custname = custname.replace('Ann Otikor','Ann Otiko')
        custname = custname.replace('Seiyefa Cordilia','Seiyefa Cordelia')
        custname = custname.replace('Phillips Ebipade','Philip Ebipade')
        custname = custname.replace('Ebimo Ometeh','Ebimo Omeleh')
        custname = custname.replace('Clement Ederekumo','Ederekumo Clement')
        custname = custname.replace('Deregoe','Daregoe')
        custname = custname.replace('Andrew Kelly','Kelly Andrew')
        custname = custname.replace('Doris Obubor','Dorris Obutor')
        custname = custname.replace('Rhoda Okomi','Okomi Rhoda')
        custname = custname.replace('Patrick Adho','Patrick Adiho')
        custname = custname.replace('Daniel Andrew','Andrew Daniel')
        
        
        custname = custname.replace('Michael Titus','Micheal Titus')
        custname = custname.replace('Godspel','Gospel')
        custname = custname.replace('Samuel Angalere','Angalaere Samuel')
        custname = custname.replace('Pere James','James Pere')
        custname = custname.replace('Samuel Forsman','Samuel Forseman')
        custname = custname.replace('Timibitei Tari','Timiebite Tari')
            
    if type == 'SALES':
       
        custname = custname.replace('Kingley','Kingsley')
        custname = custname.replace('Kroboh Oweibiagha','Kroboh Oweibigha')
        custname = custname.replace('Kroboh Owiebigha','Kroboh Oweibigha')
        custname = custname.replace('kroboh Owei','Kroboh Oweibigha')
        custname = custname.replace('Kroboh Oweibia','Kroboh Oweibigha')
  
    return custname                
# test    
    
def reformat (file,prodtype):
    
    if 'OBUN' in file:
        PRODLOC = 'OBUNNA'
    if 'KPANSIA' in file:
        PRODLOC = 'KPANSIA'
    if 'DISP' in file:
        PRODLOC = 'DISPENSER'
    if 'CRATE' in file:
        PRODLOC = "CRATES"
    
    # Store Fido Dictionary into Python Dictionary 
    outfile = open("OUT/out_%s" %(file), 'w')
    reader1 = csv.reader(DICTFILE)
        
    for row in reader1:
        sperson = row[2].strip()
        sperson = sperson.upper()
        salesid[sperson] = row[3].strip()
        cperson = (row[0].upper()).strip()
        custid[cperson] = row[1].strip()
    
    
    file1 = 'data/'+file 
    reader2 = csv.reader(open(file1, 'rt'))
    CSVHEADER = 'id,payment_term_id/id,account_id/id,user_id/id,user_id/name,partner_id/id,partner_id/display_name,date_invoice,invoice_line_ids/product_id/id,invoice_line_ids/name,invoice_line_ids/account_id/id,invoice_line_ids/quantity,invoice_line_ids/price_unit'
    print (CSVHEADER,file=outfile)
    rcount = 0
    ercount = 0
    prdid = ""
    acctid = ""
    for row in reader2:
        try:
            rcount = rcount + 1
            custname = customerqc(row[1],'CUSTOMER')
            custname = custname.upper()
            salesperson = customerqc(row[0],'SALES')
            salesperson = salesperson.upper()
            # print(custname,' master ',salesperson)
            sid = ('__export__.res_users_' + salesid[salesperson]).rstrip()
            cid = ('__export__.res_partner_' + custid[custname]).rstrip()
            if 'OBUN' in file:
                printstr = ',__export__.account_payment_term_7,__export__.account_account_7,'+sid +','+salesperson+','+cid+','+custname+','+DATEINVOICE +','+'__export__.product_product_421'+','+prodtype+','+'__export__.account_account_204'+','+row[5]+','+row[6]
                
            if 'KPANSIA' in file:                
                printstr = ',__export__.account_payment_term_8,__export__.account_account_7,'+sid +','+salesperson+','+cid+','+custname+','+DATEINVOICE +','+'__export__.product_product_421'+','+prodtype+','+'__export__.account_account_204'+','+row[5]+','+row[6]
            
            if 'DISP' in file:                
                printstr = ',__export__.account_payment_term_9,__export__.account_account_7,'+sid +','+salesperson+','+cid+','+custname+','+DATEINVOICE +','+'__export__.product_product_433'+','+prodtype+','+'__export__.account_account_98'+','+row[5]+','+row[6]
            
            if 'CRATE' in file:
                prodtyp = row[3]
                if '50CL' in prodtyp:
                    prdid = '__export__.product_product_883'
                    acctid = '__export__.account_account_205'
                    
                if '60CL' in prodtyp:
                    prdid = '__export__.product_product_572' 
                    acctid = '__export__.account_account_206'
                
                printstr = ',__export__.account_payment_term_10,__export__.account_account_7,'+sid +','+salesperson+','+cid+','+custname+','+DATEINVOICE +','+prdid+','+prodtyp+','+acctid+','+row[5]+','+row[6]
            if (row[5]) and (row[6]):
                
                if (int(row[5]) != 0)  and (int(row[6]) != 0):
                    print (printstr,file=outfile)
        except KeyError as e:
            ercount = ercount + 1
            print('Customer,',row[1],',Salesperson,',row[0],',',e.args[0],',**** ',file,file=errfile)
            continue
    
    print ('PRODLOC,LINES,Errors,Errpct\n%s,%s,%s,%.2f%%\n' %(PRODLOC, rcount,ercount,((ercount-1)/(rcount-1))*100))
    outfile.close()
    
# Make import-ready files
def convfiles(folder):
    print ('ANALYSIS\n--------')
    for file in os.listdir(folder):
        if ('KPANSIA' in file) or ('OBUN' in file.upper()):
            reformat(file,'Purewater')
        elif ('DISPENSER' in file.upper()):
            reformat(file,'Dispenser')
        elif ('CRATE' in file.upper()):
            reformat(file,'Crate')
      
# Create csv files from sheets in Sales Workbook
def csvextract():
    wb = open_workbook(WBFILE)
    delfiles(DATAFOLDER)
    print ('SHEETS IN SALES FILE')
    
    for i in range(0, wb.nsheets-1):
        sheet = wb.sheet_by_index(i)
        print (sheet.name)
 
        path =  DATAFOLDER + '/%s.csv'
        with open( path %(sheet.name.replace(" ","")+DATEINVOICE), "w") as file:
            writer = csv.writer(file, delimiter = ",")
           
            header = [cell.value for cell in sheet.row(0)]
            writer.writerow(header)
 
            for row_idx in range(1, sheet.nrows):
                row = [int(cell.value) if isinstance(cell.value, float) else cell.value
                   for cell in sheet.row(row_idx)]
                writer.writerow(row)

def main():
    # extract csv from sheets in workbook
    csvextract()
    print ("\n")
    # Actual reformating
    convfiles(DATAFOLDER)
            
    print('See %s for error and data/ for source csv files\n and OUT/ directory for import-ready files' % ERRORFILE)        
            
if __name__ == '__main__':
    main()