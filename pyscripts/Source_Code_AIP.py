from selenium import webdriver as wd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as wdw
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import TimeoutException as te
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import time
from datetime import datetime
import openpyxl as XL
import os

fls = []
vendat = []
vendatE = []
TP = 0.6
CT = time.time()
filePath = 'FileS'
filePathCP = 'GoogleChrome'

Path = os.path.join(os.path.dirname(os.path.abspath(__name__)),filePath)
CPPath = os.path.join(os.path.dirname(os.path.abspath(__name__)),filePathCP)
CDPath = os.path.join(os.path.dirname(os.path.abspath(__name__)))

#print(CPPath)
#print(CDPath)



class VMTermsUpdate:
    def __init__(self):
        li = os.listdir(Path)
        for flod, subflo, file in os.walk(Path):
            for fil in file:
                if fil.endswith('.xlsx'):
                    fls.append(os.path.join(flod,fil))
        WBR = XL.load_workbook(fls[0],data_only=True)
        sheet = WBR[WBR.sheetnames[0]]
        VaL = str(input('Enter last row: '))
        print(VaL)
        for inu,c in enumerate(sheet['A2':'B'+VaL]):
            vendat.append([])
            vendatE.append([])
            for va in c:
                vendat[inu].append(str(va.value))
                vendatE[inu].append(str(va.value))        


    def Driv():
        options = Options()
        prefs = {
            "profile.default_content_setting_values.plugins": 1,
            "profile.content_settings.plugin_whitelist.adobe-flash-player": 1,
            "profile.content_settings.exceptions.plugins.*,*.per_resource.adobe-flash-player": 1,
            "PluginsAllowedForUrls": "https://kof.bizsys.pearson.com/markview/MVT_Web_Inquiry.ShowInquiry"
        }
        # options.binary_location = f"{CPPath}\\chrome.exe"
        options.add_experimental_option('prefs',prefs)
        options.add_argument('--start-maximized')
        options.add_argument('--disable-extensions')
        # return wd.Chrome(executable_path=f'{CDPath}\\chromedriver.exe',options=options)
        return wd.Firefox()

    D = Driv()
    def StaRt(self,ent):
        # D = VMTermsUpdate.Driv(self)
        if ent.lower() == 'us':
            VMTermsUpdate.D.get('https://ebs.bizsys.pearson.com/OA_HTML/RF.jsp?function_id=1348&resp_id=52134&resp_appl_id=200&security_group_id=0&lang_code=US&oas=UkWdsrjY0wo5zsHT94VgLg..&params=3186AIGAI8sS0D7oPtsTLlA3r.pdeS9lKBabtgXem3U') #US
        elif ent.lower() == 'ca':
            VMTermsUpdate.D.get(' https://ebs.bizsys.pearson.com/OA_HTML/RF.jsp?function_id=1348&resp_id=52088&resp_appl_id=200&security_group_id=0&lang_code=US&oas=AbzQpn5vQmmzmzzaKOIxWg..&params=3186AIGAI8sS0D7oPtsTLlA3r.pdeS9lKBabtgXem3U') #Cannada            
            # D.get('http://localhost:7000/')

        timeout = 10
        try:
            wdw(VMTermsUpdate.D, timeout).until(ec.visibility_of_element_located((By.XPATH, '//*[@id="user-name-txt"]')))
            print('Complete MFA')

        except te:
            print('Time Up')
            VMTermsUpdate.D.quit()

        timeout = 60
        try:
            wdw(VMTermsUpdate.D, timeout).until(ec.visibility_of_element_located((By.XPATH, '//*[@id="region10"]/div[1]/table/tbody/tr/td/h1')))

        except te:
            print('You Did not Enter the MFA Code')
            VMTermsUpdate.D.quit()
    
    def ElemX(E):
        time.sleep(TP)
        return VMTermsUpdate.D.find_element_by_xpath(E)                    
    message = 'None'
    
    def sivin(self,venNum,venNam):
        print('_'*61)
        if venNum.lower() == 'y' and venNam.lower() == 'y':
            def venSerchLogic(H):
                VMTermsUpdate.ElemX('//*[@id="SearchSuppName"]').send_keys(H[0])
                VMTermsUpdate.ElemX('//*[@id="SearchSuppNum"]').send_keys(H[1])
                # print(f'{venNum} and {venNam}')
        elif venNum.lower() == 'n' and venNam.lower() == 'y':
            def venSerchLogic(H):
                VMTermsUpdate.ElemX('//*[@id="SearchSuppName"]').send_keys(H[0])
                # print(f'{venNum} and {venNam}')
        elif venNum.lower() == 'y' and venNam.lower() == 'n':
            def venSerchLogic(H):
                VMTermsUpdate.ElemX('//*[@id="SearchSuppNum"]').send_keys(H[1])
                # print(f'{venNum} and {venNam}')
        corrLink = None
        sitenotfound = ''
        for P,H in enumerate(vendat):
            venSerchLogic(H)

            VMTermsUpdate.ElemX('//*[@id="GoButton"]').click()
            
            ec.visibility_of_element_located((By.XPATH, '//*[@id="p_SwanPageLayout"]/div[1]/div[1]/table/tbody/tr/td[1]/h1'))
            SITEMAT = False
            try:
                #Gone inside venodr Data
                VMTermsUpdate.ElemX('//*[@id="POS_HT_SP_B_TAX_RPT"]').click()
                time.sleep(1)
                TICK = VMTermsUpdate.ElemX('//*[@id="N48:incTaxSite:0"]').is_selected() 
                if TICK:
                    VMTermsUpdate.ElemX('//*[@id="cancelBtn"]').click()
                    VMTermsUpdate.message = ('{}|{}|{}|{}|{}|{}'.format(H[0],H[1],'','Tax Reporting Site was Already Ticked ','No Changes Made By Automation','Success'))
                    print('{}|{}|{}|{}|{}|{}'.format(H[0],H[1],'','Tax Reporting Site was Already Ticked ','No Changes Made By Automation','Success'))
                    VMTermsUpdate.ElemX('//*[@id="POS_HT_SP_B_SUPP"]').click()
                else:
                    VMTermsUpdate.ElemX('//*[@id="N48:incTaxSite:0"]').click()
                    VMTermsUpdate.ElemX('//*[@id="saveBtn"]').click()

                    if VMTermsUpdate.ElemX('//*[@id="FwkErrorBeanId"]/tbody/tr/td/table/tbody/tr/td/div[3]').get_attribute('innerText') != 'Changes to Tax and Reporting have been saved':
                        VMTermsUpdate.ElemX('//*[@id="confirmBtn"]').click()
                        
                    SITEMAT = True
                    VMTermsUpdate.ElemX('//*[@id="POS_HT_SP_B_SUPP"]').click()

                
                    if SITEMAT:
                        VMTermsUpdate.message = ('{}|{}|{}|{}|{}|{}'.format(H[0],H[1],'','Tax Reporting Site Ticked','Changes Made By Automation','Success'))
                        print('{}|{}|{}|{}|{}|{}'.format(H[0],H[1],'','Tax Reporting Site Ticked','Changes Made By Automation','Success'))
                    else:
                        VMTermsUpdate.message = ('{}|{}|{}|{}|{}|{}'.format(H[0],H[1],'','Something got error in Tax an Reporting','No Changes Made By Automation','Error'))
                        print('{}|{}|{}|{}|{}|{}'.format(H[0],H[1],'','Something got error in Tax an Reporting','No Changes Made By Automation','Error'))

            except:
                VENNOT = VMTermsUpdate.ElemX('//*[@id="ResultRN:Content"]/tbody/tr/td[1]').get_attribute('innerText')
                if VENNOT == 'No results found.':
                    # print('{}|Vendor Not Found|No Changes Made.|Error'.format(H[0]))
                    # print('_'*88)
                    VMTermsUpdate.message = ('{}|{}|{}|{}|{}|{}'.format(H[0],H[1],'','Vendor Not Found','No Changes Made By Automation','Error'))
                    print('{}|{}|{}|{}|{}|{}'.format(H[0],H[1],'','Vendor Not Found','No Changes Made By Automation','Error'))
                    VMTermsUpdate.ElemX('//*[@id="POS_HT_SP_B_SUPP"]').click()

                VENMUL = VMTermsUpdate.ElemX('//*[@id="ResultRN:Content"]/tbody').find_elements_by_tag_name('tr')
                if len(VENMUL) > 1:
                    # print('{}|Multiple Vendor Found|No Changes Made.|Error'.format(H[0]))
                    # print('_'*79)
                    VMTermsUpdate.message = ('{}|{}|{}|{}|{}|{}'.format(H[0],H[1],'','Multiple Vendor Found','No Changes Made By Automation','Error'))
                    print('{}|{}|{}|{}|{}|{}'.format(H[0],H[1],'','Multiple Vendor Found','No Changes Made By Automation','Error'))
                    VMTermsUpdate.ElemX('//*[@id="POS_HT_SP_B_SUPP"]').click()
            with open('venNam'+'.txt','a') as FLS:
                FLS.write(VMTermsUpdate.message+'\n')
        VMTermsUpdate.ElemX('//*[@id="POS_HT_SP_B_SUPP"]').click()



VMTU1 = VMTermsUpdate()

def askEntity():
    Q = input('US / CA ? ')
    # VMTU1.Driv()
    VMTU1.StaRt(Q)

def askQuest():
    Q = input('Vendor Number? | Vendor Name? ')
    dat = Q.split(',')
    VMTU1.sivin(dat[0],dat[1])



askEntity()
askQuest()

print('Completed')
