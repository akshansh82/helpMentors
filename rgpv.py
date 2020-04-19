import os
import time
import openpyxl
import urllib.request
import pytesseract as tess
import backcall
from PIL import Image
from tqdm import tqdm
from colorama import Fore
from colorama import Style
from openpyxl import workbook
from selenium import webdriver 


from selenium.webdriver.common.keys import Keys  
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.common.exceptions import TimeoutException 
from selenium.common.exceptions import NoSuchElementException 
from selenium.common.exceptions import UnexpectedAlertPresentException 


os.system('cls')
#########################################################################################################################



def res(start,sem):

    success = True
    while start <= end :
    
        if len(str(start)) == 1:
            roll = '00' + str(start)
        elif len(str(start)) == 2:
            roll = '0' + str(start)
        else :
            roll = str(start)

        #time.sleep(0.5)

        try:
            driver.find_element_by_xpath('//*[@id="radlstProgram_1"]').click()
            time.sleep(0.5)
            driver.find_element_by_id("ctl00_ContentPlaceHolder1_txtrollno").send_keys(common + roll)
            sems = Select(driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_drpSemester"]'))
            sems.select_by_value(sem)
            start = captcha(start)+1
        except NoSuchElementException:
            print(f'{Fore.RED}Please check your internet connection{Style.RESET_ALL}')
            success = False
            break

        
        driver.get('http://result.rgpv.ac.in/Result/ProgramSelect.aspx')
        

    driver.close()
    if success == True:
        print(f'{Fore.GREEN}Successfull saved Excel file  {Fore.YELLOW}{shti} .xlsx at  {dest}{Style.RESET_ALL}')



###########################################################################################################################



def captcha(start):
     

    global prev
    global count

    src = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_pnlCaptcha"]/table/tbody/tr[1]/td/div/img').get_attribute('src')
    urllib.request.urlretrieve(src,'captcha.jpg')
    for i in tqdm(range(100), leave= False, desc= 'Connecting',bar_format= "{l_bar}%s{bar}%s{r_bar}" % (Fore.CYAN, Fore.RESET)):
        time.sleep(0.01)

    img = Image.open('captcha.jpg')
    text = tess.image_to_string(img).replace(' ', '').upper()
    for i in tqdm(range(100), desc= 'Extracting ', leave=False, bar_format= "{l_bar}%s{bar}%s{r_bar}" % (Fore.CYAN, Fore.RESET)):
        time.sleep(0.02)

    driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_TextBox1"]').clear()
    driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_TextBox1"]').send_keys(text)
    for i in tqdm(range(100), desc= 'Processing ', leave=False, bar_format= "{l_bar}%s{bar}%s{r_bar}" % (Fore.CYAN, Fore.RESET)):
        time.sleep(0.01)
    driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_btnviewresult"]').click()


    try:
        WebDriverWait(driver,3).until(EC.alert_is_present(),
                                    'Timed out waiting for PA creation ' +
                                    'confirmation popup to appear.')

        alert = driver.switch_to.alert
        alert.accept()
        if(prev == start):
            if(count <= 3):
                count += 1
                return start - 1
            else:
                count = 1
                sh['A' + str(start)].value = 'No Result Found'
                prev = start + 1
                return start
        
        
    except TimeoutException:
        #time.sleep(2)
        count = 1
        save()
        prev = start + 1
        return start




############################################################################################################################



def save():
        name = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_lblNameGrading"]').text
        enRO = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_lblRollNoGrading"]').text
        sgpa = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_lblSGPA"]').text
        cgpa = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_lblcgpa"]').text
        dess = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_lblResultNewGrading"]').text
        en   = int(enRO[-3:])
        sh['A'+ str(en)].value = name
        sh['B'+ str(en)].value = float(sgpa) 
        sh['C'+ str(en)].value = float(cgpa)
        sh['D'+ str(en)].value = "     " + dess
        wb.save(path+str(shti)+'.xlsx')
        print(f'{Fore.LIGHTBLUE_EX}{enRO}\t{name:{25}}\t{sgpa}\t{cgpa}\t{dess}{Style.RESET_ALL}')



#########################################################################################################################     
start = ''
end   = ''
sem   = ''
common = ''
inputs = True
try:
    start  = str(input(f'enter starting enrollment number :\t'))
    end    = str(input('enter the last enrollment number :\t'))
    if start[:-3] == end[:-3]:
        common = start[:-3]
        start = int(start[-3:])
        end   = int(end[-3:])
        try:
            sem    = int(input('enter the semester :  \t\t\t'))
        except ValueError:
            print(f'{Fore.RED}Semester Must be Intger Number')
            inputs = False
    else:
        print(f'{Fore.RED}ERROR :  ' ,start[:-3] , ' and ' , end[:-3] , f'Do Not Match{Style.RESET_ALL}')
        inputs = False

except ValueError:
    print(f'{Fore.RED}ERROR : Check Enrollment Number{Style.RESET_ALL}')
    inputs = False



#########################################################################################################################



count = 1
prev = start
if inputs == True:
    print('\n\n \tFetching Results Please Wait For a Minute \t\n\n')
    roll = '000'
    options = Options()
    options.headless =  True
    options.add_argument('--start-maximized')


    driver = webdriver.Chrome(executable_path = 'chromedriver.exe',options=options)
    tess.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    # driver.maximize_window()        
    driver.get('http://result.rgpv.ac.in/Result/ProgramSelect.aspx')



    ########################################################################################################################



    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    desktop = desktop.split('\\')
    path = ''
    dest = ''
    for i in desktop:
        path += i+'\\\\'
    for i in desktop:
        dest += i+'\\'



    #########################################################################################################################



    wb = openpyxl.Workbook()
    sh = wb.active
    shti = str(start) + ' - ' + str(end)
    sh.column_dimensions['A'].width = 25
    os.system('cls')
    res(start,str(sem))


