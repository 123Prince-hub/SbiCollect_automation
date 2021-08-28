# import all python libraries
import re
import os
import time
import pytesseract
import xlwings as xw
import datetime as dt
from PIL import Image
from selenium import webdriver
from selenium.webdriver.support.select import Select 

ws = xw.Book(r'anjushree.xlsx').sheets("data")
rows = ws.range("A2").expand().options(numbers=int).value     # excel file read from second row
driver = webdriver.Chrome(executable_path="C:\Program Files (x86)\chromedriver.exe")

for l in range(4):

    num = 2
    progress = 'Server Not Response'
    for row in rows:
        if (row[0] != "Recharge payment") or (row[0] == None) or (any(map(str.isdigit, row[0])) or (type(row[0])) == int):
            progress = "Payment_Category Not Match in sheet"
            ws.range("M"+str(num)).value = progress
            exit()
        
        if (row[1]==None) or (any(map(str.isdigit, row[1])) or (type(row[1])) == int):
            progress = "Name is Empty in sheet"
            ws.range("M"+str(num)).value = progress
            exit()
        
        if (row[2]==None) or (len(str(row[2])) != 10):
            progress = "Mobile Number is Not Define in sheet"
            ws.range("M"+str(num)).value = progress
            exit()

        row[3] = int(row[3])
        if (row[3]==None) or (row[3] < 0):
            progress = "Amount is Invalid in sheet"
            ws.range("M"+str(num)).value = progress
            exit()
        
        if (row[4]==None):
            progress = "DOB is not valid or empty in sheet"
            ws.range("M"+str(num)).value = progress
            exit()

        if (len(str(row[5])) != 16):
            progress = "Card is Not define in sheet"
            ws.range("M"+str(num)).value = progress
            exit()
          
        if (row[7]==None) or ("/" not in str(row[7])):
            progress = "Exp Date Not Define in sheet"
            ws.range("M"+str(num)).value = progress
            exit()
          
        if (row[8]==None) or (len(str(row[8])) != 3 ):
            progress = "CVV not Define in sheet"
            ws.range("M"+str(num)).value = progress
            exit()
          
        if (row[9]==None) or (len(str(row[9])) != 4):
            progress = "Ipin Not define in sheet"
            ws.range("M"+str(num)).value = progress
            exit()

        # set dob format using diffrent length 
        if len(str(row[4])) > 15:
            row[4] = row[4].strftime('%d/%m/%Y')
        else:
            row[4] = row[4]

        # exp date seprate in month(mmm) & Year(yy)
        date = row[7]
        mm = date[1:2]
        yy = date[3:]

        col = ws.range("L"+str(num)).value
        if (col == None) or (col == "Pending") or (col == "NA"):

            # Start Automation
            try:
                try:
                    driver.maximize_window()
                    driver.set_page_load_timeout(8)
                    driver.get("https://www.onlinesbi.com/sbicollect/icollecthome.htm?corpID=3736558")
                except Exception as e:
                    progress = "101"

                # Click on Checkbox
                try:
                    driver.find_element_by_xpath('//*[@id="proceedcheck_english"]').click()
                except NoSuchElementException:
                    progress = "102"

                # Enter Proceed Button
                try:
                    driver.find_element_by_xpath('/html/body/div[1]/section/div/div/div[1]/form/div[2]/button').click()
                except NoSuchElementException:
                    progress = "103"

                # Second Page .......> Dropdown Button

                time.sleep(1)
                                # get 2nd Screen Url.......get Server response error
                try:
                    driver.current_url
                except NoSuchElementException:
                    progress = "104" 
                
                # Second Screen Dropdown button
                try:
                    driver.find_element_by_xpath('/html/body/div/section/div/div/div[1]/form/div/div/div[2]/div/div[2]/div/button').click()
                except NoSuchElementException:
                    progress = "105"

                # pic value from 1st coloums from sheet
                try:                    
                    pyment = row[0]
                except NoSuchElementException:
                    progress = "106"

                # Select Value from Dropdwon
                try:    
                    driver.find_element_by_link_text("Recharge payment").click()
                except NoSuchElementException:
                    progress = "107"

                # Third Screen
                time.sleep(1)
                try:
                    driver.current_url
                except NoSuchElementException:
                    progress = "108"

                # 1st Name Field
                try:
                    Name = driver.find_element_by_xpath('//*[@id="outref11"]').send_keys(row[1])
                except NoSuchElementException:
                    progress = "109"

                # 1s Mobile Field
                try:
                    Mobile = driver.find_element_by_xpath('//*[@id="outref12"]').send_keys(row[2])
                except NoSuchElementException:
                    progress = "110"

                # Amount Field
                try:    
                    Amount = driver.find_element_by_xpath('//*[@id="outref14"]').send_keys(row[3])
                except NoSuchElementException:
                    progress = "111"

                # 2nd Name Field
                try:
                    Name = driver.find_element_by_xpath('//*[@id="cusName"]').send_keys(row[1])
                except NoSuchElementException:
                    progress = "112"

                # Dob Field
                try:
                    driver.execute_script("document.getElementById('dateOfBirth').value= '"+row[4]+"'")
                except NoSuchElementException:
                    progress = "113"

                # 1s Mobile Field
                try:    
                    Mobile = driver.find_element_by_xpath('//*[@id="mobileNo"]').send_keys(row[2])
                except NoSuchElementException:
                    progress = "114"

                # email = driver.find_element_by_xpath('//*[@id="emailId"]').send_keys(row[])

                # captcha Code scan
                driver.find_element_by_id('captchaImage').screenshot('screenshot.png')
                pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
                a = pytesseract.image_to_string(r'screenshot.png')
                a = a.upper()
                driver.find_element_by_id('captchaValue').send_keys(a.lstrip())

                # Click Submit
                try:
                    driver.find_element_by_xpath('/html/body/div[1]/section/div/div/div/form[2]/div[3]/button[1]').click()
                except NoSuchElementException:
                    progress = "115"
                
                time.sleep(1)
                try:
                    alt = driver.switch_to.alert
                    # err = driver.find_elements_by_id('captchaValue-error').text   Please enter valid text as shown in the image
                    i = 0
                    while (alt.text=="Please enter valid captcha") and (i<10):
                        alt.accept()
                        # captcha Code scan
                        time.sleep(1)
                        driver.find_element_by_id('captchaImage').screenshot('screenshot.png')
                        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
                        a = pytesseract.image_to_string(r'screenshot.png')
                        a = a.upper()
                        driver.find_element_by_id('captchaValue').send_keys(a.lstrip())
                        # Submit Button
                        try:
                            driver.find_element_by_xpath('/html/body/div[1]/section/div/div/div/form[2]/div[3]/button[1]').click()
                        except NoSuchElementException:
                            progress = "116"

                        time.sleep(1)
                        i = i+1
                except:
                    os.remove('screenshot.png')

                    # Third Page
                    time.sleep(1)

                    # get 3rd Screen Url.......get Server response error
                    try:
                        driver.current_url
                    except NoSuchElementException:
                        progress = "117"

                    try:
                        driver.find_element_by_xpath('//*[@id="collect"]/div[3]/button[1]').click()
                    except NoSuchElementException:
                        progress = "118"

                    #  Fourth Page
                    time.sleep(1)

                    # get 4th Screen Url.......get Server response error
                    try:
                        driver.current_url
                    except NoSuchElementException:
                        progress = "119"

                    # Fourth Page Card Select
                    try:
                        driver.find_element_by_xpath('/html/body/div/form/section/div[2]/div/div[4]/div/a').click()
                    except NoSuchElementException:
                        progress = "120"


                    # ****************** Sixth Page......Fill Card Details ****************    

                    # get 6th Screen Url.......get Server response error
                    try:
                        driver.current_url
                    except NoSuchElementException:
                        progress = "121"


                    # ****************** Fill Card Details ****************    
                    # Card No Field....Enter Card Number

                    time.sleep(1)
                    try:
                        driver.find_element_by_xpath('//*[@id="cardNumber"]').send_keys(row[5])
                    except NoSuchElementException:
                        progress = "122"
                    
                    # Exp Month Date Dropdown
                    try:
                        s2= Select(driver.find_element_by_id('expMnthSelect'))
                    except NoSuchElementException:
                        progress = "123"

                    try:
                        s2.select_by_value(mm)
                    except NoSuchElementException:
                        progress = "124"
                    
                    # Exp Year Date Dropdown
                    try:
                        s2= Select(driver.find_element_by_id('expYearSelect'))
                    except NoSuchElementException:
                        progress = "125"

                    try:    
                        s2.select_by_value(yy)
                    except NoSuchElementException:
                        progress = "126"

                    # Card Holder Name
                    try:
                        driver.find_element_by_xpath('//*[@id="cardholderName"]').send_keys(row[1])
                    except NoSuchElementException:
                        progress = "127"
                    
                    # CVV Field
                    try:
                        driver.find_element_by_xpath('//*[@id="cvd2"]').send_keys(row[8])            
                    except NoSuchElementException:
                        progress = "128"

                    # captcha Code scan
                    time.sleep(1)
                    driver.find_element_by_id('captcha_image').screenshot('screenshot2.png')
                    img = Image.open('screenshot2.png')
                    imgGray = img.convert('L')
                    imgGray.save('screenshot2.png')
                    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
                    time.sleep(1)
                    a = pytesseract.image_to_string(r'screenshot2.png').replace(" ", "")
                    a = a.replace(")", "J")
                    time.sleep(1)
                    driver.find_element_by_id('passline').send_keys(a.upper())

                    # Click Submit
                    time.sleep(1)
                    try:
                        driver.find_element_by_xpath('//*[@id="proceed_button"]').click()
                    except NoSuchElementException:
                        progress = "129"

                    try:
                        time.sleep(1)
                        error = driver.find_element_by_xpath('//*[@id="cardCaptchaMsg"]')
                        time.sleep(1)
                        j = 0
                        while ((error.text=="Invalid captcha") or (error.text=="Missing captcha")) and (j<6):
                            driver.find_element_by_xpath('//*[@id="cardNumber"]').send_keys(row[5])
                            driver.find_element_by_xpath('//*[@id="cardholderName"]').send_keys(row[1])
                            driver.find_element_by_xpath('//*[@id="cvd2"]').send_keys(row[8])
                            time.sleep(1)
                            driver.find_element_by_id('captcha_image').screenshot('screenshot2.png')
                    
                            img = Image.open('screenshot2.png')
                            imgGray = img.convert('L')
                            imgGray.save('screenshot2.png')
            
                            pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
                            a = pytesseract.image_to_string(r'screenshot2.png').replace(" ", "")
                            a = a.replace(")", "J")
                            driver.find_element_by_id('passline').send_keys(a.upper())
                            time.sleep(1)
                            # Click Submit
                            driver.find_element_by_xpath('//*[@id="proceed_button"]').click()
                            time.sleep(1)
                            j = j+1

                    except:
                        os.remove('screenshot2.png')
                        # Sixth Page ... Ipin Screen
                        time.sleep(1)

                        # get 6th Screen Url.......get Server response error
                        try:
                            driver.current_url
                        except NoSuchElementException:
                            progress = "130"

                        try:
                            ipin = driver.execute_script("document.querySelector('#txtipin').value = '"+str(row[9])+"'")
                        except NoSuchElementException:
                            progress = "131"

                        time.sleep(1)
                        try:
                            driver.find_element_by_xpath('//*[@id="btnverify"]').click()
                        except NoSuchElementException:
                            progress = "132"

                        # save SBCollect Reference Number value in variable to excelsheet
                        time.sleep(1)        
                
                # Sevnth Page..........Duo Number Screen   
                         
                # get 7th Screen Url.......get Server response error
                try:
                    driver.current_url
                except NoSuchElementException:
                    progress = "133"

                deo_num = driver.find_element_by_xpath('//*[@id="collect"]/div[2]/div/div[2]/span/strong | //*[@id="printdetailsformtop"]/div/div/div[2]/span/strong').text
                pass_reson = driver.find_element_by_xpath('//*[@id="collect"]/div[2]/div/p[2]/strong | //*[@id="printdetailsformtop"]/div/div/p[1]/strong').text
                
                if "successfully" in pass_reson:
                    ws.range("K"+str(num)).value = deo_num
                    ws.range("L"+str(num)).value = "Success"
                    ws.range("M"+str(num)).value = "Ok"
                        

                elif "Failure" in pass_reson:
                    ws.range("K"+str(num)).value = deo_num
                    ws.range("L"+str(num)).value = "Pending"
                    ws.range("M"+str(num)).value = "403"
                else:
                    pass
        
            except:
                # print("errroor")
                ws.range("K"+str(num)).value = "NA"
                ws.range("L"+str(num)).value = "NA"
                ws.range("M"+str(num)).value = progress
                progress = "Server Not Response"

        num += 1
driver.close()