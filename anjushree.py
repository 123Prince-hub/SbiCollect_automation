import xlwings as xw
import datetime as dt
from PIL import Image
import os, time, pytesseract
from selenium import webdriver
from selenium.webdriver.support.select import Select
from webdriver_manager.chrome import ChromeDriverManager

ws = xw.Book(r'anjushree.xlsx').sheets("data")
rows = ws.range("A2").expand().options(numbers=int).value
driver = webdriver.Chrome(ChromeDriverManager().install()) 

for l in range(4):
    num = 2
    for row in rows:

        if len(str(row[4])) > 15:
            row[4] = row[4].strftime('%d/%m/%Y')
        else:
            row[4] = row[4]

# *************** exp date ************************
        date = row[7]
        if date[0]=="0":
            mm = date[1]
        else:
            mm = date[:2]
        yy = date[3:]

        col = ws.range("L"+str(num)).value
        if (col == None) or (col == "Pending") or (col == "NA"):
            try:
                driver.implicitly_wait(60)
                driver.set_page_load_timeout(8)
                driver.maximize_window()
                driver.get("https://www.onlinesbi.com/sbicollect/icollecthome.htm?corpID=3736558")

                # Click on Checkbox
                driver.find_element_by_xpath('//input[@type="checkbox"]').click()

                # Enter Proceed Button
                driver.find_element_by_xpath('//button[contains(text(),"Proceed")]').click()

                # Second Page
                time.sleep(1)
                driver.find_element_by_xpath('//span[contains(text(),"-- Select Category --")]').click()
                driver.find_element_by_xpath("//span[contains(text(),'Recharge payment')]").click()

                time.sleep(1)
                Name = driver.find_element_by_xpath('//label[contains(text(), "Name *")]//following::input').send_keys(row[1])
                Mobile = driver.find_element_by_xpath('//label[contains(text(), "Mobile No. *")]//following::input').send_keys(row[2])
                Remark = driver.find_element_by_xpath('//label[contains(text(), "Remark")]//following::input').send_keys(row[5])
                Amount = driver.find_element_by_xpath('//label[contains(text(), "recharge to be done  *")]//following::input').send_keys(row[3])
                Name = driver.find_element_by_xpath('//*[@id="cusName"]').send_keys(row[1])
                driver.execute_script("document.getElementById('dateOfBirth').value= '"+row[4]+"'")
                Mobile = driver.find_element_by_xpath('//*[@id="mobileNo"]').send_keys(row[2])

                # captcha Code scan
                driver.find_element_by_id('captchaImage').screenshot('screenshot.png')
                pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
                a = pytesseract.image_to_string(r'screenshot.png')
                a = a.upper()
                driver.find_element_by_xpath('//label[contains(text(), "Enter the text as shown in the image *")]//following::input').send_keys(a.lstrip())

                # Click Submit
                driver.find_element_by_xpath('//button[contains(text(),"Submit")]').click()
                

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
                        driver.find_element_by_xpath('//label[contains(text(), "Enter the text as shown in the image *")]//following::input').send_keys(a.lstrip())
                        driver.find_element_by_xpath('//button[contains(text(),"Submit")]').click()
                        time.sleep(1)
                        i = i+1
                except:
                    # Next Page
                    os.remove('screenshot.png')
                    time.sleep(1)
                    driver.find_element_by_xpath('//button[contains(text(),"Confirm")]').click()

                    # Next Page
                    time.sleep(1)
                    driver.find_element_by_xpath('//span[contains(text(),"Other Bank Debit Cards")]//following::a').click()

                    # ****************** Fill Card Details ****************    
                    time.sleep(1)
                    driver.find_element_by_xpath('//label[contains(text()," Card Number ")]//following::input').send_keys(row[5])
                    s2= Select(driver.find_element_by_id('expMnthSelect'))
                    s2.select_by_value(mm)
                    
                    s2= Select(driver.find_element_by_id('expYearSelect'))
                    s2.select_by_value(yy)
                    driver.find_element_by_xpath('//label[contains(text(),"Card Holders Name")]//following::input').send_keys(row[1])
                    driver.find_element_by_xpath('//*[@id="cvd2"]').send_keys(row[8])            
                    time.sleep(1)
                    # captcha Code scan
                    driver.find_element_by_id('captcha_image').screenshot('screenshot2.png')
                    img = Image.open('screenshot2.png')
                    imgGray = img.convert('L')
                    imgGray.save('screenshot2.png')
                    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
                    time.sleep(1)
                    a = pytesseract.image_to_string(r'screenshot2.png').replace(" ", "")
                    a = a.replace(")", "J")
                    driver.find_element_by_xpath('//label[contains(text(),"Type the characters")]//following::input').send_keys(a.upper())

                    # Click Submit
                    time.sleep(1)
                    driver.find_element_by_xpath('//input[@value="Pay"]').click()
                    try:
                        time.sleep(1)
                        error = driver.find_element_by_xpath('//*[@id="cardCaptchaMsg"]')
                        time.sleep(1)
                        j = 0
                        while ((error.text=="Invalid captcha") or (error.text=="Missing captcha")) and (j<51):
                            driver.find_element_by_xpath('//*[@id="cardNumber"]').send_keys(row[5])
                            driver.find_element_by_xpath('//*[@id="cardholderName"]').send_keys(row[1])
                            time.sleep(1)
                            driver.find_element_by_xpath('//*[@id="cvd2"]').send_keys(row[8])
                            time.sleep(1)
                            driver.find_element_by_id('captcha_image').screenshot('screenshot2.png')
                    
                            img = Image.open('screenshot2.png')
                            imgGray = img.convert('L')
                            imgGray.save('screenshot2.png')
            
                            pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
                            a = pytesseract.image_to_string(r'screenshot2.png').replace(" ", "")
                            a = a.replace(")", "J")
                            driver.find_element_by_xpath('//label[contains(text(),"Type the characters")]//following::input').send_keys(a.upper())
                            time.sleep(1)
                            driver.find_element_by_xpath('//input[@value="Pay"]').click()
                            time.sleep(1)
                            j = j+1

                    except:
                        os.remove('screenshot2.png')
                    # Next Page
                    time.sleep(1)
                    ipin = driver.execute_script("document.querySelector('#txtipin').value = '"+str(row[9])+"'")
                    # ipin = driver.find_element_by_xpath('//label[contains(text(), "IPIN :")]//following::input').send_keys(row[9])
                    driver.find_element_by_xpath('//input[@value="Submit"]').click()
                    time.sleep(1)        
                
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
                ws.range("K"+str(num)).value = "NA"
                ws.range("L"+str(num)).value = "NA"
                ws.range("M"+str(num)).value = "Server not Response"

        num += 1
driver.close()