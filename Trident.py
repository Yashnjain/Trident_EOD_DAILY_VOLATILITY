from ast import Global
import csv
from operator import index
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
import time
from datetime import date
import logging
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager
import os
from bu_config import get_config
import bu_alerts
import smtplib
import email.mime.multipart
import email.mime.text
import email.mime.base
import email.encoders as encoders
import tabula
import xlwings.constants as win32c
import xlwings as xw
import bu_snowflake
import pandas as pd
from snowflake.connector.pandas_tools import pd_writer
import functools

locations_list=[]

today_date=date.today()
# log progress --
logfile = os.getcwd() +"\\logs\\"+'TRIDENT_EOD_DAILY_VOLATILITY_Logfile'+str(today_date)+'.txt'

logging.basicConfig(filename=logfile, filemode='w',
                    format='%(asctime)s %(message)s')
logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s [%(levelname)s] - %(message)s',
    filename=logfile)

logger = logging.getLogger()
logger.setLevel(logging.INFO)
logging.info('setting paTH TO DOWNLOAD')
path = os.getcwd() + "\\"+"Download"
logging.info('SETTING PROFILE SETTINGS FOR FIREFOX')



profile = webdriver.FirefoxProfile()
profile.set_preference('browser.download.folderList', 2)
profile.set_preference('browser.download.dir', path)
profile.set_preference('browser.download.useDownloadDir', True)
profile.set_preference('browser.download.viewableInternally.enabledTypes', "")
profile.set_preference('browser.helperApps.neverAsk.saveToDisk','Portable Document Format (PDF), application/pdf')
profile.set_preference('pdfjs.disabled', True)
logging.info('Adding firefox profile')
driver=webdriver.Firefox(executable_path=GeckoDriverManager().install(),firefox_profile=profile)

credential_dict = get_config('TRIDENT_EOD_DAILY_VOLATILITY','TRIDENT_EOD_DAILY_VOLATILITY')
username = credential_dict['USERNAME'].split(';')[0]
password = credential_dict['PASSWORD'].split(';')[0]
table_name = credential_dict['TABLE_NAME']
Database = credential_dict['DATABASE']
SCHEMA = credential_dict['TABLE_SCHEMA']
# receiver_email = credential_dict['EMAIL_LIST'].split(';')[0]
receiver_email = 'yashn.jain@biourja.com'
download_path=os.getcwd() + "\\Download"
file_name= os.listdir(os.getcwd() + "\\Download")
output_location= os.getcwd()+"\\Generated_CSV"
today_date=date.today()
download_path=os.getcwd() + "\\Download"
file_name= os.listdir(os.getcwd() + "\\Download")
file2=file_name[0]

test_area_date = ["67,47,85,160"]
df_date = tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages="1",area=test_area_date,silent=True,guess=False)
Trade_date=df_date[0].columns[1]
def num_to_col_letters(num):
    try:
        letters = ''
        while num:
            mod = (num - 1) % 26
            letters += chr(mod + 65)
            num = (num - 1) // 26
        return ''.join(reversed(letters))
    except Exception as e:
        raise e
def send_mail(receiver_email: str, mail_subject: str, mail_body: str, attachment_locations: list = None, sender_email: str = None, sender_password: str=None) -> bool:
    """The Function responsible to do all the mail sending logic.

    Args:
        sender_email (str): Email Id of the sender.
        sender_password (str): Password of the sender.
        receiver_email (str): Email Id of the receiver.
        mail_subject (str): Subject line of the email.
        mail_body (str): Message body of the Email.
        attachment_locations (list, optional): Absolute path of the attachment. Defaults to None.

    Returns:
        bool: [description]
    """
    logging.info("INTO THE SEND MAIL FUNCTION")
    done = False
    try:
        logging.info("GIVING CREDENTIALS FOR SENDING MAIL")
        if not sender_email or sender_password:
            sender_email = "biourjapowerdata@biourja.com"
            sender_password = r"bY3mLSQ-\Q!9QmXJ"
            # sender_email = r"virtual-out@biourja.com"
            # sender_password = "t?%;`p39&Pv[L<6Y^cz$z2bn"
        receivers = receiver_email.split(",")
        msg = email.mime.multipart.MIMEMultipart()
        msg['From'] = "biourjapowerdata@biourja.com"
        msg['To'] = receiver_email
        msg['Subject'] = mail_subject
        body = mail_body
        logging.info("Attaching mail body")
        msg.attach(email.mime.text.MIMEText(body, 'html'))
        logging.info("Attching files in the mail")
        for files_locations in attachment_locations:
            with open(files_locations, 'r+b') as attachment:
                # instance of MIMEBase and named as p
                p = email.mime.base.MIMEBase('application', 'octet-stream')
                # To change the payload into encoded form
                p.set_payload((attachment).read())
                encoders.encode_base64(p)  # encode into base64
                p.add_header('Content-Disposition',
                             "attachment; filename= %s" % files_locations)
                msg.attach(p)  # attach the instance 'p' to instance 'msg'

        # s = smtplib.SMTP('smtp.gmail.com', 587) # creates SMTP session
        s = smtplib.SMTP('us-smtp-outbound-1.mimecast.com',
                         587)  # creates SMTP session
        s.starttls()  # start TLS for security
        s.login(sender_email, sender_password)  # Authentication
        text = msg.as_string()  # Converts the Multipart msg into a string

        s.sendmail(sender_email, receivers, text)  # sending the mail
        s.quit()  # terminating the session
        done = True
        logging.info("Email sent successfully")
        print("Email sent successfully.")
    except Exception as e:
        print(
            f"Could not send the email, error occured, More Details : {e}")
    finally:
        return done
def remove_existing_files(files_location):
    """_summary_

    Args:
        files_location (_type_): _description_

    Raises:
        e: _description_
    """           
    logger.info("Inside remove_existing_files function")
    try:
        files = os.listdir(files_location)
        if len(files) > 0:
            for file in files:
                os.remove(files_location + "\\" + file)
            logger.info("Existing files removed successfully")
        else:
            print("No existing files available to reomve")
        print("Pause")
    except Exception as e:
        logger.info(e)
        raise e

def login_and_download():  
    '''This function downloads log in to the website'''
    try:
        logging.info('Accesing website')
        driver.get("https://outlook.office365.com/owa/biourja.com/")
        logging.info('providing id and passwords')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "i0116"))).send_keys(username)
        time.sleep(1)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        time.sleep(1)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys(password)
        time.sleep(1)
        logging.info('click on No Button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        time.sleep(1)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idBtn_Back"]'))).click()
        time.sleep(5)
        logging.info('Accessing search box')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "searchBoxId-Mail"))).click()
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH,"//span[@id='searchScopeButtonId-option']"))).click()
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH,"//body/div[@data-portal-element='true']/div/div/div/div/div/div[@aria-label='Search Scope Selector.']/button[2]/span[1]"))).click()
        time.sleep(5)
        logging.info('Clearing Search Bar')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input'))).clear()
        # driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input').clear()
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//button[@id='filtersButtonId']//span[@data-automationid='splitbuttonprimary']"))).click()
        time.sleep(5)
        field=WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID,'From-PICKER-ID')))
        field.click()
        field.clear()
        field.send_keys('manan')
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Manan Ahuja (manan.ahuja@biourja.com)']//span[@data-automationid='splitbuttonprimary']"))).click()
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Search']//span[@data-automationid='splitbuttonprimary']"))).click()
        logging.info('Clicking recent mail')
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/div[2]/div[2]/div/div/div/div[3]/div[2]/div/div[1]/div[2]/div/div/div/div/div/div/div'))).click()
        logging.info('Clicking more action button')
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//button[@title='More actions']"))).click()
        time.sleep(5)
        logging.info('Clicking download button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//button[@name='Download']"))).click()
        time.sleep(20)
        driver.close()
    except Exception as e:
        raise e 
def trim_process_csv(structure_name,file_name:str):
    try:
        input_sheet=os.getcwd()+f'\\{file_name}' 
        # logger.info("Opening operating workbook instance of excel")
        retry=0
        while retry < 10:
            try:
                wb=xw.Book(input_sheet)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e
        ws1=wb.sheets[0]
        ws1.autofit()
        ws1.api.Range("A1").EntireColumn.Delete()
        ws1.range("1:1").api.Delete()
        list1=["CONTRACT","OPTION_EXPIRY DATE","TRADE_DAYS_TO_EXPIRY","FUTURES_PRICE","ATM_STRADDLE","BREAK_EVEN","ATM_IMP_VOL","ATM_IMP_VOL_1D_CHG","ATM_IMP_VOL_1WK_CHG","ATM_IMP_VOL_1MO_CHG","REAL_VOL_AVG_HIST","REAL_VOL_30_DAY","IMP_VOL_10D_PUT","IMP_VOL_25D_PUT","IMP_VOL_ATM","IMP_VOL_25D_CALL","IMP_VOL_10D_CALL"]
        column_list = ws1.range("A1").expand('right').value
        for index2,values in enumerate(column_list):
            values_column_no=column_list.index(values)+1
            values_letter_column = num_to_col_letters(column_list.index(values)+1)
            ws1.range(f"{values_letter_column}1").value = list1[index2]                  
        # print("Pause")
        
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        ws1.api.Range("A1").EntireColumn.Insert()
        ws1.range("A1").value="STRUCTURE_NAME"
        ws1.range(f"A2:A{last_row}").value=structure_name
        ws1.api.Range("A1").EntireColumn.Insert()
        ws1.range("A1").value="TRADE_DATE"
        ws1.range(f"A2:A{last_row}").value=Trade_date
        column_list = ws1.range("A1").expand('right').value
        insert_letter=num_to_col_letters(len(column_list)+1)
        ws1.range(f"{insert_letter}1").value="INSERTDATE"
        ws1.range(f"{insert_letter}2:{insert_letter}{last_row}").value=today_date
        insert_letter=num_to_col_letters(len(column_list)+2)
        ws1.range(f"{insert_letter}1").value="UPDATEDATE"
        ws1.api.Range("S:S").Cut()
        ws1.api.Range("R:R").Insert(win32c.Direction.xlToRight)
        ws1.api.Range("I:I").EntireColumn.Delete()
        ws1.autofit()
        wb.save(f"{output_location}\\{file_name}")
        try:
            wb.app.quit()
        except Exception as e:
            wb.app.quit()          
    except Exception as e:
                    pass          
def read_pdf():
    try:
        column_values0=["82","114","151","181","213","242","273","296","327","357","391","420","450","481","509","539"]
        test_area0 = ["293.378,50,526.173,569.543"]
        column_values1=["82","114","151","181","213","242","272","296","327","357","391","420","450","481","509","539"]
        test_area1=["282.996","50.0","506.376","567.905"]
        test_area2=["530.0","50.0","584","568"]
        column_values2=["82","114","151","181","213","242","272","296","327","357","391","420","450","481","509","539"]
        df = tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages="all",silent=True,guess=False)
        for index, page in enumerate(df):
            if index == 0:
                test_area_structure = ["291,48,304,569"]
                df_structure = tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages="1",area=test_area_structure,silent=True,guess=False)
                structure_name=df_structure[0].columns[0]
                df0=tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages=[index + 1],area=test_area0,columns=column_values0,silent=True,guess=False)    
                file_name='file.csv'
                df0[0].to_csv(file_name)
                trim_process_csv(structure_name,file_name)
                #for second table
                test_area_structure2 = ["524.408,37.026,542,588.591"]
                df_structure2 = tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages="1",area=test_area_structure2,silent=True,guess=False)
                structure_name=df_structure2[0].columns[0]
                df1=tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages=[index + 1],area=test_area2,columns=column_values2,silent=True,guess=False)    
                file_name='file1.csv'
                df1[0].to_csv(file_name)
                # df1[0].drop(0,inplace=True)
                trim_process_csv(structure_name,file_name)
            if index == 1:
                test_area_structure3 = ["279.608,42.381,294.908,579.411"]
                df_structure3 = tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages=[index+1],area=test_area_structure3,silent=True,guess=False)
                structure_name=df_structure3[0].columns[0]
                df2=tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages=[index + 1],area=test_area1,columns=column_values1,silent=True,guess=False)    
                file_name='file2.csv'
                df2[0].to_csv(file_name)
                trim_process_csv(structure_name,file_name)
                break
        print("Done")
    except Exception as e:
        print(e) 
def csv_to_dataframe():
    try:
        csvsToUpload = os.listdir(f"{output_location}")
        df = pd.DataFrame()
        for files in csvsToUpload:
            data = pd.read_csv (f"{output_location}\\{files}")   
            df1 = pd.DataFrame(data)
            df = df.append(df1, ignore_index=True)
        return df    
    except Exception as e:
        print(e)
def snowflake_dump(df):
    engine = bu_snowflake.get_engine(
            username= "YASH",
            password= "Yash22",
            role= "POWERDB_DEV",
            schema= SCHEMA,
            database= Database
            )
    try:
        query = f"""select * from "POWERDB_DEV"."PMACRO"."TRIDENT_EOD_DAILY_VOLATILITY" where                    
        TRADE_DATE = '{Trade_date}' and STRUCTURE_NAME = '{df["STRUCTURE_NAME"][0]}' and CONTRACT = '{df["CONTRACT"][0]}' and OPTION_EXPIRY = '{df["OPTION_EXPIRY"][0]}'"""            
        
        with engine.connect() as con:
            db_df=engine.execute(query)
            if len(db_df)>0:
                pass
            else:
                df.to_sql('TRIDENT_EOD_DAILY_VOLATILITY', con=con,if_exists='append',index = False)
    except Exception as e:
        logger.exception(f"error occurred : {e}")
    finally:
        engine.dispose()
def main():
    try:
        remove_existing_files(files_location)
        login_and_download()
        read_pdf()
        df=csv_to_dataframe()
        snowflake_dump(df)       
        locations_list.append(logfile)
        send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = f'{job_name} completed successfully, Attached PDF and Logs',attachment_locations = locations_list)
    except Exception as e:
        logging.exception(str(e))
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed, Attached logs',attachment_location = logfile)
               
if __name__ == "__main__": 
    logging.info("Execution Started")
    time_start=time.time()
    directories_created=["Download","Logs","Generated_CSV"]
    for directory in directories_created:
        path3 = os.path.join(os.getcwd(),directory)  
        try:
            os.makedirs(path3, exist_ok = True)
            print("Directory '%s' created successfully" % directory)
        except OSError as error:
            print("Directory '%s' can not be created" % directory)       
    files_location=os.getcwd() + "\\Download"
    filesToUpload = os.listdir(os.getcwd() + "\\Download")
    # share_point_path = credential_dict['API_KEY'].split('/')[4:]
    
    # receiver_email='yashn.jain@biourja.com'
    job_name='TRIDENT_EOD_DAILY_VOLATILITY_AUTOMATION'
    main()
    time_end=time.time()
    logging.info(f'It takes {time_start-time_end} seconds to run')