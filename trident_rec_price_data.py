from more_itertools import strip
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
import time
from datetime import datetime,date
from datetime import timedelta
import logging
from selenium.webdriver.support import expected_conditions as EC
import os
import re
import bu_config
from bu_config import get_config
import bu_alerts
import tabula
import bu_snowflake
import pandas as pd
from snowflake.connector.pandas_tools import pd_writer
import functools
import numpy as np
from selenium.webdriver.firefox.options import Options
#from download_attachment import login_and_download


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
            logger.info("No existing files available to remove")
            print("No existing files available to remove")
    except Exception as e:
        logger.exception(e)
        raise e
def login_and_download(username, password, url, subject, download_path, logger):  
    '''This function downloads log in to the website'''
    try:
        logger.info("Inside login and download function")
        exe_path = os.getcwd() + '\\geckodriver.exe'
        options = Options()
        profile = webdriver.FirefoxProfile()
        profile.set_preference('browser.download.folderList', 2)
        profile.set_preference('browser.download.dir', download_path)
        profile.set_preference('browser.download.useDownloadDir', True)
        profile.set_preference('browser.download.viewableInternally.enabledTypes', "")
        profile.set_preference('browser.helperApps.neverAsk.saveToDisk','Portable Document Format (PDF), application/pdf')
        profile.set_preference('pdfjs.disabled', True)
        profile.update_preferences()
        driver = webdriver.Firefox(firefox_profile=profile,options=options, executable_path=exe_path)
        logger.info('Accesing website')
        driver.get(url)
        driver.maximize_window()
        logger.info('providing id and passwords')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "i0116"))).send_keys(username)
        time.sleep(1)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        time.sleep(1)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys(password)
        time.sleep(1)
        logger.info('click on No Button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idBtn_Back"]'))).click()
        time.sleep(10)       
        retry=0
        while retry < 10:
            try:
                logger.info('Accessing search box')
                WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "searchBoxId-Mail"))).click()
                time.sleep(5)
                logger.info("setting search for only inbox")
                WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable(
                    (By.XPATH,"//span[@id='searchScopeButtonId-option']"))).click()
                time.sleep(10)
                try:
                    WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable(
                        (By.XPATH,"//span[@data-automationid='splitbuttonprimary']//span[contains(text(),'Inbox')]"))).click()
                except:
                    time.sleep(1)
                    WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable(
                        (By.XPATH,"(//span[@class='ms-Button-flexContainer flexContainer-204'])[2]"))).click()                
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e
        time.sleep(5)
        logger.info('Clearing Search Bar')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR,"input[placeholder='Search']"))).clear()
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR,"input[placeholder='Search']"))).send_keys(subject)
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable(
            (By.XPATH, "//button[@aria-label='Search']//span[@data-automationid='splitbuttonprimary']"))).click()
        logger.info('Clicking recent mail')
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable(
            (By.XPATH, """/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/
            div/div/div/div[6]/div/div/div[1]/div[2]"""))).click()        
        logger.info('Clicking more action button')
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//button[@title='More actions']"))).click()
        time.sleep(5)
        logger.info('Clicking download button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//button[@name='Download']"))).click()
        time.sleep(20)
        try:
            driver.close()
        except Exception as e: 
            logger.exception('driver not closed')
            print("driver not closed") 
            try:
                driver.quit()
            except Exception as e: 
                logger.exception('driver quit failed')
                print("driver quit failed")
    except Exception as e:
        logger.exception(e)
        raise e 
def extract_and_upload_pdf(download_path):

    try:
        logger.info("Inside extract_and_upload_pdf function")
        file_name= os.listdir(download_path)
        rows = 0
        for file in file_name:
            if '.pdf' in file:            #checking the downloaded file is the recent one or not.Else, required file was not yet sent to mail.
                logger.info(f"Process started for file {file} from pdf")
                # Taking iso name from the 1st table in the pdf
                iso_name_1= tabula.read_pdf(download_path + '\\' + file, pages = 1,area =["109.013,49.725,122.018,544.68"])[0].columns[0]
                logger.info("Taking the areas of iso_name_1")
                # Taking the area co-ordinates(individually) for the tables specified in 1st iso. 
                areas_1= [["119.733,47.708,177.873,172.403"],["122.028,174.698,176.343,294.038"],["118.958,296.601,177.863,419.001"],["121.263,425.618,178.638,545.723"],["177.097,44.916,234.472,170.376"],\
                              ["177.097,174.201,234.472,293.541"],["177.097,297.366,234.472,418.236"],["175.567,421.296,234.472,542.166"],["233.707,49.506,291.847,165.021"],["232.942,171.141,291.847,295.071"],\
                                ["232.942,299.661,291.082,418.236"],["232.942,422.826,291.082,550.581"],["289.552,422.061,350.752,547.521"],["290.317,49.506,349.987,167.316"],["291.082,171.906,352.282,292.776"],\
                                   ["290.317,297.366,351.517,419.001"]]  
                list_dfs=[]
                for table in areas_1:
                    #using co_ordinates,forming them into table using below method and making changes in it to get the required date frame. 
                    table_1_df= tabula.read_pdf(download_path + '\\' + file, pages = 1,area =table)[0]
                    contract_name = table_1_df.columns[1]
                    table_1_df.columns=['STRIP','BID','ASK']
                    temp_df=table_1_df.drop(index=0)
                    temp_df['CONTRACT_NAME']=contract_name
                    temp_df = temp_df[['CONTRACT_NAME','STRIP','BID','ASK']] 
                    list_dfs.append(temp_df)              
                #concatinating all the dataframes to get a single dataframe
                iso_1_df = pd.concat(list_dfs,ignore_index=True)
                iso_1_df["ISO"]=iso_name_1
                trade_date = file.split(' ')[-1].split('.')[0]
                trade_date = datetime.strptime(trade_date,'%d%b%Y')
                temp_date = trade_date.strftime("%Y-%m-%d")
                iso_1_df["TRADE_DATE"]= temp_date
                iso_1_df['INSERT_DATE'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                final_df_1=iso_1_df[['TRADE_DATE','ISO','CONTRACT_NAME','STRIP','BID','ASK','INSERT_DATE']]
                print(final_df_1)
                # Taking iso name from the 2nd table in the pdf               
                iso_name_2= tabula.read_pdf(download_path + '\\' + file, pages = 1,area =["357.638,50.49,367.583,542.385"])[0].columns[0]
                logger.info("Taking the areas of iso_name_2")
                # Taking the area co-ordinates(individually) for the tables specified in 2nd iso.
                areas_2= [["366.053,40.326,429.548,174.201"],["366.053,171.141,430.313,294.306"],["367.583,296.601,415.013,417.471"],["366.818,419.766,415.778,543.696"],["433.373,47.976,471.623,168.846"],\
                    ["432.608,173.436,471.623,294.306"],["415.013,298.896,453.263,412.881"],["412.718,419.766,453.263,543.696"],["472.388,44.916,509.873,169.611"],\
                        ["470.858,171.906,510.638,292.011"],["451.733,292.776,491.513,413.646"],["450.203,416.706,500.693,541.401"],["509.108,45.681,569.543,166.551"],\
                            ["508.343,173.436,571.073,292.776"],["526.703,295.836,571.838,417.471"],["489.983,295.071,528.998,415.941"],["498.398,420.531,552.713,542.931"]]
                list_dfs_2=[]
                for table2 in areas_2:
                    try:
                        table_2_df= tabula.read_pdf(download_path + '\\' + file, pages = 1,area =table2)[0]
                        contract_name = table_2_df.columns[1]
                        table_2_df.columns=['STRIP','BID','ASK']
                        temp_df_2=table_2_df.drop(index=0)
                        temp_df_2['CONTRACT_NAME']=contract_name
                        temp_df_2 = temp_df_2[['CONTRACT_NAME','STRIP','BID','ASK']] 
                        list_dfs_2.append(temp_df_2)
                    except ValueError as e:
                        print(e,f'For Area with {table2}')
                        #error getting raised that one area has only two columns but the requirement has three.
                        #BID and ASK are two diff columns. But they are coming in single column. 
                        #forming table with that area and, splitting that column(BID&ASK) into two seperate columns(BID,ASK)
                        ta=tabula.read_pdf(download_path + '\\' + file, pages = 1,area =table2)[0]
                        contract_new= ta.columns[1]
                        bid =ta.iloc[:,1].apply(lambda x: x.split(' ')[0])
                        ask = ta.iloc[:,1].apply(lambda x: x.split(' ')[1])
                        ta.drop(ta.columns[1],axis = 1,inplace = True)
                        ta.insert(1,'BID',bid)
                        ta.insert(2,'ASK',ask)
                        contract_name = ta.columns[1]
                        ta.columns=['STRIP','BID','ASK']
                        temp_df_ta=ta.drop(index=0)
                        temp_df_ta['CONTRACT_NAME']=contract_new
                        temp_df_ta = temp_df_ta[['CONTRACT_NAME','STRIP','BID','ASK']] 
                        list_dfs_2.append(temp_df_ta)
                    except Exception as e:
                        logger.exception(e)
                        raise e
                #concatinating all the dataframes to get a single dataframe                 
                iso_2_df = pd.concat(list_dfs_2,ignore_index=True)
                iso_2_df["ISO"]=iso_name_2
                iso_2_df["TRADE_DATE"]= temp_date
                iso_2_df["INSERT_DATE"]=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                final_df_2=iso_2_df[['TRADE_DATE','ISO','CONTRACT_NAME','STRIP','BID','ASK','INSERT_DATE']]
                print(final_df_2)
                # Taking iso name from the 3rd table in the pdf
                iso_name_3=tabula.read_pdf(download_path + '\\' + file, pages = 1,area =["575.663,50.49,586.373,545.445"])[0].columns[0]
                logger.info("Taking the areas of iso_name_3")
                # Taking the area co-ordinates(individually) for the tables specified in 3rd iso.
                areas_3= [["584.843,50.49,632.273,172.89"],["584.843,172.584,629.978,294.219"],["584.078,297.279,634.568,420.444"]]
                list_dfs_3=[]
                for table3 in areas_3:
                    table_3_df= tabula.read_pdf(download_path + '\\' + file, pages = 1,area =table3)[0]
                    contract_name=table_3_df.columns[1]
                    #same problem is repeated in this table ,as BID & ASK came in single column.
                    #here all tables in this iso are with same problem
                    #split the column into two columns and add into the table with three columns
                    table_3_df.columns=['STRIP','BID&ASK']
                    temp_df_3=table_3_df.drop(index=0)
                    temp_df_3['CONTRACT_NAME']=contract_name
                    temp_df_3 = temp_df_3[['CONTRACT_NAME','STRIP','BID&ASK']] 
                    list_dfs_3.append(temp_df_3)              
                iso_3_df = pd.concat(list_dfs_3,ignore_index=True)
                iso_3_df["ISO"]=iso_name_3
                iso_3_df["TRADE_DATE"]= temp_date
                iso_3_df["INSERT_DATE"]=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                final_df_3=iso_3_df[['TRADE_DATE','ISO','CONTRACT_NAME','STRIP','BID&ASK','INSERT_DATE']]
                bid = final_df_3.iloc[:,4].apply(lambda x: x.split(' ')[0])
                ask = final_df_3.iloc[:,4].apply(lambda x: x.split(' ')[1])
                final_df_3.drop(final_df_3.columns[4],axis = 1,inplace = True)
                final_df_3.insert(4,'BID',bid)
                final_df_3.insert(4,'ASK',ask)
                print(final_df_3)
                # Taking iso name from the 4th table in the pdf
                iso_name_4 =tabula.read_pdf(download_path + '\\' + file, pages = 1,area =["575.663,50.49,586.373,545.445"])[0].columns[1]
                logger.info("Taking the areas of iso_name_4")
                areas_4 = ["584.843,422.739,642.983,547.434"]
                table_4_df= tabula.read_pdf(download_path + '\\' + file, pages = 1,area =areas_4)[0]
                contract_name = table_4_df.columns[1]
                table_4_df.columns=['STRIP','BID','ASK']
                temp_df_4=table_4_df.drop(index=0)
                temp_df_4['CONTRACT_NAME']=contract_name
                iso_4_df = temp_df_4[['CONTRACT_NAME','STRIP','BID','ASK']] 
                iso_4_df["ISO"]=iso_name_4
                iso_4_df["TRADE_DATE"]= temp_date
                iso_4_df["INSERT_DATE"]=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                final_df_4=iso_4_df[['TRADE_DATE','ISO','CONTRACT_NAME','STRIP','BID','ASK','INSERT_DATE']]
                print(final_df_4)
                #concating all the final df's into one main df.
                main_df= pd.concat([final_df_1,final_df_2,final_df_3,final_df_4],ignore_index=True)    
                #NO need of Dollar sign in sf.
                #Applying replace to take it off.
                main_df['BID'] = main_df['BID'].replace({r'\$':''}, regex = True)
                main_df['ASK'] = main_df['ASK'].replace({r'\$':''}, regex = True)
                logger.info("DateFrame with all details got ready")
                print(main_df)
            else:
                print(f"Recent file with Trade_Date still not came")
                break
            rows += upload_in_sf(main_df, temp_date)
            logger.info("Upload sf function completed")
            
    except Exception as e:
        logger.exception(e)
        raise e
    finally:
        return rows
        

def upload_in_sf(df, trade_date):
    
    logger.info("Inside upload_in_sf function")
    total_rows = 0
    try:
        engine = bu_snowflake.get_engine(
                    database=databasename,
                    role=f"OWNER_{databasename}",    
                    schema= schemaname                           
                )
        conn = engine.connect()
        logger.info("Engine object created successfully")

        check_query = f"select * from {databasename}.{schemaname}.{tablename} where TRADE_DATE = '{trade_date}'"
        check_max_date = f"select MAX(TRADE_DATE) from {databasename}.{schemaname}.{tablename}"
        max_date =conn.execute(check_max_date).fetchall()
        check_rows = conn.execute(check_query).fetchall()
        logger.info("Check if the data is already present")
        sf_max_date =max_date[0][0].strftime("%Y-%m-%d")
        if trade_date > sf_max_date:
            if len(check_rows) == 0:
                logger.info(f"NO data of {trade_date} is present in {tablename} table.Dumping got started into table")
                print(f'NO Data existed for {trade_date}')
                df.to_sql(tablename.lower(), 
                    con=engine,
                    index=False,
                    if_exists='append',
                    schema=schemaname,
                    method=functools.partial(pd_writer, quote_identifiers=False)
                    )
                print((f"values added for {trade_date}"))   
                logger.info(f"values of {trade_date} file are added into {tablename} table") 
            elif len(check_rows)> 0:
                logger.info(f"data of {trade_date} file is already present in {tablename} table")
                print(f'Data already existed for {trade_date}')

            logger.info(f"Dataframe Inserted into the table {tablename} for TRADEDATE {trade_date} and total rows are {len(df)}")
            total_rows += len(df)
        elif trade_date <= sf_max_date:
            logger.info(f"New pdf file still not received.Last dumped data is of {sf_max_date} file")
            print("New pdf file not received")    
    except Exception as e:
        logger.exception("Exception while inserting data into snowflake")
        logger.exception(e)
        raise e
    finally:
        try:        
            conn.close()      
            engine.dispose()
            logger.info("Engine object disposed successfully and connection object closed")
            return total_rows
        except Exception as e:
            logger.exception(e)
            raise e

if __name__ == '__main__':
    try:
        job_id=np.random.randint(1000000,9999999)
        logger = logging.getLogger(__name__)
        logger.setLevel(logging.INFO)
        logfilename = bu_alerts.add_file_logging(logger,process_name= 'TRIDENT_REC_PRICE_DATA')  
        credential_dict = get_config('TRIDENT_REC_PRICE_DATA','EMISSION_TRIDENT_REC_PRICES')
        processname = credential_dict['PROJECT_NAME']
        tablename = credential_dict['TABLE_NAME']
        url = credential_dict['SOURCE_URL']
        username = credential_dict['USERNAME']
        password = credential_dict['PASSWORD']
        subject = credential_dict['API_KEY']
        databasename = 'POWERDB_DEV'  #credential_dict['DATABASE']
        schemaname = credential_dict['TABLE_SCHEMA']
        process_owner =  credential_dict['IT_OWNER']
        receiver_email = 'indiapowerit@biourja.com'#'enoch.benjamin@biourja.com' #credential_dict['EMAIL_LIST']
        exe_path = os.getcwd() + '\\geckodriver.exe'
        download_path = os.getcwd() + "\\"+"download"
        log_json='[{"JOB_ID": "'+str(job_id)+'","CURRENT_DATETIME": "'+str(datetime.now())+'"}]'
        directories_created=["download","logs"]
        for directory in directories_created:
            path3 = os.path.join(os.getcwd(),directory)  
            try:
                os.makedirs(path3, exist_ok = True)
                print("Directory '%s' created successfully" % directory)
            except OSError as error:
                print("Directory '%s' can not be created" % directory) 

        bu_alerts.bulog(process_name=processname,database=databasename,status='Started',table_name=tablename,
            row_count=0, log=log_json, warehouse='ITPYTHON_WH',process_owner=process_owner)
        
        logger.info("Calling remove existing file function")
        remove_existing_files(download_path)
        logger.info("Remove existing file completed")
        logger.info("Calling login download function")
        login_and_download(username, password, url, subject, download_path, logger)
        logger.info("Login and download function completed")
        logger.info("Calling extract df function")
        rows = extract_and_upload_pdf(download_path)
        logger.info("Extract df function completed")
        
        bu_alerts.bulog(process_name=processname,database=databasename,status='Completed',table_name=tablename,
            row_count=rows, log=log_json, warehouse='ITPYTHON_WH',process_owner=process_owner)

        if rows > 0:
            subject = f"JOB SUCCESS - {tablename}  inserted {rows} rows"
        else:
            subject = f"JOB SUCCESS - {tablename} New Recent pdf file still not received."

        bu_alerts.send_mail(
            receiver_email = receiver_email, 
            mail_subject = subject,
            mail_body=f'{tablename} completed successfully, Attached logs',
            attachment_location = logfilename
        )

    except Exception as e:
        print("Exception caught during execution: ",e)
        logging.exception(f'Exception caught during execution: {e}')
        log_json='[{"JOB_ID": "'+str(job_id)+'","CURRENT_DATETIME": "'+str(datetime.now())+'"}]'
        bu_alerts.bulog(process_name= processname,database=databasename,status='Failed',table_name=tablename,
            row_count=rows, log=log_json, warehouse='ITPYTHON_WH',process_owner=process_owner)

        bu_alerts.send_mail(
            receiver_email = receiver_email,
            mail_subject = f'JOB FAILED - {tablename}',
            mail_body=f'{tablename} failed during execution, Attached logs',
            attachment_location = logfilename
        )