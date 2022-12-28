from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
import time
import logging
from selenium.webdriver.support import expected_conditions as EC
import os
from bu_config import get_config
import bu_alerts
import tabula
import bu_snowflake
import pandas as pd
from snowflake.connector.pandas_tools import pd_writer
import functools
from datetime import date, datetime
import numpy as np
from selenium.webdriver.firefox.options import Options
from webdriver_manager.firefox import GeckoDriverManager


def convert_float(s):
    '''
        This function converts data frame series values from number string to float. In case if error we return np.nan for strings other

        Params:
        -------
        s : str
            The value of each row of a dataframe column

        Returns:
        --------
        float(s): float
            The float value of the number string
    '''
    try:
        return float(s)
    except ValueError as e:
        return np.nan 

def trade_date():
    try:
        file_name= os.listdir(os.getcwd() + "\\Download")
        file2=file_name[0]
        test_area_date = ["67,47,85,160"]
        df_date = tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages="1",area=test_area_date,silent=True,guess=False)
        Trade_date=df_date[0].columns[1]
        return Trade_date
    except Exception as e:
        logger.info(e)
        raise e


def num_to_col_letters(num):
    try:
        letters = ''
        while num:
            mod = (num - 1) % 26
            letters += chr(mod + 65)
            num = (num - 1) // 26
        return ''.join(reversed(letters))
    except Exception as e:
        logger.info(e)
        raise e


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
        options = Options()
        profile = webdriver.FirefoxProfile()
        profile.set_preference('browser.download.folderList', 2)
        profile.set_preference('browser.download.dir', path)
        profile.set_preference('browser.download.useDownloadDir', True)
        profile.set_preference('browser.download.viewableInternally.enabledTypes', "")
        profile.set_preference('browser.helperApps.neverAsk.saveToDisk','Portable Document Format (PDF), application/pdf')
        profile.set_preference('pdfjs.disabled', True)
        profile.update_preferences()
        logging.info('Adding firefox profile')
        exe_path = r'S:\IT Dev\Production_Environment\trident_eod_daily_volatility-1\geckodriver.exe'
        # exe_path = r'C:\Users\Yashn.jain\OneDrive - BioUrja Trading LLC\Power\trident_eod_daily_volatility'
        # driver=webdriver.Firefox(executable_path=exe_path,firefox_profile=profile)
        driver = webdriver.Firefox(firefox_profile=profile,options=options, executable_path=GeckoDriverManager().install())
        logging.info('Accesing website')
        driver.get("https://outlook.office365.com/owa/biourja.com/")
        time.sleep(1)
        driver.maximize_window()
        time.sleep(10)
        logging.info('providing id and passwords')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "i0116"))).send_keys(username)
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys(password)
        time.sleep(5)
        logging.info('click on No Button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idBtn_Back"]'))).click()
        time.sleep(10)       
        retry=0
        while retry < 10:
            try:
                # WebDriverWait(driver, 90, poll_frequency=1).until(EC.invisibility_of_element_located((By.CLASS_NAME, "ms-Overlay ms-Overlay--dark root-256")))
                logging.info('closing unwanted overlay')
                try:
                    WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".ms-Dialog-button"))).click()
                except:
                    pass
                time.sleep(5)
                logging.info('Accessing search box')
                WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "searchBoxId-Mail"))).click()
                time.sleep(5)
                logging.info("setting search for only inbox")
                WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH,"//span[@id='searchScopeButtonId-option']"))).click()
                time.sleep(10)
                WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH,"//span[@data-automationid='splitbuttonprimary']//span[contains(text(),'Inbox')]"))).click()
                break                   
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e  
        time.sleep(5)
        logging.info('Clearing Search Bar')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR,"input[placeholder='Search']"))).clear()        # driver.find_element_by_xpath('/html/body/div[3]/div/div[1]/div/div[1]/div[2]/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div[2]/div/input').clear()
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR,"input[placeholder='Search']"))).send_keys('from:"manan.ahuja@biourja.com" AND Vol Report EOD')
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Search']//span[@data-automationid='splitbuttonprimary']"))).click()
        logging.info('Clicking recent mail')
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div/div/div/div[6]/div/div"))).click()        
        logging.info('Clicking more action button')
        time.sleep(5)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//button[@title='More actions']"))).click()
        time.sleep(5)
        logging.info('Clicking download button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//button[@name='Download']"))).click()
        time.sleep(20)
        try:
            driver.close()
        except Exception as e: 
            logging.info('driver not closed')
            print("driver not closed") 
            try:
                driver.quit()
            except Exception as e: 
                logging.info('driver quit failed')
                print("driver quit failed")
    except Exception as e:
        raise e 


def refactoring_dataframe(Trade_date,structure_name,dataframe):
    try:
        colList = list(dataframe.columns)
        for col in range(len(colList)):
            if len(list(dataframe.loc[dataframe[colList[col]]=='Contract'].index)):
                contactCol = col
                contractIndex = dataframe.loc[dataframe[colList[col]]=='Contract'].index[0]

        for i in range(0,contractIndex):
            dataframe.drop(i,inplace=True)
        for i in range(0,contactCol):
            dataframe.drop(dataframe.columns[[i]], axis=1, inplace=True)
        dataframe.dropna(axis=0,how='all',inplace=True)
        dataframe.dropna(axis=1,how='all',inplace=True)      
        dataframe.drop(1,inplace=True)
        dataframe.columns = ["CONTRACT","OPTION_EXPIRY","TRADE_DAYS_TO_EXPIRY","FUTURES_PRICE","ATM_STRADDLE","BREAK_EVEN",
                        "ATM_IMP_VOL","ATM_IMP_VOL_1D_CHG","ATM_IMP_VOL_1WK_CHG","ATM_IMP_VOL_1MO_CHG","REAL_VOL_AVG_HIST","REAL_VOL_30_DAY",
                                    "IMP_VOL_10D_PUT","IMP_VOL_25D_PUT","IMP_VOL_ATM","IMP_VOL_25D_CALL","IMP_VOL_10D_CALL"]

        columns_titles = ["CONTRACT","OPTION_EXPIRY","TRADE_DAYS_TO_EXPIRY","FUTURES_PRICE","ATM_STRADDLE",
                            "BREAK_EVEN","ATM_IMP_VOL","ATM_IMP_VOL_1D_CHG","ATM_IMP_VOL_1WK_CHG","ATM_IMP_VOL_1MO_CHG","REAL_VOL_AVG_HIST", "REAL_VOL_30_DAY",
                                "IMP_VOL_10D_PUT","IMP_VOL_25D_PUT","IMP_VOL_ATM","IMP_VOL_10D_CALL","IMP_VOL_25D_CALL"]
        dataframe=dataframe.reindex(columns=columns_titles)
        dataframe.insert(0,'STRUCTURE_NAME',structure_name) 
        dataframe.insert(0,'TRADE_DATE',Trade_date)
        dataframe['INSERTDATE'] = str(datetime.now())
        dataframe['UPDATEDATE'] = str(datetime.now())
        return dataframe
    except Exception as e:
        logger.exception(f"error occurred : {e}")
        raise(e)  


def read_pdf(Trade_date):
    try:
        file_name= os.listdir(os.getcwd() + "\\Download")
        file2=file_name[0] 
        logger.info("testing areas and column seperator values")
        #new coordinates updated(1st table)
        column_values0=["76","107","141","172","201","231","259","282.5","312","341","373","402","431","459","487","517"]
        test_area0 = ["273.488,43.911,515.993,547.281"]
        #old coordinates
        # column_values0=["82","114","151","181","213","242","273","296","327","357","391","420","450","481","509","539"]
        # test_area0 = ["293.378,50,526.173,569.543"]
        #new coordinates updated
        test_area2=["506.813,43.911,571.838,545.751"]
        column_values2=["76","107","141","172","201","231","259","282.5","312","341","373","402","431","459","487","517"]
        #old coordinates
        # test_area2=["530.0","50.0","584","568"]
        # column_values2=["82","114","151","181","213","242","272","296","327","357","391","420","450","481","509","539"]    
        #new coordinates updated
        column_values1=["76","107","141","172","201","231","259","282.5","312","341","373","402","431","459","487","517"]
        test_area1=["257.423,-0.459,497.633,603.891"]
        #old coordinates
        # column_values1=["82","114","151","181","213","242","272","296","327","357","391","420","450","481","509","539"]
        # test_area1=["282.996","50.0","508","568"]

        logger.info("reading full page tables")
        df = tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages="all",silent=True,guess=False)
        for index, page in enumerate(df):
            logger.info("applying check for index and extracting tables")
            if index == 0:
                logger.info("picking up structure name from the table")
                test_area_structure = ["274.253,1.071,295.673,606.951"] #changed
                df_structure = tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages="1",area=test_area_structure,silent=True,guess=False)
                structure_name=df_structure[0].columns[0]
                logger.info("picking up table and converting it to csv")
                df0=tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages=[index + 1],area=test_area0,columns=column_values0,silent=True,guess=False)[0]  
                logger.info("trimming the table")
                df0=refactoring_dataframe(Trade_date,structure_name,df0)
                #for second table
                logger.info("picking up structure name from the table")
                test_area_structure2 = ["506.813,0.306,526.703,611.541"] #changed
                df_structure2 = tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages="1",area=test_area_structure2,silent=True,guess=False)
                structure_name=df_structure2[0].columns[0]
                logger.info("picking up table and converting it to csv")
                df1=tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages=[index + 1],area=test_area2,columns=column_values2,silent=True,guess=False)[0]    
                logger.info("trimming the table")
                df1=refactoring_dataframe(Trade_date,structure_name,df1)
            if index == 1:
                logger.info("picking up structure name from the table")
                test_area_structure3 = ["250.538,1.071,278.843,611.541"]#changed
                df_structure3 = tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages=[index+1],area=test_area_structure3,silent=True,guess=False)
                structure_name=df_structure3[0].columns[0]
                logger.info("picking up table and converting it to csv")
                df2=tabula.read_pdf(download_path + '\\' + file2,lattice=True,stream=True, multiple_tables=True,pages=[index + 1],area=test_area1,columns=column_values1,silent=True,guess=False)[0]    
                logger.info("trimming the table")
                df2=refactoring_dataframe(Trade_date,structure_name,df2)
                break
        return df0,df1,df2
    except Exception as e:
        logger.info(e)
        print(e) 

def csv_to_dataframe(dataframe1,dataframe2,dataframe3):
    ''' convert into csv file to dataframe'''
    
    try:
        logger.info("into csv_to_dataframe")
        # csvsToUpload = os.listdir(f"{output_location}")
        # logger.info("creating empty dataframe")
        df = pd.DataFrame()
        logger.info("appending individual dataframes to single empty dataframe")
        # for files in csvsToUpload:
        #     data = pd.read_csv (f"{output_location}\\{files}")   
        #     df1 = pd.DataFrame(data)
        df = df.append([dataframe1,dataframe2,dataframe3], ignore_index=True)
        logger.info("deleting non required column")
        del df['ATM_IMP_VOL']
        logger.info("applying various operations on dataframe")
        list2=["FUTURES_PRICE","ATM_STRADDLE","BREAK_EVEN"]
        for values in list2: 
            df[values]  = [x[values].replace('$', '') for i, x in df.iterrows()]    
            df[values]  = [x[values].replace(' ', '') for i, x in df.iterrows()] 
            df.loc[df[values] == '#DIV/0!',values] = pd.np.nan 
            # df[values] = df[values].astype(float)
            df[values] = df[values].apply(convert_float)
        df.fillna(str("nan"),inplace=True)    
        list3=list(df.columns[8:18])  
        for values in list3:
            # df[values] = df[values].astype(str)
            df[values]  = [x[values].replace('%', '') for i, x in df.iterrows()]    
            df[values]  = [x[values].replace(' ', '') for i, x in df.iterrows()]
            df.loc[df[values] == '#DIV/0!',values] = pd.np.nan  
            df.loc[df[values] == '#REF!',values] = pd.np.nan
            # df[values] = df[values].astype(float)
            df[values] = df[values].apply(convert_float)
            print(values)

        df['TRADE_DAYS_TO_EXPIRY']=df['TRADE_DAYS_TO_EXPIRY'].astype(float)    
        # df['ISO_PNODE']  = [x['ISO_PNODE'].replace('*', '') for i, x in df.iterrows()]    
        # df["INSERTDATE"] = pd.to_datetime(pd.Series(df["INSERTDATE"])).apply(lambda x: datetime.strftime(x, "%Y-%m-%d"))
        # df["UPDATEDATE"] = pd.to_datetime(pd.Series(df["UPDATEDATE"])).apply(lambda x: datetime.strftime(x, "%Y-%m-%d"))
        try:
            # df["INSERTDATE"] = pd.to_datetime(df["INSERTDATE"],format='%m/%d/%Y').astype(str)
            df["TRADE_DATE"] = pd.to_datetime(df["TRADE_DATE"],format='%m/%d/%Y').astype(str)
            # df["UPDATEDATE"] = pd.to_datetime(df["UPDATEDATE"],format='%m/%d/%Y').astype(str)
        except Exception as e:
            logger.exception(f"conversion to datetime for Trade date column failed, {e}")
        df["OPTION_EXPIRY"] = pd.to_datetime(df["OPTION_EXPIRY"],format='%m/%d/%y').astype(str)
        # df["TRADE_DATE"] = pd.to_datetime(pd.Series(df["TRADE_DATE"])).apply(lambda x: datetime.strftime(x, "%Y-%m-%d"))
        # df["OPTION_EXPIRY"] = pd.to_datetime(pd.Series(df["OPTION_EXPIRY"])).apply(lambda x: datetime.strftime(x, "%Y-%m-%d"))
        
        return df    
    except Exception as e:
        logger.exception(f"Error occurred during csv to dataframe conversion {e}")
        raise e


def snowflake_dump(df):
    '''uploaded  the dataframe   in snowflake '''

    logger.info("creating engine object and providing credentials")
    engine = bu_snowflake.get_engine(
            role= f"OWNER_{Database}",
            schema= SCHEMA,
            database= Database
            )
    logger.info("connection initaited")        
    try:
        logger.info("query to check data")
        query = f"select * from {Database}.{SCHEMA}.{table_name} where TRADE_DATE = '{df['TRADE_DATE'][0]}'"       
        logger.info("applying check for values in snowflake table and inserting data")
        with engine.connect() as con:
            db_df = pd.read_sql_query(query, con)
            if len(db_df)>0:
                no_of_rows=0
            else:
                df.to_sql('TRIDENT_EOD_DAILY_VOLATILITY', con=con,if_exists='append',index = False,method=functools.partial(pd_writer, quote_identifiers=False))
                no_of_rows=len(df)
        return no_of_rows 
    except Exception as e:
        logger.exception(f"error occurred : {e}")
        raise(e)
    finally:
        engine.dispose()


def main():
    try:
        no_of_rows=0
        log_json='[{"JOB_ID": "'+str(job_id)+'","CURRENT_DATETIME": "'+str(datetime.now())+'"}]'
        bu_alerts.bulog(process_name=processname,database=Database,status='Started',table_name='',
            row_count=no_of_rows, log=log_json, warehouse='ITPYTHON_WH',process_owner=process_owner)
        logger.info("into remove_existing_files funtion")
        # remove_existing_files(files_location)
        logger.info("into login_and_download")
        login_and_download()
        Trade_date=trade_date()
        logger.info("into read_pdf")
        df0,df1,df2=read_pdf(Trade_date)
        logger.info("into csv_to_dataframe")
        df=csv_to_dataframe(df0,df1,df2)
        logger.info("into snowflake_dump")
        no_of_rows=snowflake_dump(df)  
        logger.info("appending file for mail")
        bu_alerts.bulog(process_name=processname,database=Database,status='Completed',table_name='',
            row_count=no_of_rows, log=log_json, warehouse='ITPYTHON_WH',process_owner=process_owner)     
        # locations_list.append(logfile)
        if no_of_rows>0:
            bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name} and {no_of_rows} rows updated',mail_body = f'{job_name} completed successfully, Attached Logs',attachment_location = logfile)
        else:
            bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name} and Already inserted previously and NO NEW DATA FOUND',mail_body = f'{job_name} completed successfully, Attached Logs',attachment_location = logfile)
    except Exception as e:
        log_json='[{"JOB_ID": "'+str(job_id)+'","CURRENT_DATETIME": "'+str(datetime.now())+'"}]'
        bu_alerts.bulog(process_name= processname,database=Database,status='Failed',table_name='',
            row_count=no_of_rows, log=log_json, warehouse='ITPYTHON_WH',process_owner=process_owner)
        logging.exception(str(e))
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed, Attached logs',attachment_location = logfile)


if __name__ == "__main__": 
    logging.info("Execution Started")
    time_start=time.time()
    today_date=date.today()
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    # log progress --
    logfile = os.getcwd() +"\\logs\\"+'TRIDENT_'+str(today_date)+'.txt'
    logging.basicConfig(level=logging.INFO,filename=logfile,filemode='w',format='[line :- %(lineno)d] %(asctime)s [%(levelname)s] - %(message)s ')

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logging.info('setting paTH TO DOWNLOAD')
    path = os.getcwd() + "\\"+"Download"
    logging.info('SETTING PROFILE SETTINGS FOR FIREFOX')

    credential_dict = get_config('TRIDENT_EOD_DAILY_VOLATILITY','TRIDENT_EOD_DAILY_VOLATILITY')
    username = credential_dict['USERNAME']
    password = credential_dict['PASSWORD']
    table_name = credential_dict['TABLE_NAME']
    Database = credential_dict['DATABASE']
    # Database = "POWERDB_DEV"
    SCHEMA = credential_dict['TABLE_SCHEMA']

    # receiver_email = credential_dict['EMAIL_LIST']
    # receiver_email = "yashn.jain@biourja.com"

    receiver_email = credential_dict['EMAIL_LIST']
    # receiver_email = "megha.chouhan@biourja.com, mrutunjaya.sahoo@biourja.com,radha.waswani@biourja.com"
    download_path=os.getcwd() + "\\Download"
    output_location= os.getcwd()+"\\Generated_CSV"
    today_date=date.today()
    job_id=np.random.randint(1000000,9999999)
    processname = credential_dict['PROJECT_NAME']
    process_owner = credential_dict['IT_OWNER']
    logging.info("Creating required directories")
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
    job_name='TRIDENT_EOD_DAILY_VOLATILITY_AUTOMATION'
    logging.info("Into the main function")
    main()
    time_end=time.time()
    logging.info(f'It takes {time_start-time_end} seconds to run')


