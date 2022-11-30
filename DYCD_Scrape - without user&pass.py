#--------------------------------------------IMPORTS---------------------------------------------------

import DYCD_getattendance as ga #includes all the functions that deal with getting the numbers from downloaded Excels

import time
from datetime import datetime, timedelta
import pandas as pd
import traceback
import chromedriver_autoinstaller

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#--------------------------------------------Functions--------------------------------------------------

def enter_DYCD(): #login and get to reports page
    driver.get("https://www.dycdconnect.nyc/Home/Login")
    assert "DYCD" in driver.title

    #input user and pass
    username = driver.find_element(By.ID, "UserName");
    password = driver.find_element(By.ID, "Password");

    username.send_keys('000')
    password.send_keys('000')

    #login
    buttons = driver.find_elements(By.TAG_NAME, 'button')
    buttons[1].click()

    time.sleep(t*0.5)

    #click on PTS/EMS
    pts = WebDriverWait(driver, t).until(EC.visibility_of_all_elements_located((By.CLASS_NAME, 'btn-group')))
    pts[1].click()

    #switch to the new tab
    driver.switch_to.window(driver.window_handles[1])

    #open menu and click reports
    menu = driver.find_element(By.CLASS_NAME, 'navBarTopLevelItem')
    button = menu.find_element(By.ID, 'TabMainMenu')
    button.click()

    time.sleep(t*0.1)

    nav_groups = driver.find_elements(By.CLASS_NAME, 'nav-subgroup')
    nav_groups[3].click()

    #Now at reports page! --------

    #switch to iframe within reports page
    iframe = driver.find_element(By.XPATH, '//*[@id="contentIFrame0"]')
    driver.switch_to.frame(iframe)

    time.sleep(1)

def separate():
    data.append(['','']) #separate so final excel is cleaner
    time.sleep(1.5)
    switchto_report()
    time.sleep(1.5)

def switchto_report(): #goes from report editor window back to main dycd reports page
    driver.close() #no longer need report window, close it

    driver.switch_to.window(driver.window_handles[1]) # switch to reports tab

    #switch to iframe within reports page
    iframe = driver.find_element(By.XPATH, '//*[@id="contentIFrame0"]')
    driver.switch_to.frame(iframe)

def global_prev(program_area, program_type): #speed things up in TAR by keeping certain elements plugged in
    global prev 
    prev = [program_area, program_type]

def next_page(): #click next page
    element = driver.find_element(By.XPATH, '//*[@id="_nextPageImg"]')
    driver.execute_script("arguments[0].click();", element)

def find_report(report): #while on main reports page, find report, and click on it
    #click tar
    tar = driver.find_element(By.XPATH, report) 
    a = tar.find_element(By.TAG_NAME, 'a')

    actions = ActionChains(driver)
    actions.move_to_element(a).perform()

    time.sleep(1)

    #
    driver.execute_script("arguments[0].click();", a)
    time.sleep(1)

    #Go to report builder pag
    driver.switch_to.window(driver.window_handles[2])
    driver.maximize_window()

    #switch to iframe within report window
    driver.switch_to.frame(0)

def select_element(xpath, n):
    select = Select(driver.find_element(By.XPATH, xpath))
    select.select_by_value(str(n))

def select_workscope_element(xpath, n):
    select = Select(driver.find_element(By.XPATH, xpath))
    select.select_by_value(str(n))

    return select.first_selected_option.text

def fill_dates_TAR(start_date, end_date): #fill in start date
    start = driver.find_element(By.XPATH, '//*[@id="reportViewer_ctl08_ctl18_txtValue"]')
    start.clear()
    time.sleep(1)

    start.send_keys(str(start_date))
    start.send_keys(Keys.RETURN)
    time.sleep(3)

    end = driver.find_element(By.XPATH, '//*[@id="reportViewer_ctl08_ctl20_txtValue"]')
    end.clear()
    time.sleep(2)

    end.send_keys(str(end_date))
    end.send_keys(Keys.RETURN)

def download_report(): #click view report
    # div = driver.find_element(By.ID, 'reportViewer_ReportViewer')
    # table = div.find_element(By.ID, 'reportViewer_fixedTable')
    # table.find_element(By.ID, 'reportViewer_ctl08_ctl00').click()
    driver.find_element(By.XPATH, '//*[@id="reportViewer_ctl08_ctl00"]').click()
    time.sleep(5)

    #click save
    driver.find_element(By.XPATH, '//*[@id="reportViewer_ctl09"]/div/div[5]').click()
    time.sleep(2)

    #click excel
    driver.find_element(By.XPATH, '//*[@id="reportViewer_ctl09_ctl04_ctl00_Menu"]/div[2]/a').click()  
    

def fill_in_TAR(program_area, program_type, workscope, start_date, end_date): #fill in report parameters:

    if prev[0] != program_area or prev[1] != program_type: #do this IF not the same as previous program.
        # select_program_area(program_area)
        select_element('//*[@id="reportViewer_ctl08_ctl06_ddValue"]', program_area)
        time.sleep(1)

        # select_program_type(program_type)
        select_element('//*[@id="reportViewer_ctl08_ctl08_ddValue"]', program_type)
        time.sleep(0.5)

        # select_provider(1)
        select_element('//*[@id="reportViewer_ctl08_ctl10_ddValue"]', '1')
        time.sleep(0.5)

    # name = select_workscope(workscope)
    name = select_workscope_element('//*[@id="reportViewer_ctl08_ctl12_ddValue"]', workscope)
    time.sleep(2)
    
    fill_dates_TAR(start_date, end_date)
    time.sleep(3)

    download_report()
    time.sleep(3)

    global_prev(program_area, program_type)

    return name

def fill_ROP_CES(workscope):
    time.sleep(1)

    select_element('//*[@id="reportViewer_ctl08_ctl06_ddValue"]', '1')

    time.sleep(1)

    name = select_workscope_element('//*[@id="reportViewer_ctl08_ctl08_ddValue"]', workscope)

    time.sleep(0.5)

    select_element('//*[@id="reportViewer_ctl08_ctl10_ddValue"]', '1')

    time.sleep(3)

    download_report()

    time.sleep(3)

    return name

def fill_ROP_CMS(workscope):
    select_element('//*[@id="reportViewer_ctl08_ctl06_ddValue"]', '1')
    time.sleep(0.5)

    name = select_workscope_element('//*[@id="reportViewer_ctl08_ctl08_ddValue"]', x)
    time.sleep(0.5)

    select_element('//*[@id="reportViewer_ctl08_ctl10_ddValue"]', '1')
    time.sleep(3)

    download_report()
    time.sleep(3)

    return name

def fill_ROP_CYEP(program_type):
    select_element('//*[@id="reportViewer_ctl08_ctl06_ddValue"]', program_type) # select program type
    time.sleep(0.5)

    select_element('//*[@id="reportViewer_ctl08_ctl08_ddValue"]', '1') # select fiscal year
    time.sleep(1)

    select_element('//*[@id="reportViewer_ctl08_ctl10_ddValue"]', '1') # select provider
    time.sleep(1.5)

    name = select_workscope_element('//*[@id="reportViewer_ctl08_ctl12_ddValue"]', '1') # get workscope name
    time.sleep(3)

    download_report()
    time.sleep(3)

    return name

def fill_ROP_B():
    select_element('//*[@id="reportViewer_ctl08_ctl06_ddValue"]', '1')
    time.sleep(0.5)

    name = select_workscope_element('//*[@id="reportViewer_ctl08_ctl08_ddValue"]', '1')
    time.sleep(3)

    download_report()
    time.sleep(3)

    return name

#-----------------------------------Variables and set-up------------------------------------------------------------------

#start time to calculate runtime
start = time.time()

#wait time
t = 1

#input dates
start_date = pd.to_datetime('11/7/22')
end_date = pd.to_datetime('11/13/22')
three_weeks_before = start_date - timedelta(weeks=3)

#TAR workscopes: [program area, program type, workscopes]
TARc_elementary_workscopes = [3,1,[2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22]]
TARc_explore_workscopes = [3,2,[1]]
TARc_high_workscopes = [3,3,[1]]
TARc_middle_workscopes = [3,5,[3, 5, 7, 9, 11, 13, 15]]
TARb_Beacon = [2, 1, [1]]
TARl_aLit = [7, 1, [1,2]]
TARl_AdultLit_AHdis = [7, 1, [1]]
TARl_AdultLit_BEdis = [7, 1, [1]]

TAR_workscopes = [TARb_Beacon, TARc_elementary_workscopes, TARc_explore_workscopes, TARc_high_workscopes, TARc_middle_workscopes, TARl_aLit]

ROP_CES = [2, 4, 7, 9, 11, 13, 15, 17, 19, 21]
ROP_CMS = [3, 5, 7, 9, 11, 13]
ROP_CYEP = [1, 2]

chromedriver_autoinstaller.install()  # Check if the current version of chromedriver exists
                                      # and if it doesn't exist, download it automatically,
                                      # then add chromedriver to path

#options so that webdriver does not close automatically
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
# options.add_argument("download.default_directory=P:/Automation Python/DYCD_Scrape")
# options.add_argument("--headless")

#------------------------------------------------Open Driver and Begin------------------------------


#open up driver
driver = webdriver.Chrome(options=options)

enter_DYCD()

next_page() #all reports are on page 2

data = [] # keep a list to put into final df

find_report('//*[@filename="BEACON_ROP.rdl"]') #find ROP_B
data.append(['ROP Beacon:',''])

for i in range(3): # if internet is slow or some other error, will try three times
    try:
        name = fill_ROP_B()
    except Exception as e:
        print(e)
        print(traceback.format_exc())
        time.sleep(3)
    else:
        break

time.sleep(3)
try: #if value does not exist (DYCD hasn't recorded yet), then just put n/a
    data.append([name + ' Middle School', ga.get_ROP_B('Sheet2', three_weeks_before)])
except Exception as e:
    print(e)
    data.append([name, 'n/a'])

try:
    data.append([name + ' High School', ga.get_ROP_B('Sheet3', three_weeks_before)])
except Exception as e:
    print(e)
    data.append([name, 'n/a'])

print(data)

separate()

find_report('//*[@filename="COMPASS_ROP_OVY.rdl"]') #find ROP_CYEP
data.append(['ROP CYEP:',''])

for x in ROP_CYEP:
    for i in range(3):
        try:
            name = fill_ROP_CYEP(x)
        except Exception as e:
            print(e)
            print(traceback.format_exc())
            time.sleep(2)
        else:
            break

    time.sleep(2)
    try:
        if x == 1:        
            data.append([name, ga.get_ROP_CYEPe(three_weeks_before)])
        else:
            data.append([name, ga.get_ROP_CYEPhs(three_weeks_before)])
    except Exception as e:
        print(e)
        data.append([name, 'n/a'])

    print(data)

separate()

find_report('//*[@filename="COMPASS_ROP_Middle_Unstruct.rdl"]') #find ROP_CMS
data.append(['ROP CMS:',''])

for x in ROP_CMS:
    for i in range(3):
        try:
            name = fill_ROP_CMS(x)
        except Exception as e:
            print(e)
            print(traceback.format_exc())
            time.sleep(3)
        else:
            break

    time.sleep(2)
    try:
        data.append([name, ga.get_ROP_CMS(three_weeks_before)])
    except Exception as e:
        print(e)
        data.append([name, 'n/a'])
    print(data)

separate()

find_report('//*[@filename="CompassElementary_ROP_New.rdl"]')
data.append(['ROP CES:',''])

for x in ROP_CES:
    for i in range(3):
        try:
            name = fill_ROP_CES(x)
        except Exception as e:
            print(e)
            print(traceback.format_exc())
            time.sleep(3)
        else:
            break
        
    time.sleep(2)
    try:
        data.append([name, ga.get_ROP_CES(three_weeks_before)])
    except Exception as e:
        print(e)
        data.append([name, 'n/a'])

    print(data)

separate()

find_report('//*[@filename="TheAttendanceProc_rpt.rdl"]') #Click TAR on reports page
data.append(['Attendance:',''])

global_prev(0, 0) #set prev to default nothing

for workscope in TAR_workscopes:
    for x in workscope[2]:
        for i in range(3):
            try:
                name = fill_in_TAR(workscope[0], workscope[1], x, start_date, end_date)
            except Exception as e:
                print(e)
                print(traceback.format_exc())
                time.sleep(3)
            else:
                break 

        time.sleep(2)

        #process excel based on what program it is since they are all slightly different
        if workscope[0] == 3:
            data.append([name, ga.get_attendance_COMPASS()])
        elif workscope[0] == 2:
            data.append([name, ga.get_attendance_Beacon()])
        else:
            data.append([name, ga.get_attendance_aLit()])

        print(data)

#Create df from list
df = pd.DataFrame(data=data)
print(df)

#save df as excel file
save_file = 'DYCDattendance_list.xlsx'
df.to_excel(save_file, index=False, header=False)

#see how long program took
end = time.time()
print(end - start)

#goodbye and thanks for all the fish
driver.quit()
