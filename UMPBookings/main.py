from selenium import webdriver
from selenium.webdriver.chrome import service, options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
import time
import datetime

options = options.Options()
service = service.Service('./chromedriver')
options.add_argument("--disable-notifications")
options.add_experimental_option('excludeSwitches', ['enable-logging'])
chrome = webdriver.Chrome(service=service, options=options)


def wait_ele(func):
    try:
        if (not WebDriverWait(chrome, 20.0, 1.0).until(func)):
            chrome.refresh()
            wait_ele(func)
        return True
    except:
        return False


def operate():

    chrome.get("https://visa-exam.ump.com.hk/")
    webWait = WebDriverWait(chrome, 20.0, 1.0)

    wait_ele(EC.text_to_be_present_in_element(
            (By.XPATH, "//*[contains(text(),'Visa Medical Examination Appointment System 簽証驗身預約系統')]"), "Visa Medical Examination Appointment System 簽証驗身預約系統"))

    applicant_field = Select(chrome.find_element(By.ID, "NoOfApplicants"))
    visa_selection = Select(chrome.find_element(By.ID, "SelectVisa"))
    time_slot = chrome.find_element(
        By.XPATH, f"//*[@id=\"tr_date\"]/td[2]/table/tbody/tr/td[3]/p/u")

    applicant_field.select_by_value("1")
    visa_selection.select_by_value("2/CANADA")
    time_slot.click()

    tab_list = chrome.window_handles
    print(f"Current Window {chrome.current_window_handle.title()}")
    for tab in tab_list:
        if tab is not chrome.current_window_handle:
            chrome.switch_to.window(tab)
            print(f"Switched Window {chrome.current_window_handle.title()}")

    def wait():
        try:

            if (webWait.until(EC.text_to_be_present_in_element(
                    (By.XPATH, "/html/body/p/b/font"), "Booking Availability for Canada Visa Examination"))):
                return SelectTimeSlot()

            if (wait_ele(EC.text_to_be_present_in_element(
                    (By.XPATH, "//*[contains(text(),'504 Gateway Time-out')]"), "504 Gateway Time-out"))):

                chrome.close()

                chrome.switch_to.window(tab_list[0])

                time_slot.click()

                tabs = chrome.window_handles
                print(f"Current Window {chrome.current_window_handle.title()}")

                for tab in tabs:
                    if tab is not chrome.current_window_handle:
                        chrome.switch_to.window(tab)
                        print(
                            f"Switched Window {chrome.current_window_handle.title()}")
                return wait()
        except Exception as e:
            print(e)
            chrome.close()

            chrome.switch_to.window(tab_list[0])

            time_slot.click()

            tabs = chrome.window_handles
            print(f"Current Window {chrome.current_window_handle.title()}")

            for tab in tabs:
                if tab is not chrome.current_window_handle:
                    chrome.switch_to.window(tab)
                    print(
                        f"Switched Window {chrome.current_window_handle}")
            return wait()

    def SelectTimeSlot():
        print("Exec SelectTimeSlot")
        table = chrome.find_element(
            By.CSS_SELECTOR, "table[bgcolor='#006699']")
        trs = table.find_elements(By.TAG_NAME, "tr")

        clicked = False
        for table_row in trs:
            centrals = table_row.find_elements(
                By.XPATH, "//td[contains(text(), 'Central')]")
            tsts = table_row.find_elements(
                By.XPATH, "//td[contains(text(), 'Tsim Sha Tsui')]")
            jordans = table_row.find_elements(
                By.XPATH, "//td[contains(text(), 'Jordan')]")
            print(len(centrals), len(tsts), len(jordans), clicked)
            if (len(centrals) > 0 or len(tsts) > 0 or len(jordans) > 0) and (not clicked):
                freeTimeSlot = table_row.find_elements(
                    By.XPATH, ".//td[not(contains(text(), '-'))]/*[not(contains(text(), '-'))]")
                print(len(freeTimeSlot))
                if len(freeTimeSlot) > 0:
                    for slot in freeTimeSlot:
                        try:
                            slot.click()
                            # webWait.until(EC.title_is("Visa Medical Examination Appointment System"))
                            if chrome.current_url == "https://visa-exam.ump.com.hk/":
                                clicked = True
                            else:
                                pass
                            break
                        except Exception as e:
                            print(e)
                            pass

            if clicked:
                break
        return True

    wait()

    if chrome.current_url != "https://visa-exam.ump.com.hk/":
        if len(chrome.window_handles) > 1:
            chrome.switch_to.window(chrome.window_handles[1])
            chrome.close()

        chrome.switch_to.window(chrome.window_handles[0])

    chrome.find_element(By.ID, "txtFamilyName_01").send_keys("WONG")
    chrome.find_element(By.ID, "txtGivenName_01").send_keys("POK CHUNG")
    chrome.find_element(By.ID, "txtChiName_01").send_keys("黃博頌")
    Select(chrome.find_element(By.ID, "ddlSex_01")).select_by_value("Female 女")
    chrome.find_element(By.ID, "txtPassport_01").send_keys("K06680883")
    chrome.find_element(By.ID, "txtHKID_01").send_keys("Y608527(8)")
    selects = chrome.find_element(
        By.ID, "epoch_basic_dob_01_calendar").find_elements(By.XPATH, ".//select")
    Select(selects[0]).select_by_value("10")
    Select(selects[1]).select_by_value("1997")
    chrome.find_element(By.ID, "epoch_basic_dob_01_cell_td").find_element(
        By.ID, "epoch_basic_dob_01_calcells").find_element(By.XPATH, ".//td[text()='6']").click()
    chrome.find_element(By.ID, "txtPatientEmail_01").send_keys("HARUSHO@PM.ME")
    chrome.find_element(By.ID, "txtContactPhone_01").send_keys("67384382")
    chrome.find_element(By.ID, "txtHKAddress_01").send_keys(
        "Flat F, 7/F, BLOCK 4, TSUI LAI GARDEN, SHEUNG SHUI, NT, HONG KONG")
    Select(chrome.find_element(By.ID, "ddlVisaCategory_01")).select_by_value("1")
    Select(chrome.find_element(By.ID, "Question_01_2")).select_by_value("0")
    Select(chrome.find_element(By.ID, "Question_01_6")).select_by_value("0")
    Select(chrome.find_element(By.ID, "Question_01_9")).select_by_value("0")
    Select(chrome.find_element(By.ID, "Question_01_36")).select_by_value("0")
    Select(chrome.find_element(By.ID, "Question_01_13")).select_by_value("0")
    Select(chrome.find_element(By.ID, "Question_01_22")).select_by_value("0")
    Select(chrome.find_element(By.ID, "Question_01_23")).select_by_value("0")
    Select(chrome.find_element(By.ID, "Question_01_18")).select_by_value("0")
    Select(chrome.find_element(By.ID, "Question_01_18")).select_by_value("0")

    chrome.find_element(By.ID, "basic_container_date_01_29_cell_td").find_element(
        By.ID, "basic_container_date_01_29_calcells").find_element(By.XPATH, ".//td[text()='8']").click()

    Select(chrome.find_element(By.ID, "Question_01_32")).select_by_value("1")
    Select(chrome.find_element(By.ID, "Consent_UmpExam_01")).select_by_value("1")
    Select(chrome.find_element(By.ID, "Consent_HivTest_01")).select_by_value("1")
    chrome.find_element(By.ID, "chkPrivacyConsent").click()
    chrome.find_element(By.NAME, "action").click()

    try:
        WebDriverWait(chrome, 10).until(
            EC.alert_is_present(), 'Please select Preferred Date.')

        alert = chrome.switch_to.alert
        alert.accept()
        time_slot.click()
        wait()
    except Exception as e:
        print(e)

    chrome.quit()


if __name__ == '__main__':

    sleep_time = 10
    while True:
        now = datetime.datetime.now()
        if now >= datetime.datetime(2022, 12, 22, 00, 00, 00):
            operate()
            if sleep_time > 1:
                sleep_time = 1
        print("operating", now)
        time.sleep(sleep_time)
