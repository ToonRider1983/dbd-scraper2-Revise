
from csv import excel
from time import sleep
from pandas import DataFrame
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from ExcelFileHandler import ExcelFileHandler
from enum import Enum
from bs4 import BeautifulSoup
from dotenv import load_dotenv
import re
import os
import sys
import ctypes

class EventAction(Enum):
    Click = 1
    SendKey = 2
    Clear = 3
    Submit = 4

class scraper:
    def __init__(self, site="", driver_version="", wait_delay=1):
        self.site: str = site
        self.driver_version: str = driver_version
        self.browser: webdriver 
        self.waiting_load: int = wait_delay


    def initialize_browser(self):
        try :
            chrome_options = Options()
            chrome_options.add_argument("--start-maximized")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--log-level=0")
            chrome_options.add_argument("--log-level=1")
            chrome_options.add_argument("--log-level=2")
            chrome_options.add_argument("--log-level=3")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--blink-settings=imagesEnabled=false")
            chrome_options.add_argument("--disable-javascript")
            # chrome_options.add_argument("--ignore-certificate-errors")  # Helps bypass SSL issues
            # chrome_options.add_argument("--allow-running-insecure-content")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_service = Service(ChromeDriverManager(driver_version=self.driver_version).install())
            # self.Show_Progress_OnTitle("DBD Scraping - Ready", 0, 0)
            self.browser = webdriver.Chrome(service=chrome_service, options=chrome_options)
            self.browser.set_page_load_timeout(180)
            self.browser.implicitly_wait(180)
            self.browser.set_script_timeout(180)
            self.browser.get(self.site)
            print(f"ℹ️ Browser initialized.")
            return True
        except Exception as e:
            print(f"❌ Error initializing browser: {e}")
            return False

    
    def OpenSite(self):
        try:
            print(f"🌏 Navigated to site -> {self.site}")
            self.browser.get(self.site)
            WebDriverWait(self.browser, 60).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            return True
        except Exception as e:
            print(f"❌ Failed navigate to site:\n{e}")
            return False
        
    
    def EventDriver_OnElement(self, event_action: EventAction, refer_web_element: By, element_value:any, send_key_value:str="", send_key_val_option:str=""):
        web_element = self.browser.find_element(refer_web_element, element_value)
        
        try:
            match event_action:
                case EventAction.Click      : web_element.click() 
                case EventAction.SendKey    : 
                            if send_key_val_option != "":
                                web_element.send_keys(send_key_value, send_key_val_option)
                            else:
                                web_element.send_keys(send_key_value)

                case EventAction.Clear      : web_element.clear()
                case EventAction.Submit     : web_element.submit()
            
            # sleep(self.waiting_load)
            return True
        except Exception as err:
            print(f"❌ Event driver on element error :{err}")
            return False
        
        
    def SendKey_OnElement(self, refer_web_element: By, element_value:any, send_key_value:str=""):
        try:
            wait = WebDriverWait(self.browser, 60)
            element_input = wait.until(EC.element_to_be_clickable((refer_web_element, element_value)))
            element_input.clear()
            element_input.send_keys(send_key_value + Keys.ENTER)
            return True
        except Exception as err:
            print(f"❌ Event driver on elements error :{err}")
            return False

    def CloseBrowser(self):
        try:
            self.browser.quit()
            print(f"ℹ️ Browser closed.")
            return True
        except Exception as e:
            print(f"❌ Error closing the browser: {e}")
            return False
        
    
    def RemoveUnwanted_Char(self, source_string:str):
        _new_string = ""

        source_string = re.sub(r"^[0-9]", "", source_string)

        for index, char in enumerate(source_string):
            # Remove unwanted charecter of string on start such as '1.', '2.', '53.'
            hex_char = hex(ord(char))
            
            if (hex_char < '0x41' or hex_char > '0x5a') and (hex_char < '0x61' or hex_char > '0x7a'):
                _new_string += char

        _new_string = re.sub(r"^.", "", _new_string).strip()
        _new_string = re.sub(r"[.,.$]|\($\)$\-$", "", _new_string).strip()
        return _new_string
    

    def Show_Progress_OnTitle(self, caption_on_title_bar:str="", current_progress:int = 0, total_progress:int = 0):
        on_progress:float = 0.0

        try:
            on_progress:float = (current_progress / total_progress) * 100
        except:
            pass

        if current_progress == 0 and total_progress == 0:
            ctypes.windll.kernel32.SetConsoleTitleW(f"{caption_on_title_bar} - Ready")
        elif on_progress > 0:
            ctypes.windll.kernel32.SetConsoleTitleW(f"{caption_on_title_bar}.. {on_progress:.2f}%")
        elif on_progress == 100:
            ctypes.windll.kernel32.SetConsoleTitleW(f"{caption_on_title_bar} - Completed")
        

    def DataScrap(self, data_frame:DataFrame, excel:ExcelFileHandler, progress_start:int=0):
        _excel = ExcelFileHandler()
        is_completed:bool = False
        print(f"💻 Start Scraping Data")
        last_index = -1
        for _data in data_frame.itertuples():
                last_index = _data.Index
                try :
                    if progress_start == 0 or _data.Index >= progress_start:
                        print(f"DBD Scraper II {(_data.Index + 1) * 100 / len(data_frame):.2f}% ({_data.Index + 1}/{len(data_frame)}) -> Scraping {_data.Account_Name} ...")
                        self.browser.get(f'https://datawarehouse.dbd.go.th/juristic/searchInfo?keyword={str(_data.Thai_Tax_ID).strip()}')
                        WebDriverWait(self.browser, 60).until(
                            # EC.presence_of_element_located((By.TAG_NAME, "body"))
                            EC.presence_of_element_located((By.CLASS_NAME, "page-header"))
                        )

                        soup = BeautifulSoup(self.browser.page_source, 'html.parser')
                        div_root = soup.find('div', class_='page')

                        if div_root is None:
                            print("❌ no result page for", _data.Account_Name)
                            continue

                        # Scrape Thai Tax ID
                        Thai_tax_ID = ""
                        if _data.Thai_Tax_ID == "" or _data.Thai_Tax_ID == "-":
                            try:
                                h4 = div_root.find('h4', string=re.compile("เลขทะเบียนนิติบุคคล")).get_text().strip(string=True)
                                tax_id = h4.split(" : ")[1]
                                Thai_tax_ID = str(tax_id).strip()
                            except:
                                print("error scraping tax id")
                        elif len(str(_data.Thai_Tax_ID)) > 13:
                            Thai_tax_ID = str(_data.Thai_Tax_ID).strip().replace(".0", "")
                            Thai_tax_ID = str(f"0{Thai_tax_ID}")
                        else:
                            Thai_tax_ID = str(_data.Thai_Tax_ID).strip()

                        # try:
                        #     h4 = div_root.find('h4', string=re.compile("เลขทะเบียนนิติบุคคล")).get_text().strip(string=True)
                        #     tax_id = h4.split(" : ")[1]
                        #     Thai_tax_ID = str(tax_id).strip()
                        # except:
                        #     pass
                        #     # print("error scraping tax id")

                        # if len(str(_data.Thai_Tax_ID)) > 13:
                        #     Thai_tax_ID = str(_data.Thai_Tax_ID).strip().replace(".0", "")
                        #     Thai_tax_ID = str(f"0{Thai_tax_ID}")

                        data_frame.at[_data.Index, 'Thai_Tax_ID'] = Thai_tax_ID
                        print(f"Found tax id: ({Thai_tax_ID} | {_data.Thai_Tax_ID}) for {_data.Account_Name} 👍")

                        # Scrape TSIC ID - OK
                        try:
                            target = soup.find('div', class_="card-infos")
                            target = target.find_all_next('div', class_='col-8')
                            tsic_id = re.findall(r'\d+', target[2].get_text())
                            data_frame.at[_data.Index, 'เลขประเภทธุรกิจ'] = tsic_id[0]
                            print(f"Found tsic id: {tsic_id[0]} for {_data.Account_Name} 👍")
                        except:
                            pass

                        # Scrape Company Type 
                        try:
                            company_type = div_root.find_all_next('div', class_='col-8').get_text().strip(string=True)
                            # company_type = div_root.find('div', string=re.compile("ประเภทนิติบุคคล")).find_next("div").get_text(strip=True)
                            data_frame.at[_data.Index, 'ประเภทบริษัท'] = company_type
                            print(f"Found company type: {company_type} for {_data.Account_Name} 👍")
                        except:
                            pass

                        # Scrape Billing City, Billing Street, Billing Province
                        try:
                            address_div = soup.find("div", string=re.compile("ที่ตั้งสำนักงานแห่งใหญ่")).find_next("div")
                            address_text = address_div.get_text(strip=True)

                            data_frame.at[_data.Index, 'Billing_Street'] = address_text
                            match_province = re.search(r"จ\.[^ ]+|กทม\.[^ ]+|กรุงเทพ[^ ]+|กรุงเทพมหานคร[^ ]+", address_text)
                            data_frame.at[_data.Index, 'Billing_Province'] = match_province.group() if match_province else ""
                            print(f"🌏 Extracted Address for {_data.Account_Name}: Street: {address_text}, Province: {data_frame.at[_data.Index, 'Billing_Province']}")
                        except:
                            pass

                    sleep(1)  # To avoid overwhelming the server with requests

                    if _data.Index % 300 == 0:
                        print(f"💾 Auto-saving progress to Excel at index {_data.Index}...")
                        _excel.write_excel(data_frame, f"DBD-AutoSave-{_data.Index}", "Sheet1", False)
                        print(f"✅ Auto-save completed.")

                    print(f"✅ Scraping completed successfully.")
                    is_completed = True
                except KeyboardInterrupt:
                    print(f"🛑 Stop Application by User")
                    is_completed = True
                    break
                except Exception as e:
                    print(f"❌ Error during data scraping: {e}")
                    is_completed = False
                    continue
            
        return is_completed, data_frame, last_index
                

if __name__ == "__main__":
    # ctypes.windll.kernel32.SetConsoleTitleW("DBD Scraping - Ready")
    load_dotenv()
    base_url = os.getenv("BASE_URL")
    chrome_version = os.getenv("CHROME_VERSION")
    wait_time = int(os.getenv("WAIT_TIME"))
    ascii_art = r"""
██████╗ ██████╗ ██████╗       ███████╗ ██████╗██████╗  █████╗ ██████╗ ███████╗██████╗     ██╗██╗
██╔══██╗██╔══██╗██╔══██╗      ██╔════╝██╔════╝██╔══██╗██╔══██╗██╔══██╗██╔════╝██╔══██╗    ██║██║
██║  ██║██████╔╝██████╔╝█████╗███████╗██║     ██████╔╝███████║██████╔╝█████╗  ██████╔╝    ██║██║
██║  ██║██╔══██╗██╔═══╝ ╚════╝╚════██║██║     ██╔══██╗██╔══██║██╔═══╝ ██╔══╝  ██╔══██╗    ██║██║
██████╔╝██████╔╝██║           ███████║╚██████╗██║  ██║██║  ██║██║     ███████╗██║  ██║    ██║██║
╚═════╝ ╚═════╝ ╚═╝           ╚══════╝ ╚═════╝╚═╝  ╚═╝╚═╝  ╚═╝╚═╝     ╚══════╝╚═╝  ╚═╝    ╚═╝╚═╝
    """
    print(ascii_art)
    scraper = scraper(site=base_url, driver_version=chrome_version, wait_delay=wait_time)
    try :
        filename = input("⌨️ กรุณป้อนที่อยู่ของไฟล์ : ")
        sheet_name = input("⌨️ กรุณาป้อนชื่อชีท (Sheet) ของไฟล์ : ")
        excel = ExcelFileHandler(file_path=filename, sheet_name=sheet_name)
        _data_frame = excel.read_excel()
        if _data_frame is not None:
            if scraper.initialize_browser():
                if scraper.OpenSite():
                    sleep(scraper.waiting_load)
                    if scraper.EventDriver_OnElement(EventAction.Click, By.ID, "btnWarning", "", ""):
                        # if scraper.DataScrap(_data_frame, 0)[0]:
                        #     scraper.CloseBrowser()
                        progress_index = 0
                        is_done = False

                        is_done, _data_frame, progress_index = scraper.DataScrap(_data_frame, excel, progress_index)
                        if is_done:
                            print(f"💾 Saving scraped data to Excel...")
                            excel.write_excel(_data_frame, f"DBD-Scraped2-", "Sheet1", False)
                            print(f"✅ Data scraping completed and saved to Excel.")
                            scraper.CloseBrowser()
                    
    except KeyboardInterrupt:
        print(f"🛑 Stop Application by User")
    except Exception:
        scraper.CloseBrowser()