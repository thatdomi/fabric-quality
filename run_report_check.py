import argparse
import yaml
import os
import pandas as pd
import subprocess
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from termcolor import colored

from utils.PowerBIRestHandler import PowerBIRestHandler

EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" #path to the edge installation
EDGE_PROFILE_PATH = r"C:\Users\domin\AppData\Local\Microsoft\Edge\User Data\\" # path where edge profiles are stored

def load_config(config_file):
    with open(config_file, 'r') as file:
        config = yaml.safe_load(file)
    return config

class PowerBIReportProbe:
    def __init__(self, profile_name=None):
        self.profile_name=profile_name
        self.results = [["url_report", "url_page", "page_nr", "has_error"]]
        self.has_found_any_errors = False

    def init_selenium_driver_edge(self):
        options = webdriver.EdgeOptions()
        
        # add profile options if profile is specified
        if self.profile_name:
            profilePath = EDGE_PROFILE_PATH
            
            # Here you set the path of the profile ending with User Data not the profile folder
            options.add_argument(f"--user-data-dir={profilePath}");
            # Here you specify the actual profile folder
            options.add_argument(f"--profile-directory={self.profile_name}");

        self.driver = webdriver.Edge(options=options)

    def get_report_page_url(self, report_base_url, page_number = None):
        if not page_number or page_number == 1 or page_number ==0:
            # add Report section to make sure we land on the first page of a report
            url = report_base_url + "/ReportSection"
        else:
            url = report_base_url + f"/ReportSection{page_number}"
        return url

    def get_report_page_id(self, url) -> str:
        # returns the id of the power bi report page based on the url

        # check if url contains section information
        if "ReportSection" in url:
            # getting the id of the power bi report page
            sectionId = url.split("ReportSection")[1].split("?")[0]
        
        if not sectionId:
            print("currently on page 1, section id is empty")
        
        return sectionId

    def load_report_page_by_url(self, url):
        # function to handle calls to power bi url, including waits
        print(f"loading url: {url}")
        self.driver.get(url)
        #driver.find_elements(By.TAG_NAME, "pbi-overlay-container")

        element = WebDriverWait(self.driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, "pbi-overlay-container"))
        )
        if element:
            
            # Checking for page navigation which is not expanded
            try:
                pageNavBtn = WebDriverWait(self.driver, 1).until(
                    EC.element_to_be_clickable((By.ID, "pageNavBtn"))
                )
                if pageNavBtn:
                    pageNavBtn.click()
                    print("Pages expanded")

            except:
                print("Page Expand Button not present.")

            print("loaded Page")


    def has_report_page_error_visuals(self) -> bool:
        # function to check for visuals which have errors in them
        print(f"Checking for visuals with errors in {self.driver.current_url}")
        try: 
            element = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.TAG_NAME, "canvas-visual-error-overlay"))
            )

            self.has_found_any_errors = True
            return True
        except:
            print("no errors in visuals found")
            return False

    def get_report_all_pages(self, report_base_url):
        # function to loop through all pages within a report
        self.load_report_page_by_url(report_base_url)
        try:
            # Assuming the buttons are within a mat-action-list
            mat_action_list = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//mat-action-list[@data-testid='pages-navigation-list']"))
            )
            # Locate the buttons within the mat-action-list (adjust the locator as per your HTML structure)
            buttons = mat_action_list.find_elements(By.TAG_NAME, "button")
        
        except:
            # if there is only one page, there is no page navigation
            print("no page buttons found")
            buttons = [0]

        current_page_number = 1
        report_pages_count = len(buttons)
        for button in buttons:
            if type(button) != int: # if the button is a button object and not the arbitrary 0, click the button
                button.click()
            report_page_url = self.driver.current_url
            page_id = self.get_report_page_id(self.driver.current_url)
            has_report_page_errors = self.has_report_page_error_visuals()
            
            self.log_results(
                report_base_url=report_base_url,
                url=report_page_url,
                report_page_number=f"{current_page_number}/{report_pages_count}",
                has_report_page_errors=has_report_page_errors
                )

            print(f"Page Url: {report_page_url}")
            print(f"Page {current_page_number}/{report_pages_count} in Report")
            print(f"Has page errors: {has_report_page_errors}")
                
            # Make sure global switch is set that errors have been found
            if has_report_page_errors:
                self.has_found_any_errors = True
            
            current_page_number = current_page_number+1

    def log_results(self, report_base_url, url, report_page_number, has_report_page_errors):
        # make sure results are stored somewhere
        self.results.append([report_base_url, url, report_page_number, has_report_page_errors])
    
    def show_results(self):
        df = pd.DataFrame.from_records(self.results)
        
        print(df)

        # Display the DataFrame with clickable links
        df[0][1:] = df[0][1:].apply(lambda x: f'<a href="{x}" target="_blank">link to report</a>')
        df[1][1:] = df[1][1:].apply(lambda x: f'<a href="{x}" target="_blank">link to page</a>')
        
        df.columns = df.iloc[0]
        df = df[1:]
        # Save the DataFrame to an HTML file
        dirpath = os.getcwd()
        html_file_path = dirpath +"\\output.html"
        df.to_html(html_file_path, escape=False, render_links=True)
        

if __name__ == "__main__":
    # python run_report_check.py -t "DEV" -n "Blog - Region"
    ############################## Argument Parsing ################################################
    parser = argparse.ArgumentParser(description='Check Reports in Power BI Workspaces with Selenium')

    # Add arguments
    parser.add_argument('-c', '--config', default='tenants.yaml', help='Path to the YAML config file')
    parser.add_argument('-t', '--tenant', default='DEFAULT', required=False, help='Specify path in tenants.yaml and select option accordingly')
    parser.add_argument('-n', '--name', default='Blog - Region', help='Workspace name')

    # Parse the command-line arguments
    args = parser.parse_args()

    # Access the values using args
    tenantName = args.tenant
    workspaceName = args.name

    # Load config
    config = load_config(args.config)

    # Access config values
    profile_name = config.get(tenantName)

    print(config)
    print(profile_name)

    # Replace "ProfilePath" with the actual path to the profile directory you want to use
    profile_directory = profile_name
    edge_path = EDGE_PATH
    command = [edge_path,  "--profile-directory=" + profile_directory,"--start-maximized"]

    try:
        process = subprocess.Popen(command, stdout=subprocess.PIPE) # stdout=subprocess.PIPE, 

    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")
    #################################################################################################

    client = PowerBIRestHandler()
    workspaceId = client.get_workspace_by_name(workspaceName) # Demo: "Blog - Region"
    print(workspaceId)
    report_urls = client.get_reports_in_workspace(workspaceId=workspaceId)
    print(report_urls)

    os.system('taskkill /f /im  "msedge.exe"')
    input("Press enter to continue....")
    
    # setup
    probe = PowerBIReportProbe(profile_name=profile_name)
    probe.init_selenium_driver_edge()

    print(report_urls)
    # main logic
    for report_base_url in report_urls:
        probe.get_report_all_pages(report_base_url=report_base_url)

    # results
    if probe.has_found_any_errors:
        print(colored("----------there are errors in the report---------------", "red"))
    else:
        print(colored("------------there are no errors in the report-------------", "green"))

    probe.show_results()

    # close browser
    probe.driver.quit()

    # open 
    