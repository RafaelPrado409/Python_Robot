from multiprocessing.connection import wait
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem

browser_lib = Selenium()
excel_lib = Files()
file_system = FileSystem()

def create_excel_file(path, fmt, name):
  excel_lib.create_workbook(path, fmt)
  excel_lib.save_workbook(name)

def format_excel():
  excel_lib.set_cell_value(1, 1, "Agencies")
  excel_lib.set_cell_value(1, 2, "Amount")
  excel_lib.save_workbook("Agencies.xlsx")

def open_the_website(url):
    browser_lib.open_available_browser(url)

def get_org_title():
  org_list = []
  count = 1
  for i in range(9):
    for j in range(3):
      if count == 27:
        break
      title = browser_lib.get_text("xpath://*[@id='agency-tiles-widget']/div/div[{}]/div[{}]/div/div/div/div[1]/a/span[1]".format(i+1,j+1))
      org_list.append(title)
      excel_lib.set_cell_value(count+1, 1, org_list[count-1])
      excel_lib.save_workbook("Agencies.xlsx")
      count = count + 1

def get_org_amount():
  amount_list = []
  count = 1
  for i in range(9):
    for j in range(3):
      if count == 27:
        break
      title = browser_lib.get_text("xpath://*[@id='agency-tiles-widget']/div/div[{}]/div[{}]/div/div/div/div[1]/a/span[2]".format(i+1,j+1))
      amount_list.append(title)
      excel_lib.set_cell_value(count+1, 2, amount_list[count-1])
      excel_lib.save_workbook("Agencies.xlsx")
      count = count + 1
  
def get_all_table():
  excel_lib.create_worksheet("Individual Investments")
  for i in range(7):
    table_cell_result = browser_lib.get_table_cell("xpath://*[@id='investments-table-object_wrapper']/div[3]/div[1]/div/table", 2, i+1)
    excel_lib.set_worksheet_value(1, i + 1, table_cell_result)
    excel_lib.save_workbook("Agencies.xlsx")
  rows_number = browser_lib.get_text("xpath://*[@id='investments-table-object_info']")
  rows_number = rows_number.split(" ")
  rows_number = int(rows_number[5])
  for i in range(rows_number):
    for j in range(7):
      browser_lib.wait_until_page_contains_element("xpath://*[@id='investments-table-object_wrapper']", 20)
      browser_lib.select_from_list_by_label("xpath://*[@id='investments-table-object_length']/label/select", "All")
      browser_lib.wait_until_element_contains("xpath://*[@id='investments-table-object_info']", "Showing 1 to {} of {} entries".format(rows_number, rows_number), 30)
      table_cell_result = browser_lib.get_table_cell("xpath:/html/body/main/div/div/div/div[4]/div/div/div/div[2]/div/div[1]/div/div/div/div/div[3]/div[2]/table", i+2, j+1)
      excel_lib.set_worksheet_value(i+2, j+1, table_cell_result)
      excel_lib.save_workbook("Agencies.xlsx")
      # ------> FILL THE EXCEL WITH THE ENTIRE TABLE <----------------
  for i in range(rows_number):
    if i >= 1:
      browser_lib.wait_until_page_contains_element("xpath://*[@id='investments-table-object_wrapper']", 20)
      browser_lib.select_from_list_by_label("xpath://*[@id='investments-table-object_length']/label/select", "All")
    browser_lib.wait_until_element_contains("xpath://*[@id='investments-table-object_info']", "Showing 1 to {} of {} entries".format(rows_number, rows_number), 30)
    browser_lib.wait_until_page_does_not_contain("xpath://*[@id='investments-table-object_paginate']/span/a[2]")
    is_visible = browser_lib.is_element_visible("xpath://*[@id='investments-table-object']/tbody/tr[{}]/td[1]/a".format(i+1))
    if is_visible == True:
      browser_lib.click_link("xpath://*[@id='investments-table-object']/tbody/tr[{}]/td[1]/a".format(i+1))
      browser_lib.wait_until_page_contains_element("xpath://*[@id='investment-quick-stats-widget']/div/div[5]/div[1]", 10)
      browser_lib.click_link("xpath://*[@id='business-case-pdf']/a")
      browser_lib.wait_until_element_contains("xpath://*[@id='business-case-pdf']/span", "Generating PDF...", 20)
      browser_lib.wait_until_page_does_not_contain_element("xpath://*[@id='business-case-pdf']/span", 20)
      browser_lib.go_to("https://itdashboard.gov/drupal/summary/005")
                                     
# Define a main() function that calls the other functions in order:
def main():
    try:
        create_excel_file("/Users/rafaelprado", "xlsx", "Agencies")
        format_excel()
        open_the_website("https://itdashboard.gov/")
        browser_lib.maximize_browser_window()
        browser_lib.click_element_when_visible("id:node-23", False)
        browser_lib.wait_until_element_contains("id:agency-tiles-container", "Department of Agriculture", 5000)
        get_org_amount()
        get_org_title()
        browser_lib.click_element_when_visible("xpath://*[@id='agency-tiles-widget']/div/div[1]/div[1]/div/div/div/div[2]/a")
        browser_lib.set_focus_to_element("xpath://*[@id='block-itdb-custom--5']/div/div/div/div[1]/div/h3")
        browser_lib.wait_until_page_contains_element("xpath://*[@id='investments-table-object_wrapper']", 1000)
        browser_lib.select_from_list_by_label("xpath://*[@id='investments-table-object_length']/label/select", "All")
        browser_lib.wait_until_page_does_not_contain("xpath://*[@id='investments-table-object_paginate']/span/a[2]")
        get_all_table()
       
    finally:
        browser_lib.close_all_browsers()


# Call the main() function, checking that we are running as a stand-alone script:
if __name__ == "__main__":
    main()
 