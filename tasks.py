from RPA.Browser.Selenium import Selenium
from RPA.Dialogs import Dialogs
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF

import csv

browser = Selenium()
dialog = Dialogs()
http = HTTP()
tables = Tables()
pdf = PDF()

url = 'https://robotsparebinindustries.com/#/robot-order'

def open_website_and_close_modal(url: str):
        #browser = Selenium()
        browser.open_available_browser(url)
        browser.click_button_when_visible('css:.btn.btn-dark')

def get_orders():
    #dialog = Dialogs()
    #http = HTTP()
    #tables = Tables()

    dialog.add_heading('Upload CSV File')
    dialog.add_text_input(label='Give URL for CSV-file', name='url')
    response = dialog.run_dialog()

    http.download(response.url, overwrite=True)

    with open('orders.csv') as csvFile:
        # list of dictionaries, https://stackoverflow.com/questions/21572175/convert-csv-file-to-list-of-dictionaries
        table = [{k: v for k, v in row.items()} for row in csv.DictReader(csvFile, skipinitialspace=True)]

        #csvReader = csv.DictReader(csvFile)
        #table = list(csvReader)
        #print(table)
        return table

def fill_the_form(row: str):
    #print(row)
    #print(row['Body'])
    browser.select_from_list_by_value('id:head', row['Head'])
    browser.select_radio_button('body', 'id-body-'+row['Body'])
    browser.input_text('css:.form-control', row['Legs'])
    browser.input_text('name:address', row['Address'])

def preview_the_robot():
    browser.click_button('id:preview')

def submit_the_order():
    try:
        browser.submit_form('tag:form')
    except Exception as inst:   # https://docs.python.org/3/tutorial/errors.html
        print(type(inst))    # the exception instance
        print(inst.args)

'''
--- YRITYSKERRAT, AIKA? ---
Submit the order
    Wait Until Keyword Succeeds    5x    1 sec    Submit Form    tag:form
'''

def store_the_receipt_as_a_PDF_file(row):
    browser.wait_until_element_is_visible('tag:form', 5)  # preview takes some time to load
    # variables for head and body
    head = browser.get_selected_list_label('id:head')
    body = browser.get_text('//*[@id="root"]/div/div[1]/div/div[1]/form/div[2]/div/div['+row["Body"]+']/label')
    #print(body)
    order_html = f'<html><h3>Order number: {row["Order number"]}</h3><p>Head: {head}<br>Body: {body}<br>Legs: {row["Legs"]}<br>Address: {row["Address"]}</p></html>'
    #print(order_html)
    pdf.html_to_pdf(order_html, '.\\output\\pdf\\'+row["Order number"]+'.pdf')
    #print('.\\output\\pdf\\'+row["Order number"]+'.pdf')  
    return '.\\output\\pdf\\'+row["Order number"]+'.pdf'

def take_a_screenshot_of_the_robot(row):
    browser.wait_until_element_is_visible('id:robot-preview-image', 2)
    browser.set_screenshot_directory('.\\output\\screenshots')
    scrshot = browser.capture_element_screenshot('id:robot-preview-image',  row['Order number']+'.png')
    return scrshot

def embed_the_robot_screenshot_to_the_receipt_PDF_file(screenshot, pdf_file):
    pdf.open_pdf(pdf_file)
    #files = tables.create_list(screenshot)
    files = [screenshot]
    pdf.add_files_to_pdf(files, pdf_file, True)
    pdf.close_pdf(pdf_file)

def go_to_order_another_robot():
    browser.select_from_list_by_value('id:head', '1')
    browser.select_radio_button('body', 'id-body-1')
    browser.input_text('css:.form-control', '2')
    legs =  browser.get_element_attribute('css:.form-control', 'value')
    print(legs)
    browser.input_text('name:address', 'My Address 1')
    address = browser.get_element_attribute('name:address', 'value')
    print(address)

'''
Go to order another robot
    # what was expected here? to order just some extra robot?
    # filling form
    Select From List By Value   id:head    1
    Select Radio Button    body    id-body-1
    Input Text    css:.form-control    2
    ${legs}=   Get Element Attribute    css:.form-control      value
    Input Text    name:address    My Address 1
    ${address}=    Get Element Attribute    name:address    value

    # clicking preview button
    Click Button    id:preview

    # submitting form
    Wait Until Keyword Succeeds    5x    1 sec    Submit Form    tag:form

    # making pdf
    Wait Until Element Is Visible    tag:form
    # variables for head and body
    ${head}=    Get Selected List Label    id:head
    ${body}=    Get Text    //*[@id="root"]/div/div[1]/div/div[1]/form/div[2]/div/div[1]/label

    ${order_html}=    Catenate      
    ...    <html><h3>Order number: Another robot </h3><p>Head: ${head}<br>Body: ${body}<br>Legs: ${legs}<br>Address: ${address}</p></html>

    ${path}=    Set Variable    ${CURDIR}\\output\\pdf\\another_robot.pdf
    Html To Pdf    ${order_html}    ${path}

    # taking screenshot
    Wait Until Element Is Visible    id:robot-preview-image    2
    Set Screenshot Directory    ${CURDIR}\\output\\screenshots
    ${scrshot}=    Capture Element Screenshot    id:robot-preview-image    another_robot.png

    # embed screenshot in pdf
    Open Pdf    ${path}       
    ${files}=    Create List   ${scrshot}
    Log    ${files}
    Add Files To Pdf    ${files}    ${path}    append=True
    Close Pdf    ${path} 


'''

def main():
    try:
        open_website_and_close_modal(url)
        orders = get_orders()
        #print(orders)
        for row in orders:
            fill_the_form(row)
            preview_the_robot()
            submit_the_order()
            pdf = store_the_receipt_as_a_PDF_file(row)
            screenshot = take_a_screenshot_of_the_robot(row)
            embed_the_robot_screenshot_to_the_receipt_PDF_file(screenshot, pdf)
        go_to_order_another_robot()

    finally:
        browser.close_all_browsers()

if __name__ == "__main__":
    main()