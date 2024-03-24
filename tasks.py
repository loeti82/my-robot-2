from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

import shutil

@task
@task
def order_robots_from_RobotSpareBin():
    """
    comment to trigger task
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    Complete all orders even there are failure
    """
    #browser.configure(slowmo=200)
    open_robot_order_website()
    orders = get_orders()
    for order in orders:
        fill_and_submit_form(order)
    archive_receipts()
    clean_up()

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    page = browser.page()
    page.click("button:text('OK')")

def get_orders():
    """
    Downloads excel file from the given URL
    Order number,Head,Body,Legs,Address
    """
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    table = Tables()
    orders = Tables.read_table_from_csv(table, path="./orders.csv")
    return orders

def order_another_bot():
    """Clicks on order another button to order another bot"""
    page = browser.page()
    page.click("#order-another")

def click_ok():
    """Clicks on ok whenever a new order is made for bots"""
    page = browser.page()
    page.click('text=OK')

def fill_and_submit_form(order):
    page = browser.page()
    page.select_option("#head", str(order['Head']))
    page.click("#id-body-"+str(order['Body']))
    page.fill("xpath=//html/body/div/div/div[1]/div/div[1]/form/div[3]/input", str(order['Legs']))
    page.fill("#address", str(order['Address']))
    while True:
        page.click("#order")
        order_another = page.query_selector("#order-another")
        if order_another:
            pdf_path = store_receipt_as_pdf(int(order['Order number']))
            screenshot_path = screenshot_robot(int(order['Order number']))
            embed_screenshot_in_receipt(screenshot_path, pdf_path)
            order_another_bot()
            click_ok()
            break

def store_receipt_as_pdf(order_nubmer):
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf_path = "output/receipts/{0}.pdf".format(order_nubmer)
    pdf.html_to_pdf(receipt_html, pdf_path)
    return pdf_path

def screenshot_robot(order_numer):
    page = browser.page()
    screenshot_path = "output/screenshots/{0}.png".format(order_numer)
    page.locator("#robot-preview-image").screenshot(path=screenshot_path)
    return screenshot_path

def embed_screenshot_in_receipt(screenshot_path, pdf_path):
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(image_path=screenshot_path,
                                   source_path=pdf_path,
                                   output_path=pdf_path)
    
def archive_receipts():
    lib = Archive()
    lib.archive_folder_with_zip("./output/receipts", "./output/receipts.zip")

def order_robot(order):
    page = browser.page()
    page.click("button:text('Order')")
    receipt = page.locator("#container").inner_html()
    return receipt

def clean_up():
    shutil.rmtree("./output/receipts")
    shutil.rmtree("./output/screenshots")
