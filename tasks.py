from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

from PIL import Image
import io
import os

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a   PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=100,
    )
    open_robot_order_website()    
    orders = get_orders()
    for order in orders:
        close_annoying_modal()
        fill_the_form(order) 
        pdf_path = store_receipt_as_pdf(order["Order number"])   
        sc_path = screenshot_robot(order["Order number"])
        embed_screenshot_to_receipt(sc_path,pdf_path)
        archive_receipts()

def open_robot_order_website():
    """Open given url"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    """Donwload orders ande return a table with data"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv",overwrite=True)
    lib = Tables()
    return lib.read_table_from_csv("orders.csv")

def close_annoying_modal():
    """Close that ****"""
    page = browser.page()
    page.click("button:text('OK')")
    
def fill_the_form(order):
    """Fill web form to order"""
    page = browser.page()
    page.fill("//*[@placeholder='Enter the part number for the legs']",order["Legs"])
    page.fill("#address",order["Address"])
    page.select_option("#head", str(order["Head"]))
    page.click(f"#id-body-{str(order['Body'])}")
    page.click("#preview")
    page.click("#order")
    while page.is_visible("#order-completion") == False:
        page.click("#order")
    
    
def store_receipt_as_pdf(order_number):
    """Create blank pdf order"""    
    pdf = PDF()
    pdf_path = f"output/receipts/order-{order_number}.pdf"
    pdf.html_to_pdf("",output_path=pdf_path)
    return pdf_path

def screenshot_robot(order_number):
    """Take page screenshoot"""
    page = browser.page()
    image = Image.open(io.BytesIO(page.screenshot()))    
    os.makedirs("output/receipts/images/", exist_ok=True)
    img_path = f"output/receipts/images/img_order_{order_number}.png"
    image.save(img_path)
    page.click("#order-another")
    return img_path

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Embed image into pdf"""
    pdf = PDF()
    pdf.add_files_to_pdf(files=[f'{screenshot}:x=0, y=0',],target_document=pdf_file)

def archive_receipts():
    """zip pdf receipts"""
    lib = Archive()
    lib.archive_folder_with_zip('output/receipts', 'output/receipts.zip', include='*.pdf')
    