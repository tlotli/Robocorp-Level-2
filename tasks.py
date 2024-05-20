from robocorp.tasks import task
from robocorp import browser
# from robocorp.workitems import ApplicationException

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from pathlib import Path

import zipfile

# from RPA.Browser.Selenium import Selenium

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP
    archive of the receipts and the images.
    """
    # browser.configure(slowmo=400)
    
    navigate_to_the_robot_website()
    complete_log_in_form()
    navigate_to_the_orders_page()
    # close_terms_and_condition_modal()
    download_order_csv()
    orders = read_csv_contents()
    
    for order in orders:
        populate_orders_form(order)

def navigate_to_the_robot_website():
    """Navigates to the robot spares website"""
    browser.goto('https://robotsparebinindustries.com')

def complete_log_in_form():
    """Enters log credentials and navigates to the website"""
    page = browser.page()
    page.fill('#username', 'maria')      
    page.fill('#password', 'thoushallnotpass')
    page.click("button:text('Log in')")  

def navigate_to_the_orders_page():
    """Clicks on the orders tab and navigates to the orders page"""
    page = browser.page()
    page.locator("xpath=//a[contains(@href, '#/robot-order')]").click()

# def close_terms_and_condition_modal():
#     """Close the terms and condition modal"""    
#     page = browser.page()
#     page.locator(".btn-dark").click()

def download_order_csv():
    """Downloads the orders CSV file"""
    http = HTTP()
    http.download('https://robotsparebinindustries.com/orders.csv', overwrite=True)     

def read_csv_contents():
    """Reads the orders CSV file"""
    tables = Tables()
    orders = tables.read_table_from_csv(path='orders.csv', header=True)
    return orders


def populate_orders_form(order):
    """Populate orders form from CSV"""
    page = browser.page() 
    page.click("button:text('OK')")
    page.select_option('#head', order['Head'])
    page.locator(f"xpath=//input[@type='radio' and @value='{order['Body']}']").check()
    page.locator("xpath=//label[contains(text(), '3. Legs:')]/following-sibling::input").fill(str(order['Legs']))
    page.fill('#address', order['Address'])
    
    submit_order(page, order)

def submit_order(page, order, retries=5):
    """Submit the order and retry if it fails"""
    for attempt in range(retries):
        page.click("button:text('Order')", force=True)
        
        if not is_error_present(page):
            break  # Exit loop if no error is present
        else:
            print(f"Attempt {attempt + 1} failed. Retrying...")

    order_id = order['Order number']
    capture_order_screen_shot(order_id, page)

def is_error_present(page):
    """Check if there is an error message on the page"""
    error_locator = page.locator("xpath=//div[contains(@class, 'alert alert-danger')]")
    return error_locator.is_visible()

def capture_order_screen_shot(order_id, page):    
    """Take snapshots of the order receipt and robot preview"""
    receipt_screenshot_selector = "#receipt"
    robot_screenshot_selector = "#robot-preview"

    receipt_screenshot_path = page.screenshot(path=f"output/{receipt_screenshot_selector.strip('#')}._.{order_id}.png")
    # robot_screenshot = page.screenshot(path=f"output/{robot_screenshot_selector.strip('#')}._.{order_id}.png")
    
    get_receipt_screenshot = Path((f"output/{receipt_screenshot_selector.strip('#')}._.{order_id}.png"))
    embed_screenshot_to_receipt(get_receipt_screenshot, order_id)
    page.locator("#order-another").click(force=True)


def embed_screenshot_to_receipt(receipt_screenshot,order_id):
    """Embed the screenshot of the robot to the PDF receipt"""
    doc = PDF()
    
    stored_pdf_path = f"output/receipt_{order_id}.pdf"
    
    doc.add_files_to_pdf(
        files=[
            (receipt_screenshot),
            # robot_screenshot
        ],
        target_document=stored_pdf_path
        )
    
    archive_receipts(stored_pdf_path)
    
def archive_receipts(stored_pdf_path):    
    """Create ZIP archive of the receipts"""
    zip_path = Path("output/receipts_archive.zip")
    
    # Open the zip file in append mode, create if it doesn't exist
    with zipfile.ZipFile(zip_path, 'a') as zipf:
        # Add the stored PDF to the zip archive
        zipf.write(stored_pdf_path, arcname=Path(stored_pdf_path).name)
    
    
    