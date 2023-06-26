import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from difflib import SequenceMatcher
import os


# Set up the Selenium driver
chrome_options = Options()
chrome_options.add_argument("--headless")  # Uncomment this line to run Chrome in headless mode
service = Service("path/to/chromedriver")  # Replace with the path to your chromedriver executable
driver = webdriver.Chrome(service=service, options=chrome_options)
# Load the Excel file
excel_file = "ASCO Publications (PO, OP, CCI, GO) for RK 4.11.23.xlsx"
sheets_dict = pd.read_excel(excel_file, sheet_name=None)
all_sheets = []
for name, sheet in sheets_dict.items():
    sheet['sheet'] = name
    sheet = sheet.rename(columns=lambda x: x.split('\n')[-1])
    all_sheets.append(sheet)

df = pd.concat(all_sheets)
df.reset_index(inplace=True, drop=True)

total = df.shape[0]

'''
When using selenium and dealing with redirects with this headless browser, asco would occasionally throw an intermediary page
that checked whether the system was a bot and asked the system to press the resume button, we are then bypassing this
bot checking feature through the following function.
'''

def skip_resume_button():
    try:
        resume_button = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.NAME, "resume"))
        )
        resume_button.click()
    except:
        pass  # If there's no resume button, just continue with the current page
        
'''
The following function queries the asco database by a given doi, forcing a redirect to the page with the article pdf link accesible,
then scrapes and returns that pdf link
'''

def find_pdf_url(title, doi):
    modified_doi = doi.replace("/","%2f") #corrects for slash handling such that it doesn't treat the information after the slash as a subdirectory
    article_url = f"https://oce.ovid.com/search?q=(DOI:{modified_doi})&st=false"
    url = ""

    # Use Selenium to handle the JavaScript redirect
    driver.get(article_url)
    
    skip_resume_button()
    soup = BeautifulSoup(driver.page_source, "html.parser")
    
    
    '''
    The following conditional is a workaround to test whether the first redirect is directly to the page which has the article pdf available,
    if it's an intermediate page we find that the html title tag doesn't match the article title. In this case that means our query with
    the given doi yielded multiple results (which are the same result duplicated for some odd reason), in which case we are taking the first result
    and forcing a redirect to that page
    '''
    
    if (SequenceMatcher(a=soup.title.getText(),b=title).ratio() < 0.5):
        print("\tMultiple Results Found, Redirecting to First Result")
        url = "https://oce.ovid.com" + soup.find("a", href=lambda href: href and href.startswith("/article/"))["href"]
        driver.get(url)
        skip_resume_button()
        soup = BeautifulSoup(driver.page_source, "html.parser")
        # Find PDF URLs within the page source

    for iframe in soup.find_all("iframe"):
        if "src" in iframe.attrs and iframe["src"].endswith(".pdf"):
            return iframe["src"]

searched = 0
failed_articles = [[],[]]
errors = ["Unable to be found","No link found for PDF"]
# Loop through the articles and search for their DOIs and PMIDs

for index, row in df.iterrows(): #ITerating through the rows of the excel file
    title = row["Title"]
    doi = row["DOI"]
    pdf_url=""
    sheet = row["sheet"]
    
    try:
        searched+=1
        pdf_url = find_pdf_url(title,doi)
    except:
        failed_articles[0].append(title)
        print(f"Doc {searched}/{total}: Unable to be found: {title}\n # of failed articles: {len(failed_articles[0])+len(failed_articles[1])}")
        continue
    try:
        response = requests.get(pdf_url[pdf_url.find("https://"):])
    except:
        failed_articles[1].append(title)
        print(f"Doc {searched}/{total}: No link found for PDF for {title}\n # of failed articles: {len(failed_articles[0])+len(failed_articles[1])}")
        continue
    
    with open(f"{os.getcwd()}/{sheet}/{doi[doi.find('/')+1:]}.pdf", "wb") as f:
        f.write(response.content)
        print(f"Doc {searched}/{total}: Downloaded {title}: {doi[doi.find('/')+1:]}.pdf")
        
# Clean up the driver
driver.quit()

print("The following articles failed to download for one reason or another:")
for error_num, article_list in enumerate(failed_articles):
    print(f"{errors[error_num]} for the following articles...")
    for article in article_list:
        print(article)
