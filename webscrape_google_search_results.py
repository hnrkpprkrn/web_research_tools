# Importing required libraries
import requests
from bs4 import BeautifulSoup
from datetime import date
import pandas as pd

# Function to scrape Google search results for a given query
def scrape_google_search_results(query):
    # The base URL for Google search
    base_url = "https://www.google.com/search?q="
    
    # Defining user-agent header for the request (simulating a web browser)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.66 Safari/537.36"
    }
    
    # List to store the scraped results
    results = []
    
    # Looping through 5 pages of search results (each page shows 10 results)
    for page in range(5):
        # Calculating the start index for the current page
        page_num = page * 10
        
        # Sending a GET request to Google search with the specified query and page number
        response = requests.get(f"{base_url}{query}&start={page_num}", headers=headers)
        
        # Parsing the HTML content of the response using BeautifulSoup
        soup = BeautifulSoup(response.text, "html.parser")
        
        # Finding all the links from the search results using their class name
        links = soup.find_all("div", class_="yuRUbf")
        
        # Extracting and storing the URLs of the search results
        for link in links:
            url = link.find("a").get("href")
            results.append(url)

    # Generating a file name for storing the results in XLSX format
    file_name = "google_results_" + query + "_" + str(date.today()) + ".xlsx"
    
    # Convert the list of links to a DataFrame
    df = pd.DataFrame(results, columns=['Link'])
    
    # Write the DataFrame to an Excel file
    df.to_excel(file_name, index=False, sheet_name='Tabelle1')

# Usage
# Taking input from the user for the Google search query
search_query = input("Bitte gib den Google-Suchbegriff ein: ")

# Calling the function to scrape Google search results for the provided query
websites = scrape_google_search_results(search_query)
