import requests
import pandas as pd
import os
from dotenv import load_dotenv

load_dotenv()

# API credentials
API_BASE_URL = "https://www.zefix.admin.ch/ZefixPublicREST/api/v1"
USERNAME = os.getenv("API_USERNAME")
PASSWORD = os.getenv("API_PASSWORD")

# Function to search companies
def search_companies(search_key):
    url = f"{API_BASE_URL}/company/search"
    headers = {"Content-Type": "application/json"}
    results = []

    # Add wildcards (*) to the search key for flexible matching
    wildcard_search_key = f"*{search_key}*"

    payload = {"name": wildcard_search_key}  # Use wildcards in the name field
    response = requests.post(url, json=payload, auth=(USERNAME, PASSWORD), headers=headers)

    if response.status_code == 200:
        data = response.json()

        # Collect UIDs of companies with "ACTIVE" status
        for company in data:
            if company["status"] == "ACTIVE":  # Filter by active status
                results.append(company["uid"])
    else:
        print(f"Error: {response.status_code}, {response.text}")

    return results

# Function to get details of a company by UID
def get_company_details(uid):
    url = f"{API_BASE_URL}/company/uid/{uid}"
    response = requests.get(url, auth=(USERNAME, PASSWORD))
    
    if response.status_code == 200:
        data = response.json()  # This is a list of companies
        
        # Process all companies in the list
        company_details = []
        for company in data:
            if company["status"] == "ACTIVE":
                company_details.append({
                    "Name": company["name"],
                    "Street": company["address"]["street"],
                    "HouseNumber": company["address"]["houseNumber"],
                    "City": company["address"]["city"],
                    "SwissZipCode": company["address"]["swissZipCode"],
                })
        return company_details
    else:
        print(f"Error fetching details for UID {uid}: {response.status_code}")
        return None


# Function to process and export to Excel
def create_csv(search_key):
    print(f"Searching for companies with key: {search_key}")
    uids = search_companies(search_key)
    print(f"Found {len(uids)} matching companies.")

    # Collect results
    results = []
    for uid in uids:
        details_list = get_company_details(uid)
        if details_list:
            results.extend(details_list)  # Add all details from this UID

    if results:
        filename = f"{search_key}_companies.csv"
        
        # Use pandas to create a DataFrame and save as CSV
        df = pd.DataFrame(results)
        df.to_csv(filename, index=False)  # Write to CSV without row indices
        
        print(f"CSV file created: {filename}")
    else:
        print("No matching companies found.")



# Main Function
if __name__ == "__main__":
    search_key = input("Enter the company search key: ")
    create_csv(search_key)
