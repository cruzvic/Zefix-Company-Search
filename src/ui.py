import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import requests
import re
from langdetect import detect
import spacy
import html
from dotenv import load_dotenv
import os

# Load spaCy models for German, Italian, and French
nlp_de = spacy.load("de_core_news_sm")
nlp_it = spacy.load("it_core_news_sm")
nlp_fr = spacy.load("fr_core_news_sm")

load_dotenv()

# API credentials
API_BASE_URL = "https://www.zefix.admin.ch/ZefixPublicREST/api/v1"
USERNAME = os.getenv("API_USERNAME")
PASSWORD = os.getenv("API_PASSWORD")

def clean_text(text):
    """
    Clean the input text by removing HTML entities and normalizing whitespace.
    """
    # Decode HTML entities (e.g., &apos;)
    text = html.unescape(text)
    # Normalize whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    return text

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

def extract_person_names(sogc_pub):
    """
    Extract person names from sogcPub in the format [firstName1, lastName1, firstName2, lastName2].
    If first and last names cannot be distinguished, return a flat list of names [name1, name2, ...].
    """
    person_names = []
    removed_names = set()

    for pub in sogc_pub:
        message = pub.get("message", "")
        if not message:
            continue

        # Process for Eingetragene Personen (German)
        eingetragene_section = re.search(r"Eingetragene Personen.*?:", message)
        if eingetragene_section:
            start_idx = eingetragene_section.end()
            while start_idx < len(message):
                semicolon_idx = message.find(";", start_idx)
                if semicolon_idx == -1:
                    semicolon_idx = len(message)
                segment = message[start_idx:semicolon_idx].strip()

                # Skip segments with financial patterns or standalone numeric data
                if re.search(r"parts? de CHF", segment, re.IGNORECASE) or re.search(r"\d+['.]?\d*\.\d{2}", segment) or re.match(r"^\d+$", segment):
                    start_idx = semicolon_idx + 1
                    continue

                # Remove apostrophes and clean the segment
                segment = re.sub(r"'", "", segment)
                comma_splits = segment.split(",", 2)
                if len(comma_splits) >= 2:
                    last_name = comma_splits[0].strip()
                    first_name = comma_splits[1].strip()
                    person_names.extend([first_name, last_name])  # Add as [firstName, lastName]
                else:
                    person_names.append(segment.strip())  # Add as a single name if no distinction possible

                start_idx = semicolon_idx + 1

        # Process for Ausgeschiedene Personen (German)
        ausgeschiedene_section = re.search(r"Ausgeschiedene Personen.*?:", message)
        if ausgeschiedene_section:
            start_idx = ausgeschiedene_section.end()
            while start_idx < len(message):
                semicolon_idx = message.find(";", start_idx)
                if semicolon_idx == -1:
                    semicolon_idx = len(message)
                segment = message[start_idx:semicolon_idx].strip()

                # Skip segments with financial patterns or standalone numeric data
                if re.search(r"parts? de CHF", segment, re.IGNORECASE) or re.search(r"\d+['.]?\d*\.\d{2}", segment) or re.match(r"^\d+$", segment):
                    start_idx = semicolon_idx + 1
                    continue

                # Remove apostrophes and clean the segment
                segment = re.sub(r"'", "", segment)
                comma_splits = segment.split(",", 2)
                if len(comma_splits) >= 2:
                    last_name = comma_splits[0].strip()
                    first_name = comma_splits[1].strip()
                    removed_names.add(f"{first_name},{last_name}")
                else:
                    removed_names.add(segment.strip())

                start_idx = semicolon_idx + 1

        # Process for French patterns (including Personne inscrite)
        for pattern in [
            r"Titulaire.*?:", r"Associés-gérants.*?:", r"Personne inscrite.*?:", r"Personne\(s\) inscrite\(s\).*?:"
        ]:
            section = re.search(pattern, message)
            if section:
                start_idx = section.end()
                while start_idx < len(message):
                    comma_idx = message.find(",", start_idx)
                    if comma_idx == -1:
                        break
                    full_name = message[start_idx:comma_idx].strip()
                    if not re.search(r"parts? de CHF", full_name, re.IGNORECASE) and not re.search(r"\d+['.]?\d*\.\d{2}", full_name) and not re.match(r"^\d+$", full_name):
                        full_name = re.sub(r"'", "", full_name)
                        person_names.append(full_name)

                    semicolon_idx = message.find(";", start_idx)
                    if semicolon_idx == -1:
                        break
                    start_idx = semicolon_idx + 1

        # Process for Persone iscritte (Italian)
        italian_section = re.search(r"Persone iscritte.*?:", message)
        if italian_section:
            start_idx = italian_section.end()
            while start_idx < len(message):
                semicolon_idx = message.find(";", start_idx)
                if semicolon_idx == -1:
                    semicolon_idx = len(message)
                segment = message[start_idx:semicolon_idx].strip()

                # Skip segments with financial patterns or standalone numeric data
                if re.search(r"parts? de CHF", segment, re.IGNORECASE) or re.search(r"\d+['.]?\d*\.\d{2}", segment) or re.match(r"^\d+$", segment):
                    start_idx = semicolon_idx + 1
                    continue

                # Remove apostrophes and clean the segment
                segment = re.sub(r"'", "", segment)
                comma_splits = segment.split(",", 2)
                if len(comma_splits) >= 2:
                    last_name = comma_splits[0].strip()
                    first_name = comma_splits[1].strip()
                    person_names.extend([first_name, last_name])  # Add as [firstName, lastName]
                else:
                    person_names.append(segment.strip())  # Add as a single name if no distinction possible

                start_idx = semicolon_idx + 1

    # Filter out removed names
    final_names = [
        name for name in person_names
        if name not in {n.split(",")[0] for n in removed_names}  # Remove by first name if distinction exists
    ]

    # Return the names as a flat list
    return final_names if final_names else ["No names found"]


def get_company_details(uid):
    url = f"{API_BASE_URL}/company/uid/{uid}"
    response = requests.get(url, auth=(USERNAME, PASSWORD))
    
    if response.status_code == 200:
        data = response.json()  # This is a list of companies
        
        # Process all companies in the list
        company_details = []
        for company in data:
            if company["status"] == "ACTIVE":
                person_names = extract_person_names(company.get("sogcPub", []))
                company_details.append({
                    "Name": company["name"],
                    "Street": company["address"]["street"],
                    "HouseNumber": company["address"]["houseNumber"],
                    "City": company["address"]["city"],
                    "SwissZipCode": company["address"]["swissZipCode"],
                    "OwnerNames": person_names,
                })
        return company_details
    else:
        print(f"Error fetching details for UID {uid}: {response.status_code}")
        return None

def create_csv(search_key, update_progress):
    uids = search_companies(search_key)
    if not uids:
        return None

    results = []
    total_uids = len(uids)
    update_progress(0, total_uids, f"Found {total_uids} companies. Fetching details...")

    for i, uid in enumerate(uids):
        details_list = get_company_details(uid)
        if details_list:
            results.extend(details_list)

        # Update progress bar after each UID is processed
        update_progress(i + 1, total_uids, f"Processing {i + 1}/{total_uids} companies...")

    if results:
        filename = f"{search_key}_companies.csv"
        df = pd.DataFrame(results)
        df.to_csv(filename, index=False)
        return filename
    return None

# GUI Implementation
def run_search():
    search_key = search_entry.get()
    if len(search_key) < 3:
        messagebox.showerror("Error", "Search key must be at least 3 characters long.")
        return

    # Clear previous messages and progress
    status_label.config(text=f"Searching for company search key: {search_key}")
    progress_bar["value"] = 0

    # Run the search and generate the CSV
    try:
        def update_progress(current, total, message):
            progress_bar["value"] = (current / total) * 100
            progress_bar.update()
            status_label.config(text=message)

        filename = create_csv(search_key, update_progress)
        if filename:
            status_label.config(text=f"CSV file created: {filename}")
            messagebox.showinfo("Success", f"CSV file created: {filename}")
        else:
            status_label.config(text="No matching companies found.")
            messagebox.showinfo("No Results", "No matching companies found.")
    except Exception as e:
        status_label.config(text="An error occurred.")
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Set up the main window
root = tk.Tk()
root.title("Zefix Company Search")

# Add UI components
tk.Label(root, text="Enter Search Key:").pack(pady=5)
search_entry = tk.Entry(root, width=40)
search_entry.pack(pady=5)

tk.Button(root, text="Search & Export", command=run_search).pack(pady=10)

status_label = tk.Label(root, text="", fg="blue")
status_label.pack(pady=5)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=10)

# Run the GUI event loop
root.mainloop()
