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

# Language to spaCy model mapping
LANGUAGE_MODELS = {
    "de": nlp_de,
    "it": nlp_it,
    "fr": nlp_fr
}

# Blacklist of irrelevant terms
BLACKLIST_TERMS = [
    "Statuts", "Feuille", "Organe", "Publ", "TYPE", "Zweck", "FT", "AG", "Officielle Suisse du Commerce"
]

VALID_NAME_PARTICLES = {"da", "de", "van", "von", "del", "dos", "du", "di", "la", "le", "der"}


def clean_text(text):
    """
    Clean the input text by removing HTML entities and normalizing whitespace.
    """
    # Decode HTML entities (e.g., &apos;)
    text = html.unescape(text)
    # Normalize whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def filter_names(names):
    """
    Filter out non-names based on the blacklist, structural checks, and specific fragment cleaning,
    while preserving valid name particles like 'da', 'de', 'von', etc.
    """
    filtered_names = []
    for name in names:
        original_name = name

        # Remove blacklisted terms entirely
        for term in BLACKLIST_TERMS:
            if term in name:
                print(f"Removed '{term}' from '{original_name}' due to blacklist.")
                name = name.replace(term, "").strip()

        # Remove specific patterns like "CHE-328.335.041</" or similar
        name = re.sub(r"CHE-\d+\.\d+\.\d+<.*?>", "", name)  # Remove "CHE-XXX.XXX.XXX</...>"
        name = re.sub(r"CHE-\d+\.\d+\.\d+", "", name)  # Remove "CHE-XXX.XXX.XXX"
        name = re.sub(r"<[^>]*>", "", name)  # Remove tags like "<...>"
        name = re.sub(r'=[^>]*>', '', name)  # Remove fragments like "=\"...\">"

        # Remove dangling semicolons, extra whitespace, and ensure the name is valid
        name = re.sub(r'\s+;\s*$', '', name)  # Remove trailing semicolons
        name = re.sub(r'^[;\s]+|[;\s]+$', '', name)  # Remove leading/trailing semicolons or whitespace
        name = re.sub(r'\s+', ' ', name).strip()  # Normalize whitespace

        # Allow valid particles in names (e.g., 'da', 'de', 'von')
        words = name.split()
        valid_particles = set(VALID_NAME_PARTICLES)
        valid = all(
            word.isalpha() or word.lower() in valid_particles or "-" in word
            for word in words
        )

        # Ensure name isn't empty after cleaning
        if not valid or not words or not name.strip():
            print(f"Removed '{original_name}' for not meeting name criteria.")
            continue

        # Reconstruct the cleaned and validated name
        reconstructed_name = " ".join(words)
        filtered_names.append(reconstructed_name)

    return filtered_names

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

def detect_language_and_extract(text):
    """
    Detect the language of the text and extract person names using the appropriate model.
    """

    text = clean_text(text) # Clean text before processing
    try:
        lang = detect(text)  # Detect the language
        model = LANGUAGE_MODELS.get(lang)
        if model:
            names = extract_names_with_model(text, model)
            return filter_names(names)
        else:
            print(f"Unsupported language detected: {lang}")
            return []
    except Exception as e:
        print(f"Language detection failed for text: {text}. Error: {str(e)}")
        return []

def extract_names_with_model(text, model):
    """
    Extract names using a specific spaCy language model, ensuring multi-word names are preserved.
    """
    doc = model(text)
    names = []
    for ent in doc.ents:
        if ent.label_ == "PER":  # Extract PERSON entities
            # Add additional logic to check and preserve particles
            cleaned_name = " ".join([
                token.text for token in ent
                if token.text.lower() not in BLACKLIST_TERMS
            ])
            names.append(cleaned_name.strip())
    return names


def extract_person_names(sogc_pub):
    """
    Extract person names from sogcPub using multilingual processing.
    """
    person_names = []
    for pub in sogc_pub:
        message = pub.get("message", "")
        if message:
            names = detect_language_and_extract(message)  # Detect language and extract names
            person_names.extend(names)

    # Deduplicate and join names
    unique_names = list(set(person_names))
    return "; ".join(unique_names) if unique_names else "No names found"

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
