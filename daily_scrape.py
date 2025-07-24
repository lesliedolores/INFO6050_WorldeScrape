#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jul 22 22:01:22 2025

@author: lesliedoloresgarcia

@Course: INF 6050
@University: Wayne State University
@Assignment: Scraping the daily word of the day
    
@Python Version: 3.11

@Description: scraping the daily word of the day
Locate the Excel file location by pasting in the console: 
    os.getcwd()
"""


import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import date
import os
import re

# Excel file to save to
EXCEL_FILE = 'wordle_words_today.xlsx'
print(" Excel file will be saved to:", os.path.abspath(EXCEL_FILE))

def fetch_todays_word():
    """Scrape today's Wordle word using the second <p> after heading."""
    url = 'https://www.tomsguide.com/news/what-is-todays-wordle-answer'
    print(f"🌐 Fetching page: {url}")
    resp = requests.get(url)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, 'html.parser')

    # Find the specific heading section
    section = soup.find('h2', id='section-today-s-wordle-answer')
    if not section:
        raise RuntimeError(" X Could not find <h2> with id='section-today-s-wordle-answer'.")

    # Traverse to the second <p> tag after the section
    p_tags_found = 0
    next_tag = section
    while next_tag and p_tags_found < 2:
        next_tag = next_tag.find_next_sibling()
        if next_tag and next_tag.name == 'p':
            p_tags_found += 1

    # Debug: print 6 next siblings to confirm placement
    print("🛠️  Debugging nearby sibling tags:")
    debug_tag = section
    for i in range(6):
        debug_tag = debug_tag.find_next_sibling()
        if debug_tag:
            text_preview = debug_tag.get_text(strip=True)[:80] if debug_tag.name == 'p' else ''
            print(f"  {i+1}. <{debug_tag.name}>: {text_preview}")

    # Use regex to find a 5-letter capitalized word
    if next_tag and next_tag.name == 'p':
        full_text = next_tag.get_text()
        match = re.search(r'\b([A-Z]{5})\b', full_text)
        if match:
            word = match.group(1)
            print(f" Found Wordle word: {word}")
            return word.lower()

    raise RuntimeError(" X Could not find a 5-letter capitalized word in the second <p>.")

def load_excel():
    """Load or create the Excel log file."""
    if os.path.exists(EXCEL_FILE):
        return pd.read_excel(EXCEL_FILE)
    else:
        return pd.DataFrame(columns=["Date", "Word"])

def save_todays_word():
    """Fetch, check for duplicates, and save today's Wordle word."""
    today = date.today().isoformat()
    word = fetch_todays_word()

    df = load_excel()
    if (df["Date"] == today).any():
        print("ℹ️ Word already logged for today.")
        return

    new_row = pd.DataFrame([{"Date": today, "Word": word}])
    df = pd.concat([df, new_row], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)
    print(f" Word saved to Excel: {word.upper()} on {today}")

if __name__ == '__main__':
    try:
        save_todays_word()
    except Exception as e:
        print(" Error:", e)
