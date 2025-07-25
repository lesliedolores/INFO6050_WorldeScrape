#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
INF 6050 - Final Project: Wordle Tracker with Multi-Sheet Logging

Scrapes today's and yesterday's Wordle answers and saves each to its own tab
in a single Excel file using pandas and openpyxl.
Locate the Excel file location by pasting in the console: 
    os.getcwd()
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import date
import os
import re
import csv
from openpyxl import load_workbook
from pandas import ExcelWriter

# === File Path ===
EXCEL_FILE = 'wordle_words_today_combined.xlsx'
print(" Excel file will be saved to:", os.path.abspath(EXCEL_FILE))

URL = 'https://www.tomsguide.com/news/what-is-todays-wordle-answer'

# === Load a specific sheet or create new DataFrame ===
def load_sheet(sheet_name):
    if os.path.exists(EXCEL_FILE):
        try:
            return pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
        except ValueError:
            return pd.DataFrame(columns=["Date", "Word"])
    else:
        return pd.DataFrame(columns=["Date", "Word"])

# === Save word to correct sheet, preserving others ===
def save_word(word, source):
    sheet_name = "Today" if source == "today" else "Yesterday"
    log_date = date.today().isoformat()

    # Load all sheets (to preserve when rewriting)
    if os.path.exists(EXCEL_FILE):
        all_sheets = pd.read_excel(EXCEL_FILE, sheet_name=None)
    else:
        all_sheets = {}

    # Get or initialize the target sheet
    sheet_df = all_sheets.get(sheet_name, pd.DataFrame(columns=["Date", "Word"]))

    # Avoid duplicates
    if (sheet_df["Date"] == log_date).any():
        print(f"ℹ️ Word already logged for {log_date} in sheet '{sheet_name}'.")
        return

    # Append the new row
    new_row = pd.DataFrame([{"Date": log_date, "Word": word}])
    sheet_df = pd.concat([sheet_df, new_row], ignore_index=True)
    all_sheets[sheet_name] = sheet_df

    # Write all sheets back into the Excel file
    with ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='w') as writer:
        for name, frame in all_sheets.items():
            frame.to_excel(writer, sheet_name=name, index=False)

    print(f" Saved '{word.upper()}' to sheet '{sheet_name}' for {log_date}")

# === Fetch today's Wordle word ===
def fetch_todays_word():
    print(" Trying to fetch TODAY'S Wordle word...")
    resp = requests.get(URL)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, 'html.parser')

    section = soup.find('h2', id='section-today-s-wordle-answer')
    if not section:
        raise RuntimeError("X Could not find section for today's Wordle.")

    # Get second <p> after heading
    p_tags_found = 0
    next_tag = section
    while next_tag and p_tags_found < 2:
        next_tag = next_tag.find_next_sibling()
        if next_tag and next_tag.name == 'p':
            p_tags_found += 1

    print("🛠️  Debugging nearby sibling tags:")
    debug = section
    for i in range(5):
        debug = debug.find_next_sibling()
        if debug:
            snippet = debug.get_text(strip=True)[:80] if debug.name == 'p' else ''
            print(f"  {i+1}. <{debug.name}>: {snippet}")

    if next_tag and next_tag.name == 'p':
        text = next_tag.get_text()
        match = re.search(r'\b([A-Z]{5})\b', text)
        if match:
            word = match.group(1).lower()
            print(f" Found TODAY'S Wordle word: {word.upper()}")
            return word, "today"

    raise RuntimeError("X Could not extract today's Wordle word.")

# === Fetch yesterday's Wordle word ===
def fetch_yesterdays_word():
    print("🌐 Trying to fetch YESTERDAY'S Wordle word...")
    resp = requests.get(URL)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, 'html.parser')

    section = soup.find('h2', id='section-yesterday-s-wordle-answer')
    if not section:
        raise RuntimeError("X Could not find section for yesterday's Wordle.")

    p_tag = section.find_next_sibling('p')
    if p_tag:
        text = p_tag.get_text()
        match = re.search(r'\b([A-Z]{5})\b', text)
        if match:
            word = match.group(1).lower()
            print(f" Found YESTERDAY'S Wordle word: {word.upper()}")
            return word, "yesterday"

    raise RuntimeError("X Could not extract yesterday's Wordle word.")

# === Run Tracker ===
def run_tracker():
    try:
        word_today, source_today = fetch_todays_word()
        save_word(word_today, source_today)
    except Exception as e:
        print(f"Failed to fetch or save today's word: {e}")

    try:
        word_yesterday, source_yesterday = fetch_yesterdays_word()
        save_word(word_yesterday, source_yesterday)
    except Exception as e:
        print(f"Failed to fetch or save yesterday's word: {e}")

# === Entry Point ===
if __name__ == '__main__':
    run_tracker()
