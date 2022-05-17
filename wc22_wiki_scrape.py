import requests
from bs4 import BeautifulSoup
import xlsxwriter

COLUMN_NAMES = ["Date", "Time", "Match"]

# Captures Wikipedia article content.
content = requests.get(
        "https://en.wikipedia.org/wiki/2022_FIFA_World_Cup"
    )

# Parses Wikipedia article HTML.
soup = BeautifulSoup(content.text, "html.parser")
tags = soup.find_all("div", {"class": "footballbox"})

matches = []
for tag in tags:
    match_dictionary = {}
    date = tag.find("span", {"class": "bday dtstart published updated"})
    match_dictionary["Date"] = date.get_text()
    time = tag.find("div", {"class": "ftime"})
    match_dictionary["Time"] = time.get_text()
    home_team = tag.find("th", {"itemprop": "homeTeam"})
    away_team = tag.find("th", {"itemprop": "awayTeam"})
    match_dictionary["Match"] = home_team.get_text(strip=True) + " vs " + away_team.get_text(strip=True)
    matches.append(match_dictionary)

# Write match information to an Excel workbook.
workbook = xlsxwriter.Workbook("World Cup 2022 Match Schedule.xlsx")
worksheet = workbook.add_worksheet()
 
for index, column in enumerate(COLUMN_NAMES):
    worksheet.write(0, index , column)

row_index = 1 # Ensures match information never overwrites column headers.
column_index = 0
for match in matches:
    for value in match.values():
        worksheet.write(row_index, column_index, value)
        column_index += 1
    row_index += 1
    column_index = 0 

workbook.close()
