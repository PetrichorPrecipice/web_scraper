# WebScraper.py - Finds PGMs (Platinum Group Metals) related R&D IPs and put into another excel spreadsheet

# Import
import xlsxwriter
import requests
from bs4 import BeautifulSoup
import re
from time import sleep
from random import randint

# Store data
names = []
patentNos = []
inventors = []
assignees = []
abstracts = []
Data = [names, patentNos, inventors, assignees, abstracts]


def user_input():
    # Ask user for search criteria
    search_query = str(input('Please input what patents you would like to search for')).upper()
    return search_query


# Worksheet
ws = xlsxwriter.Workbook('Results.xlsx')
usWb = ws.add_worksheet()


def collect_data(query):
    # Iterate for each page

    for i in range(1, 10):
        url = 'http://patft.uspto.gov/netacgi/nph-Parser?Sect1=PTO2&Sect2=HITOFF&p=' + str(
            int((i / 50) + 1)) + '&u=%2Fnetahtml%2FPTO%2Fsearch' \
                                 '-bool.html&r=' + str(
            i) + '&f=G&l=50&co1=AND&d=PTXT&s1=' + query + '.ABTX.&OS=ABST/' + query + '&RS=ABST/' + query
        print(url)
        headers = {'user-agent': 'my-app/0.0.1'}
        page = requests.get(url, timeout=5, headers=headers)
        soup = BeautifulSoup(page.text, 'html.parser')
        sleep(randint(2, 8))

        # Names

        name_finder = soup('font', {'size': '+1'})
        name = re.sub(r'<.*>|]|\[|\n\s?', " ", (str(name_finder)))
        name = name.strip()
        names.append(name)

        # Patent Number

        patentno_finder = soup('td', {'align': 'right'})
        patent_temp = str(patentno_finder[1]).replace(',', '')
        patentNo = re.findall(r'(\d{6,8}|RE\d+|PP\d+|D\d+|AI\d+|X\d+|T\d+|H\d+)', patent_temp)
        patentNos.append(patentNo[0])

        # Inventors

        inventor_finder = soup('td', {'align': 'left'})
        inventor = re.sub(r'<\w>|</\w>|</?td.*?>|\n\s?', '', str(inventor_finder[3]))
        inventors.append(inventor)

        # Assignees

        assignee_finder = soup('td', {'align': 'left'})
        assignee = re.sub(r'<\w>|</\w>|</?td.*?>|\n\s?|<\w+/>', '', str(assignee_finder[6]))
        assignees.append(assignee)

        # Abstract

        abstract_finder = soup('p')
        abstract = re.sub(r'<.*>|\n\s?', "", (str(abstract_finder[0])))
        abstract = abstract.strip()
        abstracts.append(abstract)


def worksheet_creator():
    # Fill Cells
    for i in range(0, 5):
        for row, data in enumerate(Data[i]):
            usWb.write(row + 1, i, data)
    # Make Bold

    bold = ws.add_format({'bold': True})

    # Make Headers

    usWb.write(0, 0, 'Patent Name', bold)
    usWb.write(0, 1, 'Patent Number', bold)
    usWb.write(0, 2, 'Inventor(s)', bold)
    usWb.write(0, 3, 'Assignee(s)', bold)
    usWb.write(0, 4, 'Description', bold)

    # Close

    ws.close()


collect_data(user_input())
worksheet_creator()