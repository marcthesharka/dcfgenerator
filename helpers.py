import csv
import urllib.request

from flask import redirect, render_template, request, session
from functools import wraps

# Create a list of dictionaries to store CSV file
incomestatementdict = []

# Create a list of dictionaries to store CSV file
balancesheetdict = []

# Create a list of dictionaries to store CSV file
cashflowdict = []

def lookupis(company):
    """Look up Income Statement for company."""

    # reject company if it starts with caret
    if company.startswith("^"):
        return None

    # Reject company if it contains comma
    if "," in company:
        return None

    # Query Intrinio for Income Statement
    try:

        # GET CSV Income Statement
        url = f"https://financials.morningstar.com/ajax/ReportProcess4CSV.html?t={company}&reportType=is&period=12&dataType=A&order=asc&columnYear=5&number=3"
        webpage = urllib.request.urlopen(url)

        # Read CSV
        incomestatement = csv.reader(webpage.read().decode("utf-8").splitlines())

        # Ignore first row
        next(incomestatement)

        # Ignore second row
        next(incomestatement)

        totalColumns = 7
        # Iterate through rows in CSV
        for line in incomestatement:
            if(len(line) < totalColumns):
                continue
            lineitem = line[0]
            yr1 = line[1]
            yr2 = line[2]
            yr3 = line[3]
            yr4 = line[4]
            yr5 = line[5]
            ttm = line[6]

            # Append items to dictionary
            incomestatementdict.append({"lineitem": lineitem, "yr1": yr1, "yr2": yr2, "yr3": yr3, "yr4": yr4, "yr5": yr5, "ttm": ttm})

        # return the dictionary
        return incomestatementdict

    except:
        pass

def lookupbs(company):
    """Look up Balance Sheet for company."""

    # reject company if it starts with caret
    if company.startswith("^"):
        return None

    # Reject company if it contains comma
    if "," in company:
        return None

    # Query Intrinio for Balance Sheet
    try:

        # GET CSV Income Statement
        url = f"https://financials.morningstar.com/ajax/ReportProcess4CSV.html?t={company}&reportType=bs&period=12&dataType=A&order=asc&columnYear=5&number=3"
        webpage = urllib.request.urlopen(url)

        # Read CSV
        balancesheet = csv.reader(webpage.read().decode("utf-8").splitlines())

        # Ignore first row
        next(balancesheet)

        # Ignore second row
        next(balancesheet)

        totalColumns = 6
        # Iterate through rows in CSV
        for line in balancesheet:
            if(len(line) < totalColumns):
                continue
            lineitem = line[0]
            yr1 = line[1]
            yr2 = line[2]
            yr3 = line[3]
            yr4 = line[4]
            yr5 = line[5]
            #yr1 = int(line[1])
            #yr2 = int(line[2])
            #yr3 = int(line[3])
            #yr4 = int(line[4])
            #yr5 = int(line[5])

            # Append items to dictionary
            balancesheetdict.append({"lineitem": lineitem, "yr1": yr1, "yr2": yr2, "yr3": yr3, "yr4": yr4, "yr5": yr5})

        # return the dictionary
        return balancesheetdict

    except:
        pass

def lookupcf(company):
    """Look up Cash Flow for company."""

    # reject company if it starts with caret
    if company.startswith("^"):
        return None

    # Reject company if it contains comma
    if "," in company:
        return None

    # Query Intrinio for Cash Flow
    try:

        # GET CSV Cash Flow
        url = f"https://financials.morningstar.com/ajax/ReportProcess4CSV.html?t={company}&reportType=cf&period=12&dataType=A&order=asc&columnYear=5&number=3"
        webpage = urllib.request.urlopen(url)

        # Read CSV
        cashflow = csv.reader(webpage.read().decode("utf-8").splitlines())

        # Ignore first row
        next(cashflow)

        # Ignore second row
        next(cashflow)

        totalColumns = 6
        # Iterate through rows in CSV
        for line in cashflow:
            if(len(line) < totalColumns):
                continue
            lineitem = line[0]
            yr1 = line[1]
            yr2 = line[2]
            yr3 = line[3]
            yr4 = line[4]
            yr5 = line[5]
            # yr1 = int(line[1])
            # yr2 = int(line[2])
            # yr3 = int(line[3])
            # yr4 = int(line[4])
            # yr5 = int(line[5])

            # Append items to dictionary
            cashflowdict.append({"lineitem": lineitem, "yr1": yr1, "yr2": yr2, "yr3": yr3, "yr4": yr4, "yr5": yr5})

        print(cashflowdict)

        # return the dictionary
        return cashflowdict

    except:
        pass

def usd(value):
    """Formats value as USD."""
    return f"${value:,.2f}"


