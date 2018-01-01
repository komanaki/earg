#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Excel Accounting Report Generator
"""

import calendar
import locale
import argparse

import xlwings as xw

parser = argparse.ArgumentParser(description='Builds a year-long Excel workbook of financial reports with one month per sheet.')
parser.add_argument('year', type=int, help='Year of the financial report')
parser.add_argument('--template', type=str, default='template_fr.xlsx', help='Path to the template file')

def build_workbook(year, template):
    """
    Open a template workbook and build it for a given year
    """

    # Open the Excel workbook
    print("Opening Excel...")
    book = xw.Book(template)

    # Duplicate the template sheet for each month
    print("Duplicating sheets...")
    for i in range(0, 12):
        book.sheets[0].api.Copy(Before=book.sheets[0].api)

    # Delete the last sheet (the template one)
    book.sheets[-1].delete()

    # Fill out each month sheet
    print("Filling data...")
    for i in range(0, 12):
        sheet = book.sheets[i]
        num = i + 1

        periods = get_month_periods(year, num)
        sheet.name = calendar.month_name[num].title()

        # Year and month
        sheet.range('B2').value = year
        sheet.range('B3').value = calendar.month_name[num].title()

        # Periods
        sheet.range('C6').value = "'%s - %s" % (periods[0][0], periods[0][1])
        sheet.range('D6').value = "'%s - %s" % (periods[1][0], periods[1][1])
        sheet.range('E6').value = "'%s - %s" % (periods[2][0], periods[2][1])
        sheet.range('F6').value = "'%s - %s" % (periods[3][0], periods[3][1])
        sheet.range('G6').value = "'%s - %s" % (periods[4][0], periods[4][1])

        if i == 0:
            # Remove useless cells on first sheet
            sheet.range('E2:F3').clear()
        else:
            # Link to previous month balance
            previous_sheet = calendar.month_name[num - 1].title()
            sheet.range('E3').value = "='%s'!$O$3" % (previous_sheet)

    print("Done !\nYou can now 'save as' the workbook currently shown in Excel to prevent the overwriting of the template file.")

def get_month_periods(year, month):
    """
    Returns a maximum of 5 "week-like" periods for a given month
    """

    # Divide a month in "calendar weeks"
    calendar_weeks = calendar.monthcalendar(year, month)
    periods = []

    # Convert calendar weeks to a maximum of 5 periods
    if len(calendar_weeks) == 5:
        for i in range(0, 5):
            week = [i for i in calendar_weeks[i] if i != 0]
            periods.append([week[0], week[-1]])

    else:
        # Merge the smallest "extra week" into the first or last week of the month
        if calendar_weeks[0].count(0) > calendar_weeks[-1].count(0):
            periods.append([1, calendar_weeks[1][-1]])
            for i in range(2, 6):
                week = [i for i in calendar_weeks[i] if i != 0]
                periods.append([week[0], week[-1]])
        else:
            for i in range(0, 4):
                week = [i for i in calendar_weeks[i] if i != 0]
                periods.append([week[0], week[-1]])
            last_week = [i for i in calendar_weeks[-1] if i != 0]
            periods.append([calendar_weeks[-2][0], last_week[-1]])

    return periods

if __name__ == '__main__':
    locale.setlocale(locale.LC_ALL, '')
    args = parser.parse_args()
    build_workbook(args.year, args.template)