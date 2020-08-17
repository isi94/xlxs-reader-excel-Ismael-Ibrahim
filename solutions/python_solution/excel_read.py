# Author: Ismael Ibrahim
# Date: 17.08.2020
# Content: Excel xlsx Parser, calculates the sum up the skill value from
# a given date

import datetime, xlrd

#compares the title row with the given param year and month
#returns the column index
def getCol(worksheet, wb_date_mode, year, month):
    cols_size = worksheet.ncols

    for col in range(cols_size) :
        title = worksheet.cell_value(0,col)
        if (col > 2):
            try:
                title_as_date = \
                    datetime.datetime(*xlrd.xldate_as_tuple(title, wb_date_mode))
                if(int(title_as_date.month) == month
                        and int(title_as_date.year) == year):
                    return col
            except:
                print("Unexpected cell content")
    return 0


#sum up the skill from the given column(date) if the skill attribute matches
#returns the sum of the skill
def getSumSkill(worksheet,col, requested_skill_name):
    rows_size = worksheet.nrows
    sum_of_skill = 0.0
    found_skill = "n"
    for row in range(rows_size):
        skill_name = worksheet.cell_value(row,1)
        if(skill_name == requested_skill_name):
            found_skill = "y"
            try:
                skill_val = float(worksheet.cell_value(row,col))
                sum_of_skill = sum_of_skill + skill_val
            except:
                skill_val = 0
    if(found_skill == "n"):
        return found_skill
    return sum_of_skill


def main():
    filename = "Site_Capacity.xlsx"
    workbook_index = 0

    try:
        workbook = xlrd.open_workbook(filename)
    except:
        print("File could not be opened")

    try:
        worksheet = workbook.sheet_by_index(workbook_index)
    except:
        print("No sheet with this index")
    quit_or_next = ""

    while(quit_or_next is not "q"):
        year = "";
        month = "";
        skill = input("Skill: ")

        while(year == ""):
            try:
                year = int(input("Year: "))
            except:
                print("input Year is no digit!")

        while(month == ""):
            try:
                month = int(input("Month: "))
            except:
                print("input Month is no digit!")

        column_index = getCol(worksheet,workbook.datemode,year,month)
        if(column_index == 0):
            print("Date not found\n")
        else:
            sum_of_skill = getSumSkill(worksheet,column_index, skill)
            if(sum_of_skill == "n"):
                print("Skill not found\n")
            else:
                print("Skill result: ",sum_of_skill, "\n")


        quit_or_next = input("Enter \"q\" to quit or any to repeat: ")
    workbook.release_resources()

if __name__ == "__main__":
    main()


