#!/usr/bin/python3
import openpyxl
import glob
import xlsxwriter
import datetime
import re
import os
import termcolor

print("Hello, nice to see you. I am starting...")
termcolor.cprint(text="8888888b.  8888888888 .d8888b.  \n888   Y88b 888       d88P  Y88b \n888    888 888       888    888 \n888   d88P 8888888   888        \n8888888P\"  888       888        \n888 T88b   888       888    888 \n888  T88b  888       Y88b  d88P \n888   T88b 888        \"Y8888P\"", color="yellow", on_color="on_white")

os.chdir(os.path.dirname(os.path.realpath(__file__)))
xlsx_name=datetime.datetime.now().strftime("%d_%m_%Y-%H_%M")
xlsx_sheer_for_write=xlsxwriter.Workbook("RFC_report_for_{xlsx_name}.xlsx".format(xlsx_name=xlsx_name))

format_border=xlsx_sheer_for_write.add_format()
format_border.set_border(1)


col_with_rfc_num=(0,3,7,11,15)

header=("Responsible Engineer name", "Team name", "Grade", "Type of activities", "Spent time (hours)")

idx_all_grade = idx_all_engineer = idx_all_team = idx_all_type = 1
total_by_grades={}; total_by_engineers={}; total_by_teams={}; total_by_types={}

total_list=[]
total_list.append(total_by_engineers); total_list.append(total_by_teams); total_list.append(total_by_grades); total_list.append(total_by_types)

total_time_worksheet=xlsx_sheer_for_write.add_worksheet("Total")
total_metrics=xlsx_sheer_for_write.add_worksheet("Total_metrics")

for current_col_with_rrc_num in col_with_rfc_num:
    total_time_worksheet.write(0, current_col_with_rrc_num, "RFC_number", format_border)
    total_time_worksheet.set_column(current_col_with_rrc_num, current_col_with_rrc_num, 11)
for current_col_total_time in range(1,18,4):
    total_time_worksheet.write(0, current_col_total_time, "Total_time", format_border)
    total_time_worksheet.set_column(current_col_total_time, current_col_total_time, 9)

total_sheet_columns=("Total_grades","Total_engineer_name","Total_team_name","Total_activities_type")

for i in range(4,17,4):
    total_time_worksheet.write(0, i, total_sheet_columns[int(i/4-1)], format_border)

def get_data_from_files(file_name):
    '''Function for read necesary data from 'Resources involved' sheet and return in via lists on lists'''
    list_of_data =[] ; list_of_data_temporary = []
    #open xlsx-file
    xlsx_file=openpyxl.load_workbook(file_name, read_only=True, data_only=True)
    try:
        active_sheet=xlsx_file['Resources involved']
    except KeyError:
        termcolor.cprint("Resource involved sheet is not found in file... Skipping the {file} file...".format(file=file_name), color="red", on_color="on_white")
        return None
    need_proceed=True
    for cell_obj in active_sheet["A2:E100"]:
        if need_proceed:
            #for cells in rows
            for cell in cell_obj:
                #if cell is empty break the all loops
                if not cell.value:
                    need_proceed=False
                    break
                #if not add cell's values to list
                else:
                    list_of_data_temporary.append(cell.value)
            #we need this to avoid adding last empty list to list
            if need_proceed:
                list_of_data.append(list_of_data_temporary)
                list_of_data_temporary=[]
            #break all loops
            else:
                break
    return list_of_data

def process_data():
    # get list of all files with xlsx-extention
    print("Trying find all xlsx-file in ./RFC/ directory")
    xlsx_files = glob.glob('RFC/*.xlsx')
    files_count=len(xlsx_files)
    if files_count>0:
        print("All OK, we have found {count} files".format(count=files_count))
    else:
        print("No any xlsx-files found... Nothing to do, existing...")
        exit()
    errors_count=0
    for idx, current_file in enumerate(xlsx_files):
        print("Processing a {file} file".format(file=current_file))
        #get list with data from current file, raise a function for cwrite these data to one file
        result_list=get_data_from_files(current_file)
        if not result_list:
            termcolor.cprint("Error during processing the {file} file".format(file=current_file), color="red", on_color="on_white")
            errors_count+=1
            continue
        create_separate_sheet_for_each_rfc(result_list, re.sub("[^0-9]", '',re.sub(".xlsx", '', re.sub(".*/", '', current_file))), idx-errors_count)

def create_separate_sheet_for_each_rfc(data, sheet_name, index):
    #counters for rows (index of rows)
    global idx_all_grade; global idx_all_engineer; global idx_all_team; global idx_all_type
    all_grades=[]; all_engineers=[]; all_teams=[]; all_types=[]
    #move the'Resources involved' sheet from each file to our our new file and create separate sheet on our file
    worksheet=xlsx_sheer_for_write.add_worksheet(name="RFC_{sheet_name}".format(sheet_name=sheet_name))
    header=("Responsible Engineer name", "Team name", "Grade", "Type of activities", "Spent time (hours)")
    worksheet.write_row(0,0,header, format_border)
    column_width=(23, 10, 6, 16, 7)
    for i in range(0,5):
        worksheet.set_column(i,i,column_width[i])
    sum_all_spent_time=0
    # add to lists our variables, later we will create set to extract only unique values
    #and wtite the data to row; also calculate the total spent time for RFC
    for idx, row in enumerate(data):
        all_grades.append(row[2]); all_engineers.append(row[0]); all_teams.append(row[1]); all_types.append(row[3])
        worksheet.write_row(idx+1, 0, row, format_border)
        sum_all_spent_time+=row[4]
    total_time_worksheet.write(index+1, 0, sheet_name, format_border)
    total_time_worksheet.write(index+1, 1, sum_all_spent_time, format_border)
    #convert list to set, we need only unique values
    all_grades=set(all_grades); all_engineers=set(all_engineers); all_teams=set(all_teams); all_types=set(all_types)
    grades_time={}; engineers_time={}; teams_time={}; types_time={}
    for current_grade in all_grades:
        grades_time[current_grade]=calculate_needed_time(data, current_grade, 2)
    for current_engineers in all_engineers:
        engineers_time[current_engineers]=calculate_needed_time(data, current_engineers, 0)
    for current_team in all_teams:
        teams_time[current_team]=calculate_needed_time(data, current_team, 1)
    for current_types in all_types:
        types_time[current_types]=calculate_needed_time(data, current_types, 3)
    idx_all_grade=write_to_totaL_sheet(grades_time, 3, sheet_name, idx_all_grade)
    idx_all_engineer = write_to_totaL_sheet(engineers_time, 7, sheet_name, idx_all_engineer)
    idx_all_team = write_to_totaL_sheet(teams_time, 11, sheet_name, idx_all_team)
    idx_all_type = write_to_totaL_sheet(types_time, 15, sheet_name, idx_all_type)


def write_to_totaL_sheet(my_dict, column, rfc_number, idx_name):
    for i in my_dict.keys():
        total_time_worksheet.write(idx_name, column, rfc_number, format_border)
        total_time_worksheet.write(idx_name, column+1, i, format_border)
        total_time_worksheet.write(idx_name, column+2, my_dict[i], format_border)
        idx_name+=1
    return idx_name


def calculate_needed_time(data, criteria, idx):
    '''Function for calculate total spent time by needed criterias'''
    temp=0
    for current_row in data:
        if criteria==current_row[idx]:
            temp+=current_row[4]
            try:
                total_list[idx][criteria]+=current_row[4]
            except:
                total_list[idx][criteria] = current_row[4]
    return temp


def create_total_of_total_sheet():
    #total_list.append(total_by_engineers); total_list.append(total_by_teams); total_list.append(total_by_grades); total_list.append(total_by_types)
    columns_with_data=("Engineer name","Team name","Grades","Type of activity")
    column_with_data_width=(20,10,7,15)
    for i in range(0,10,3):
        total_metrics.write(0,i,columns_with_data[int(i/3)], format_border)
        total_metrics.set_column(i,i, column_with_data_width[int(i/3)])
    for i in range(1,11,3):
        total_metrics.write(0, i, "Spent time", format_border)
        total_metrics.set_column(i,i, 9)
    for col, current_metric in enumerate(total_list):
        for row, current_key in enumerate(current_metric.keys()):
            total_metrics.write(row+1, col*3, current_key, format_border)
            total_metrics.write(row+1, col*3+1, current_metric[current_key], format_border)

def set_column_width():
    width_tiple=(11,18,16,17)
    for idx in range(4,17,4):
        total_time_worksheet.set_column(idx, idx, width_tiple[int(idx/4-1)])


process_data()
create_total_of_total_sheet()
#print(total_list)
set_column_width()
xlsx_sheer_for_write.close()
