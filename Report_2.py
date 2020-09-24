import random
import time
import os
import csv
import pandas as pd
from docxtpl import DocxTemplate
from PyPDF2 import PdfFileWriter, PdfFileReader
from docx2pdf import convert
import matplotlib.pyplot as plt
import seaborn as sns

sns.set()

# Source CSV - column names that must match the * that are {{***}} inside "template.docx"
csv_file = "Marks.csv"
save_path = r'C:\Users\Safee\PycharmProjects\Report\Output_Reports'
graph_path = r'C:\Users\Safee\PycharmProjects\Report\Output_Graphs'


def make_graphs_2():
    df = pd.read_csv(csv_file)
    #   We have to convert the dataframe into a dictionary
    dfd = df.to_dict(orient='records')
    #   Assigning df,column to the variable Header_List
    Header_List = list(df.columns)
    print("The headers in the file are: ")
    print(Header_List)
    #   Start subject and End subject name should match the column name in the csv/file
    Start_Subject = str(input("Starting Subject:"))
    End_Subject = str(input("End Subject:"))

    if Start_Subject and End_Subject in df.columns:
        #   Index tells that location in the list
        start_index = Header_List.index(Start_Subject)
        end_index = Header_List.index(End_Subject)
        #   Printing the location to the user
        print("{} found in Headers and is at location {}".format(Start_Subject, start_index))
        print("{} found in Headers and is at location {}\n".format(End_Subject, end_index))

    #   Once we have the start index and the end index
    #   We select the colums of the file and specifying the start index and end index+1
    subjects_headers = df.columns[start_index:end_index + 1]
    #   Then we convert that subject headers to a list
    subject_list = list(subjects_headers)
    #   Converting Input to the list
    #   We separate it by ','
    #   Then we convert it into a list
    sub_list = input("Enter Subject Headings for X axis(separated by commas):")
    sub_list = sub_list.split(',')
    sub_list = [str(x) for x in sub_list]
    print(sub_list)

    #   In this loop we go over the total number of rows
    #   We make a new list 'y'
    #   First we get the name
    for i in range(0, len(dfd)):
        y = []
        print("\n")
        print("Name: {}".format(dfd[i].get('name')))

        #   We then loop over the subject list made earlier
        #   We append the marks in the list 'y'
        for subject in subject_list:
            y.append(dfd[i].get(str(subject)))
        print(y)

        #   PLotting Graph using the data
        plt.figure(figsize=(6, 6))
        fig = plt.figure()
        ax = fig.add_axes([0, 0, 1, 1])

        plt.xticks(fontsize=20, rotation=45)
        plt.yticks(fontsize=20)
        plt.ylim(0, 100)
        plt.ylabel('Percentage(%)', fontsize=20, fontweight='bold')
        plt.title("Student Performance", fontsize=40, pad=20, loc='center')
        #   x is the list sublist and y is the marks list
        bars = ax.bar(sub_list, y, width=0.5, color='dodgerblue', align='center')

        #This is use to get the values of the bar
        for bar in bars:
            yval = bar.get_height()
            plt.text(bar.get_x() + .1, yval - 5, yval, fontsize=10,
                     fontweight='bold', color='white')

        #   How the graph will be named.
        name_format = "{}_{}_{}_{}.png".format(str(i + 1), dfd[i]['name'], dfd[i]['year'], dfd[i]['sec'])
        complete_name = os.path.join(graph_path, name_format)
        plt.savefig(complete_name, bbox_inches='tight')

        # plt.show()
        plt.close()


# def make_graph(n):
#     df=pd.read_csv(csv_file)
#     dfd = df.to_dict(orient='records')
#     plt.figure(figsize=(6, 6))
#     fig = plt.figure()
#
#     # plt.style.use('seaborn-dark-palette')
#     ax = fig.add_axes([0, 0, 1, 1])
#
#     x = ['English', 'Math', 'Science', 'Islamic', 'Arabic', 'Average']
#     y = [dfd[n].get('eng'),
#          dfd[n].get('math'),
#          dfd[n].get('sci'),
#          dfd[n].get('isl'),
#          dfd[n].get('ar'),
#          dfd[n].get('avg')]
#
#     plt.xticks(fontsize=20, rotation=45)
#     plt.yticks(fontsize=20)
#     # plt.yticks(np.arange(0,100,5))
#     plt.ylim(0, 100)
#     # plt.xlabel('Subjects',labelpad=10)
#     plt.ylabel('Percentage(%)', fontsize=20, fontweight='bold')
#     plt.title("Student Performance", fontsize=40, pad=20, loc='center')
#
#     bars = ax.bar(x, y, width=0.5, color='dodgerblue', align='center')
#
#     for bar in bars:
#         yval = bar.get_height()
#         plt.text(bar.get_x() + .1, yval - 5, yval, fontsize=20,
#                  fontweight='bold', color='white')
#     name_format = "{}_{}_{}_{}.png".format(str(i + 1), dfd[i]['name'], dfd[i]['year'], dfd[i]['sec'])
#     complete_name = os.path.join(graph_path, name_format)
#     plt.savefig(complete_name, bbox_inches='tight')
#     plt.close()


def make_reports(n):
    # tpl = DocxTemplate("Template.docx")  # In same directory
    tpl = DocxTemplate("Template2.docx")  # In same directory
    df = pd.read_csv(csv_file)
    context = df.to_dict(orient='records')

    graph_format = "{}_{}_{}_{}.png".format(str(n + 1), context[n]['name'], context[n]['year'], context[n]['sec'])
    complete_graph_location = os.path.join(graph_path, graph_format)
    # tpl.replace_pic('test.png','graph '+str(n)+'.png')

    #       Replacing the graph with the test.png in the template
    tpl.replace_pic('test.png', complete_graph_location)
    tpl.render(context[n])

    name_format = "{}_{}_{}_{}.docx".format(str(n + 1), context[n]['name'], context[n]['year'], context[n]['sec'])
    complete_name = os.path.join(save_path, name_format)
    tpl.save(complete_name)

    # Wait a random time - increase to (60,180) for real production run.
    # time.sleep(random.randint(1, 2))


### Main ###

df2 = len(pd.read_csv(csv_file))
print("Total Entries{}\n".format(df2))
make_graphs_2()
print("")
print("All Graphs Are Generated")

for i in range(0, df2):
    print("Graphs {} is Saved in Output_Grpahs".format(i))
    # make_graph(i)

    print("Generating Report {} with Graph {}. Please Wait  ".format(i, i))
    make_reports(i)
    print("Converting Report {} to PDF".format(i))
    # convert(save_path)
    print("\n")

print("Reports are Generated")
