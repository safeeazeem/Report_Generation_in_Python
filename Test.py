import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
sns.set()
save_path = r'C:\Users\Safee\PycharmProjects\Report\Output_Graphs'

csv_file = "Marks2.csv"
df = pd.read_csv(csv_file)
dfd = df.to_dict(orient='records')
# for loc in range(len(df.columns)):
#     print("{}_{}\t".format(loc,df.columns[loc]))

Header_List = list(df.columns)
print("The headers in the file are: ")
print(Header_List)
Start_Subject = str(input("Starting Subject:"))
End_Subject = str(input("End Subject:"))

if Start_Subject and End_Subject in df.columns:
    print("{} found in Headers".format(Start_Subject))
    print("{} found in Headers".format(End_Subject))
    start_index= Header_List.index(Start_Subject)
    end_index= Header_List.index(End_Subject)
    print(start_index)
    print(end_index)

subjects_headers = df.columns[start_index:end_index+1]
subject_list = list(subjects_headers)

# Converting Input to the list
sub_list = input("Enter Subject list:")
sub_list = sub_list.split(',')
sub_list =[str(x) for x in sub_list]
print(sub_list)

#going through the entries
for i in range(0,len(dfd)):
    y = []
    print("\n")
    print("Name: {}".format(dfd[i].get('name')))
    #Getting Values of each subject
    for subject in subject_list:
        y.append(dfd[i].get(str(subject)))
    print(y)

    plt.figure(figsize=(6, 6))
    fig = plt.figure()

    # plt.style.use('seaborn-dark-palette')
    ax = fig.add_axes([0, 0, 1, 1])

    plt.xticks(fontsize=20, rotation=45)
    plt.yticks(fontsize=20)
    plt.ylim(0, 100)
    plt.ylabel('Percentage(%)', fontsize=20, fontweight='bold')
    plt.title("Student Performance", fontsize=40, pad=20, loc='center')

    bars = ax.bar(sub_list, y, width=0.5, color='dodgerblue', align='center')

    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + .1, yval - 5, yval, fontsize=10,
                 fontweight='bold', color='white')

    name_format = "{}_{}_{}_{}.png".format(str(i + 1), dfd[i]['name'], dfd[i]['year'], dfd[i]['sec'])
    complete_name = os.path.join(save_path, name_format)
    plt.savefig(complete_name, bbox_inches='tight')

    # plt.show()
    plt.close()

