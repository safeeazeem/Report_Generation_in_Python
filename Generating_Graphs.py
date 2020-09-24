import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
sns.set()

save_path = r'C:\Users\Safee\PycharmProjects\Report\Output_Graphs'


df = pd.read_csv('Marks.csv')

# COnverting to dictionary
dfd = df.to_dict(orient='records')

df2 = len(pd.read_csv('Marks.csv'))
print("CSV HAS: {} Entries".format(df2))
for i in range(0, df2):
    plt.figure(figsize=(6, 6))
    fig = plt.figure()

    # plt.style.use('seaborn-dark-palette')
    ax = fig.add_axes([0, 0, 1, 1])

    x = ['English', 'Math', 'Science', 'Islamic', 'Arabic', 'Average']
    y = [dfd[i].get('eng'),
         dfd[i].get('math'),
         dfd[i].get('sci'),
         dfd[i].get('isl'),
         dfd[i].get('ar'),
         dfd[i].get('avg')]

    plt.xticks(fontsize=20, rotation=45)
    plt.yticks(fontsize=20)
    # plt.yticks(np.arange(0,100,5))
    plt.ylim(0, 100)
    # plt.xlabel('Subjects',labelpad=10)
    plt.ylabel('Percentage(%)', fontsize=20, fontweight='bold')
    plt.title("Student Performance", fontsize=40, pad=20, loc='center')

    bars = ax.bar(x, y, width=0.5, color='dodgerblue', align='center')

    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + .1, yval - 5, yval, fontsize=10,
                 fontweight='bold', color='white')

    name_format = "{}_{}_{}_{}.png".format(str(i + 1), dfd[i]['name'], dfd[i]['year'], dfd[i]['sec'])
    complete_name = os.path.join(save_path, name_format)
    plt.savefig(complete_name, bbox_inches='tight')


    # plt.show()
    plt.close()
print("{}_Graphs are Saved in Output_Grpahs".format(df2))