import datetime
import glob
import os
import shutil
import numpy as np
import pandas as pd
from pathlib import Path
import matplotlib.pyplot as plt
import seaborn as sns

'''
Load data
'''

# set up
path = Path('/Users/meganmason491/Documents/snowex/2020/timeseries/qaqc_pits/pits_csv_edited')
fname_pkl = 'SNEX20_TS_IB_newSnow_v01.pkl'

# load dataframe (from pickel file)
df = pd.read_pickle(Path.joinpath(path, fname_pkl))

'''
Create Weeks Dictionary
'''

weeks_dict = dict(zip(df['Week No.'], df['Date/Local Time']))


def week_range(d, weekday):
    days_ahead = d.weekday() - weekday
    if days_ahead <= 0: # Target day already happened this week
        days_ahead += 7
    return d + datetime.timedelta(days_ahead)

d = datetime.date(2020, 3, 4)
monday = week_range(d, 0) # 0 = Monday, 1=Tuesday, 2=Wednesday...
print(next_monday)


'''
plotting interval board data
'''

# Nested barplot by locations
grouped_df = df.groupby('Location')

for name, group in grouped_df:
    print(f"Processing: {name}")

    g = sns.catplot(
        data=group, kind="bar",
        x="Week No.", y="AVG SWE (mm)", hue="Site",
        ci="sd", palette="colorblind", alpha=.6, height=6
    )
    g.despine(left=True)
    g.set_axis_labels("Week Number", "SWE (mm)")
    # g.set_title(f"{name} Sites")
    g.legend.set_title(f"{name} Sites")
    plt.tight_layout()
    plt.savefig(f"/Users/meganmason491/Downloads/{''.join(name.split())}_2020_intervalboard.png")


# Nested barplot by weeks
grouped_df = df.groupby('Week No.')

for week, group in grouped_df:
    print(f"Processing: {week}")
    print('group is:', group)

    g = sns.catplot(
        data=group, kind="bar",
        x="Site", y="AVG SWE (mm)", hue="Location",
        ci="sd", palette="Set1", alpha=.9, height=6, dodge=False
    )
    g.despine(left=True)
    g.set_axis_labels("Date", "SWE (mm)")
    g.set_xticklabels(rotation=45, ha='right')
    # g.set_title(f"{name} Sites")
    g.legend.set_title(f"Week {week}")
    plt.tight_layout(rect=[0.05, 0.05, .75, 0.9])
    plt.title(f"Week {week}")
    # plt.savefig(f"/Users/meganmason491/Downloads/{''.join(name.split())}_2020_intervalboard.png")
    plt.savefig(f"/Users/meganmason491/Downloads/week{week}_2020_intervalboard.png")


    # round 2
    # Plot each year's time series in its own facet
g = sns.relplot(
    data=df,
    x="Week No.", y="AVG HN (cm)", col="Location", hue="Site",
    kind="scatter", palette="crest", linewidth=4, zorder=5,
    col_wrap=3, height=2, aspect=1.5, legend=False,
)

# # Iterate over each subplot to customize further
# for year, ax in g.axes_dict.items():
#
#     # Add the title as an annotation within the plot
#     ax.text(.8, .85, year, transform=ax.transAxes, fontweight="bold")
#
#     # Plot every year's time series in the background
#     sns.lineplot(
#         data=flights, x="month", y="passengers", units="year",
#         estimator=None, color=".7", linewidth=1, ax=ax,
#     )

# Reduce the frequency of the x axis ticks
# ax.set_xticks(ax.get_xticks()[::2])

# Tweak the supporting aspects of the plot
g.set_titles("")
g.set_axis_labels("", "Height of New Snow (cm)")
g.tight_layout()
plt.show()
