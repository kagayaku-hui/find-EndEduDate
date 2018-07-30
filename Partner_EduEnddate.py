# Purpose: Find out the end date of audit partner's education
# Author: kagayaku-hui


import operator
import pandas as pd
from pandas import DataFrame
from helper import partner, date
# import excel spreadsheet "Partner_EduEndDate"
path = r'Path\Parter_EduEndDate.xlsx'


df = pd.read_excel(path)
print(df.head())
print(df['major'][1])
print(df.shape)
num_rows,num_cols = df.shape
partner_dict = {}

# save partner into dictionary--partner_dict, where key is partner id and value is a list of dates. date is a pair of (start date, end date)
for i in range(num_rows):
    id = df['Engagement Partner ID'][i]
    if partner_dict.__contains__(id):
        current_partner = partner_dict[id]
    else:
        current_partner = partner(id)
        current_partner.dates = []
    startDate = df['startDate'][i]
    endDate = df['endDate'][i]
    current_date = date(startDate, endDate)
    current_partner.dates.append(current_date)
    partner_dict[id] = current_partner

# go through each partner, and save the partner id and "start working date" into dictionary--partner_work_dict
partner_work_dict = {}
for id in partner_dict.keys():
    current_partner = partner_dict.get(id)
    dates = current_partner.dates
    dates.sort(key=operator.attrgetter('startDate'))
    last_degree_end_year = dates.__getitem__(0).endDate
    workYear = last_degree_end_year
    for i in range(dates.__len__()):
        current = dates[i]
        # if the current degree start date is greater than the last degree end date, the work year is the last degree end date
        if i > 0 and current.startDate > last_degree_end_year:
            workYear = last_degree_end_year
            break
        last_degree_end_year = current.endDate
        workYear = last_degree_end_year
    partner_work_dict[id] = workYear
df = DataFrame({'partner id': list(partner_work_dict.keys()), 'start work date': list(partner_work_dict.values())})
print(partner_work_dict)

#save results to the following path
destination = r'Path'
df.to_excel(destination, sheet_name='sheet1', index=False)





