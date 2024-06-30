import pandas as pd
from datetime import datetime, date

# Intake
df1 = pd.read_csv('D:/pythonProject/trans.CSV')
pmt = (df1[df1['Gross'] > 0][['Name', 'Gross']]).rename(columns={'Gross': 'Payment'})
pmt['Payment'] = -pmt['Payment']
pmt = pmt.groupby('Name', as_index=False)[['Payment']].sum()


info = pd.read_excel('D:/pythonProject/balances.xlsx')
info = info[info['Name'].notna()]


# Update
rawlist = info.merge(pmt, how='outer')
print('rawlist\n', rawlist)
rawlist['Copay_owed_period'] = rawlist['Copay'] * rawlist['Sessions_period']
rawlist['Balance'] += rawlist['Copay_owed_period'] + rawlist['Payment']
print(rawlist)

# Present
final_bals = rawlist.drop(['Copay_owed_period'], axis=1)
final_bals['Sessions_period'] = [0 for x in range(0, len(final_bals.index))]
#final_bals['Balance'] = [0 for x in range(0, len(final_bals.index)) if x<0]

print(final_bals)

with pd.ExcelWriter('balances.xlsx',
                    if_sheet_exists='new',
                    mode='a') as writer:
    final_bals.set_index('Name').to_excel(writer, sheet_name=f'Balances as of {date.today().strftime('%m-%d-%Y')} ')
