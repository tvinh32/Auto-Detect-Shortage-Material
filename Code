from google.colab import drive
drive.mount('/content/drive')
# Import library
import glob
import pandas as pd
import os
import shutil
!pwd
%cd /content/drive/My Drive/Colab Notebooks

df_demand = pd.read_excel('Use case/Data.xlsx', skiprows=1, usecols=['Job start date','Job open Qty',"Job Item",'On Hand Qty','On Hold Qty'],sheet_name ='FNDWRR' ).fillna(0)
df_demand.head()

df_supply = pd.read_excel('Use case/Data.xlsx', skiprows=6,sheet_name ='supply')
df_supply.head()

# List columns of Job item, Onhand stock v& On Hold Qty
list_stock_cols = ['Job Item','On Hand Qty','On Hold Qty']

# Remove duplicates
df_on_hand_stock = df_demand.drop_duplicates(subset=list_stock_cols)[list_stock_cols]
df_on_hand_stock['On Hand Total']= df_on_hand_stock['On Hand Qty']+ df_on_hand_stock['On Hold Qty']
df_on_hand_stock.drop(columns=['On Hand Qty','On Hold Qty'],inplace = True)
df_on_hand_stock

#Rename column Part_no in sheet supply 
df_supply.rename(columns={"Part_No     ": "Job Item"}, inplace= True)

#Get on-transit qty
df_supply_melted = df_supply.melt(id_vars=['Job Item'],var_name='supply_week',value_name='Planning OH qty')
df_supply_melted.head()

#Insert 1 time column into "on hand" sheet with value 2000 (regard "on hand" was available a long time ago) 
df_on_hand_stock.insert(1,'supply_week', pd.to_datetime('2000-01-01'))
df_on_hand_stock.columns = df_supply_melted.columns
df_on_hand_stock.head(2)

# Merge stock-on-hand and on-transit into 1 for supply sheet 
df_supply_all = pd.concat([df_on_hand_stock,df_supply_melted])
df_supply_all.head()

# Groupby by Job item and calculate cumsum Supply Qty
df_supply_all['cumsum_supply'] = df_supply_all.groupby('Job Item').cumsum()
df_supply_all.head()

# Remove any Job Item has demand = 0
df_real_demand = df_demand[df_demand['Job open Qty'] != 0].drop(columns=['On Hand Qty','On Hold Qty'])

# Groupby by Job item and calculate cumsum Demand Qty
df_real_demand['cumsum_demand'] = df_real_demand.groupby('Job Item').cumsum()
df_real_demand

# Merge demand and supply sheet
df_demand_supply = df_real_demand.merge(df_supply_all,how='right', on='Job Item')
df_demand_supply.head()

df_supply_all[df_supply_all['Job Item'] == 103323003]

df_demand_supply[df_demand_supply['Job Item'] == 103323003]

# Remove rows that had insufficent supply for demand
df_enough = df_demand_supply[df_demand_supply['cumsum_demand'] <= df_demand_supply['cumsum_supply']]

# Get only the first row of sufficent supply 
df_first_supply = df_enough.groupby(['Job Item','cumsum_demand']).first()
df_first_supply.head()

# Retain necessary rows
df_first_supply = df_first_supply.reset_index()[['Job Item','Job open Qty','Job start date','cumsum_demand','cumsum_supply','supply_week','Planning OH qty']]
df_first_supply.head()

# Label rows having insufficent supply
df_final = df_real_demand.merge(df_first_supply,how='left')

# Format 'Job start date' into datetime format
df_final['Job start date'] = pd.to_datetime(df_final['Job start date'])#.dt.strftime('%d-%m-%Y')

# Calculate late days
df_final['Late (day>0)']=(df_final['supply_week']-df_final['Job start date']).dt.days

# Note for insufficient supply or late supply
df_final.loc[df_final['Late (day>0)'] > 0,['Note']]='Tre hang'
df_final.loc[df_final['cumsum_supply'].isnull(),['Note']]='Khong co hang'

#Format 'Job start date' and 'supply_week' into day-montnh-year
df_final['supply_week']= df_final['supply_week'].dt.strftime('%d-%m-%Y')
df_final['Job start date']= df_final['Job start date'].dt.strftime('%d-%m-%Y')
df_final.head()

# Format color 
def highlight_rows(row):
    val = row.loc['Note']
    if val == 'Tre hang':
       color = '#FFB3BA'
    elif val == 'Khong co hang':
       color =  '#BAFFC9'
    else: color =  None
    return ['background-color: {}'.format(color) for r in row]

df_format = df_final.style.apply(highlight_rows,axis=1)

# Xuất file excel
df_format.to_excel('Use case/Output data.xlsx')
