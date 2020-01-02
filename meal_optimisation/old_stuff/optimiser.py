
import pandas as pd
from datetime import date
import numpy as np
import openpyxl as oxl

# ----------------------------------------------------------------------
# Creating data frame from csv of todays date 

today=date.today()
OrderDate=today.strftime("%Y-%m-%d")
CSVOrderFile="FuelzOrders-2019-11-22.csv"
df=pd.read_csv(CSVOrderFile)
weeknum=date.isocalendar(today)[1]

# ----------------------------------------------------------------------
# Searching orders for meal plans
sub ='Plan'
df['Plan?']=df["Item Name"].str.find(sub)  
c=0      
for entry in df['Plan?']:
    if entry <=0:
        df.at[c,'Plan?']=0
    else:
        df.at[c,'Plan?']=1
    c=c+1

# Text searching for meals per day from plans (MPD)      
Meals_per_day = df["Product Variation"].str.extract(pat = '(: .)')
for i in range(len(Meals_per_day)):
    if str(df.at[i,'Product Variation'])[0]!='M':
        Meals_per_day.iat[i,0]=int(0)
    else:
        Meals_per_day.iat[i,0]=int(str(Meals_per_day.iat[i,0])[2])
df['MPD']=Meals_per_day

# ----------------------------------------------------------------------
# Creating classic/asian orders from MPD

"""
Method to search table for plan type and add the number of meals requested
to the appropriate column
    POrds - table
    planType - The subscription plan (E.g. Asian Plan / Subscription)
    columnName - The POrds column to add the meals to
"""
def createPOrds(POrds, planType, columnName, i) :
    if df.at[i,'Item Name'] == planType:
        for s in range(0, df.at[i,'MPD'] * 5):
            POrds.at[s,columnName] += 1
        
POrds=pd.DataFrame(np.array([[0,0],[0,0],[0,0],[0,0],[0,0],[0,0],[0,0],[0,0],[0,0],[0,0]]),columns=['Classic','Asian'])

for i in range(0, len(df['Order Number'])):

    # Normal meal plans
    createPOrds(POrds, "Subscription Plan", "Classic", i)
    createPOrds(POrds, "Asian Plan", "Asian", i)

    # Fusion Meals
    if df.at[i,'Item Name'] == 'Fusion Meal Plan' and df.at[i,'MPD'] == 1:
        #remember to change this
        weeknum = 1
        if (weeknum % 2) == 0:
            for s in range(0,2):
                POrds.at[s,'Asian'] += 1
            for s in range(0,3):
                POrds.at[s,'Classic'] += 1
        else:
            for s in range(0,3):
                POrds.at[s,'Asian'] += 1
            for s in range(0,2):
                POrds.at[s,'Classic'] += 1          

    if df.at[i,'Item Name']=='Fusion Meal Plan' and df.at[i,'MPD']==2:
        for s in range(0, 5):
            POrds.at[s,'Asian'] += 1
            POrds.at[s,'Classic'] += 1

print(POrds)

"""
#----------------------------------------------------------------------
# Loading whole menu, Types: V=Veg, R=Red meant, C=Chicken
OMR=oxl.load_workbook(filename="Meal Orders-Meals-Recipes.xlsx")
sheet=OMR['Menus']

data=sheet.values
data2 = list(data)[1:]
MenuImport = pd.DataFrame(data2, columns=['Menu','Dish','Type'])
"""


# ----------------------------------------------------------------------
# Loading previous 2 weeks orders

HistoricalPOrds1=pd.DataFrame(np.array([[0, 0, 0, 0], [0, 0, 0, 0]]),columns=['Classic', 'num', 'Asian', 'num'])
HistoricalPOrds2=pd.DataFrame(np.array([[0, 0, 0, 0], [0, 0, 0, 0]]),columns=['Classic', 'num', 'Asian', 'num'])

print(HistoricalPOrds1)
OMR=oxl.load_workbook(filename="Meal Orders-Meals-Recipes.xlsx")
sheet=OMR['Order History']

# 1 week back
if weeknum<46:
    step1=(6+weeknum)*11+2
else:
    step1=(weeknum-46)*11+2

Classics1=[0,0,0,0,0,0,0,0,0,0]
for i in range(10):
    Classics1[i]=sheet['D'+str(i+step1)].value 
    
# 2 weeks back
if weeknum<46:
    step2=(5+weeknum)*11+2
else:
    step2=(weeknum-47)*11+2


Classics2=[0,0,0,0,0,0,0,0,0,0]
for i in range(10):
    Classics2[i]=sheet['D'+str(i+step2)].value 




# ----------------------------------------------------------------------
# Putting PLAN ONLY orders into excel sheet (in future change this POrds to the true total orders)
            
sheet=OMR['Order History']

if weeknum<46:
    step=(7+weeknum)*11+2
else:
    step=(weeknum-45)*11+2

for i in range(10):
    ECell='E'+str(i+step)
    sheet[ECell]=int(POrds.iat[i,0])
    HCell='H'+str(i+step)
    sheet[HCell]=int(POrds.iat[i,1])

OMR.save("Meal Orders-Meals-Recipes.xlsx")




























