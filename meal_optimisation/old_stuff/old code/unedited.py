
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

del Meals_per_day

# ----------------------------------------------------------------------
# Creating classic/asian orders from MPD
POrds=pd.DataFrame(np.array([[0,0],[0,0],[0,0],[0,0],[0,0],[0,0],[0,0],[0,0],[0,0],[0,0]]),columns=['Classic','Asian'])

for i in range(len(df['Order Number'])):
    if df.at[i,'Item Name']=='Subscription Plan' and df.at[i,'MPD']==1:
        for s in range(0,5):
            POrds.at[s,'Classic']=POrds.at[s,'Classic']+1

    if df.at[i,'Item Name']=='Subscription Plan' and df.at[i,'MPD']==2:
        for s in range(0,10):
            POrds.at[s,'Classic']=POrds.at[s,'Classic']+1

    if df.at[i,'Item Name']=='Asian Plan' and df.at[i,'MPD']==1:
        for s in range(0,5):
            POrds.at[s,'Asian']=POrds.at[s,'Asian']+1

    if df.at[i,'Item Name']=='Asian Plan' and df.at[i,'MPD']==2:
        for s in range(0,10):
            POrds.at[s,'Asian']=POrds.at[s,'Asian']+1

    if df.at[i,'Item Name']=='Fusion Meal Plan' and df.at[i,'MPD']==1:
        if (weeknum % 2) == 0:
            for s in range(0,2):
                POrds.at[s,'Asian']=POrds.at[s,'Asian']+1
            for s in range(0,3):
                POrds.at[s,'Classic']=POrds.at[s,'Classic']+1
        else:
            for s in range(0,3):
                POrds.at[s,'Asian']=POrds.at[s,'Asian']+1
            for s in range(0,2):
                POrds.at[s,'Classic']=POrds.at[s,'Classic']+1            

    if df.at[i,'Item Name']=='Fusion Meal Plan' and df.at[i,'MPD']==2:
        for s in range(0,5):
            POrds.at[s,'Asian']=POrds.at[s,'Asian']+1
            POrds.at[s,'Classic']=POrds.at[s,'Classic']+1


# ----------------------------------------------------------------------
# Loading whole menu, Types: V=Veg, R=Red meant, C=Chicken
OMR=oxl.load_workbook(filename="Meal Orders-Meals-Recipes.xlsx")
sheet=OMR['Menus']

data=sheet.values
data2 = list(data)[1:]
MenuImport = pd.DataFrame(data2, columns=['Menu','Dish','Type'])


# ----------------------------------------------------------------------
# Loading previous 2 weeks orders
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




print(POrds)










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




























