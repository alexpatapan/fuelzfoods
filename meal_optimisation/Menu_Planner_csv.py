
import pandas as pd
from datetime import date
import numpy as np
import openpyxl as oxl
from random import seed
import time
import random

random.seed();


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
    if entry <= 0:
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
# Creating classic/asian Order table from MPD

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
        
POrds=pd.DataFrame(np.zeros((10, 4)), columns=['Classic', 'numC', 'Asian', 'numA'], dtype=object)

for i in range(0, len(df['Order Number'])):

    # Normal meal plans
    createPOrds(POrds, "Subscription Plan", "numC", i)
    createPOrds(POrds, "Asian Plan", "numA", i)

    # Fusion Meals
    if df.at[i,'Item Name'] == 'Fusion Meal Plan' and df.at[i,'MPD'] == 1:
        #remember to change weeknum back
        weeknum = 1
        if (weeknum % 2) == 0:
            for s in range(0,2):
                POrds.at[s,'numA'] += 1
            for s in range(0,3):
                POrds.at[s,'numC'] += 1
        else:
            for s in range(0,3):
                POrds.at[s,'numA'] += 1
            for s in range(0,2):
                POrds.at[s,'numC'] += 1          

    if df.at[i,'Item Name']=='Fusion Meal Plan' and df.at[i,'MPD']==2:
        for s in range(0, 5):
            POrds.at[s,'numA'] += 1
            POrds.at[s,'numC'] += 1

# ----------------------------------------------------------------------
# Loading previous 2 weeks orders into dataframes (HistoricalWeek1 and HistoricalWeek2)

HistoricalWeek1=pd.DataFrame(np.zeros((10, 4)), columns=['Classic', 'numC', 'Asian', 'numA'], dtype=object)
HistoricalWeek2=pd.DataFrame(np.zeros((10, 4)), columns=['Classic', 'numC', 'Asian', 'numA'], dtype=object)

OMR=oxl.load_workbook(filename="Past 2 Weeks.xlsx")
sheet=OMR['Order History']

def populateHistoricalOrders(HistoricalWeek, weeknum, excelColumn1, excelColumn2, orderColumn1, orderColumn2):
    inc = 0
    for i in range(weeknum*11+2, (weeknum+1)*11+1):
        HistoricalWeek.at[inc, orderColumn1]=sheet[excelColumn1+str(i)].value
        HistoricalWeek.at[inc, orderColumn2]=sheet[excelColumn2+str(i)].value
        inc += 1
weeknum = 0
populateHistoricalOrders(HistoricalWeek1, weeknum, 'D', 'E', 'Classic', 'numC')
populateHistoricalOrders(HistoricalWeek1, weeknum, 'G', 'H', 'Asian', 'numA')

populateHistoricalOrders(HistoricalWeek2, weeknum+1, 'D', 'E', 'Classic', 'numC')
populateHistoricalOrders(HistoricalWeek2, weeknum+1, 'G', 'H', 'Asian', 'numA')

# ----------------------------------------------------------------------
# Make list of custom order dishes for this week (cusomOrders)
customOrders = []
for i, j in enumerate(df['Plan?']):
    if j == 0:
        customOrders.append(df.at[i, 'Item Name'])

#----------------------------------------------------------------------
# Loading whole menu into a dictionary (mealMenu), Types: V=Veg, R=Red meant, C=Chicken
OMR=oxl.load_workbook(filename="Past 2 Weeks.xlsx")
sheet=OMR['Menus']

data=sheet.values
data2 = list(data)[1:]
MenuImport = pd.DataFrame(data2, columns=["ID", 'Menu','Dish','Type'])

mealMenu = {}
for i in (range(43)):
    mealMenu[MenuImport.at[i, 'Dish']] = MenuImport.at[i, 'Menu']

# ----------------------------------------------------------------------
#Find number of meals needed to be made (numClassics, numAsian)
numClassics = -1
numAsian = -1
for i in (range(10)):
    if (POrds.at[i, 'numA'] == 0 and numClassics == -1):
        numClassics = i
    if (POrds.at[i, 'numC'] == 0 and numAsian == -1):
        numAsian = i

# ----------------------------------------------------------------------
#Calculate most optimal dish!
completedClassics = 0 #num optimal dishes we have found
completedAsian = 0

#Possibly put custom order dishes which haven't been made last week into the meal packs (50% chance)
for meal in customOrders:
    if (random.random() < 0.5):
        continue
    if (mealMenu[meal] == 'A'):
        if not (meal in iter(HistoricalWeek1['Asian'])):
            POrds.at[completedAsian, 'Asian'] = meal
            completedAsian += 1
    else:
        if not (meal in iter(HistoricalWeek1['Classic'])):
            POrds.at[completedClassics, 'Classic'] = meal
            completedClassics += 1

#At this point we have (possibly) added custom orders -> Pick random meal which hasn't been made in last week
def inLastWeek(meal, menuType, num):
    for i in (range(num)):
        if (meal == HistoricalWeek1.at[i, menuType]):
            return True
    return False

def inPOrds(meal, menuType):
    for i in (range(10)):
        if (meal == POrds.at[i, menuType]):
            return True
    return False
            
def pickRandomMeal(menuName, numMeals, completedMeals):
    while (completedMeals <= 4):
        meal = MenuImport.at[random.randint(1, 42), 'Dish']
        #check if meal wasnt in first 5 of last week
        if (not inLastWeek(meal, menuName, 5) and (not inPOrds(meal, menuName))):
            POrds.at[completedMeals, menuName] = meal
            completedMeals += 1

    while (completedMeals > 4 and completedMeals <= 9):
        numMealsInLastWeek = 0
        meal = MenuImport.at[random.randint(1, 42), 'Dish']
        #we are allowed upto 2 meals which have been made in past week for T meals/week
        if (not inPOrds(meal, menuName)):
            if (inLastWeek(meal, menuName, 10)):
                numMealsInLastWeek += 1

            if (numMealsInLastWeek > 2):
                continue
            POrds.at[completedMeals, menuName] = meal
            completedMeals += 1
        
pickRandomMeal('Asian', numAsian, completedAsian)
pickRandomMeal('Classic', numClassics, completedClassics)
print(POrds)
print(customOrders)

#Need to add number of custom orders to POrds
#Then save POrds to an excel sheet


"""
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
"""



























