# Importing required library 
import pygsheets
import pandas as pd

startingDataRow = 6
sevenDaySalesColumn = "BF"
fourteenDaySalesColumn = "BG"
thirtyDaySalesColumn = "BH"
ninetyDaySalesColumn = "BI"
marketRegion = "US"

client = pygsheets.authorize(service_account_file="ecomspot-414511-51041f900256.json") 

# opens a spreadsheet by its name/title 
spreadsht = client.open('12-Feb-Mckillans Car Care') 
  
# opens a worksheet by its name/title 
worksht = spreadsht.worksheet("title", "MAIN - reorder suggestion") 
  
last_row_index = len(worksht.get_col(7, include_tailing_empty=False))
  
marketplace = worksht.get_values('G'+str(startingDataRow),"G"+str(last_row_index))
countUS = -1
countCA = 0
for region in marketplace:
    if region[0] == "US": 
        countUS = countUS + 1
    elif region[0] == "CA":
        countCA = countCA + 1


SevenDaySales = []
if marketRegion == "US":
    SevenDaySales = worksht.get_values(sevenDaySalesColumn + str(startingDataRow),sevenDaySalesColumn + str(startingDataRow+countUS))
elif marketRegion == "CA":
    SevenDaySales = worksht.get_values(sevenDaySalesColumn + str(startingDataRow + countUS + 1),sevenDaySalesColumn + str(startingDataRow + countUS + 1 +countCA))
    
sevenDaySum = 0
for unit in SevenDaySales: 
    sevenDaySum = int(unit[0]) + sevenDaySum

FourteenDaySales = []
if marketRegion == "US":
    FourteenDaySales = worksht.get_values(fourteenDaySalesColumn + str(startingDataRow),fourteenDaySalesColumn + str(startingDataRow+countUS))
elif marketRegion == "CA":
    FourteenDaySales = worksht.get_values(fourteenDaySalesColumn + str(startingDataRow + countUS + 1),fourteenDaySalesColumn + str(startingDataRow + countUS + 1 +countCA))
fourteenDaySum = 0
for unit in FourteenDaySales: 
    fourteenDaySum = int(unit[0]) + fourteenDaySum

thirtyDaySales = []
if marketRegion == "US":
    thirtyDaySales = worksht.get_values(thirtyDaySalesColumn + str(startingDataRow),thirtyDaySalesColumn + str(startingDataRow+countUS))
elif marketRegion == "CA":
    thirtyDaySales = worksht.get_values(thirtyDaySalesColumn + str(startingDataRow + countUS + 1),thirtyDaySalesColumn + str(startingDataRow + countUS + 1 +countCA))
    
thirtyDaySum = 0
for unit in thirtyDaySales: 
    thirtyDaySum = int(unit[0]) + thirtyDaySum

ninetyDaySales = []
if marketRegion == "US":
    ninetyDaySales = worksht.get_values(ninetyDaySalesColumn + str(startingDataRow),ninetyDaySalesColumn + str(startingDataRow+countUS))
elif marketRegion == "CA":
    ninetyDaySales = worksht.get_values(ninetyDaySalesColumn + str(startingDataRow + countUS + 1),ninetyDaySalesColumn + str(startingDataRow + countUS + 1 +countCA))
    # print(countCA)
ninetyDaySum = 0
for unit in ninetyDaySales: 
    ninetyDaySum = int(unit[0]) + ninetyDaySum

# Fetching files from desktop:

UnitsOrderedSevenDays = pd.read_csv('7s.csv')["Units Ordered"].to_numpy()
UnitsOrderedFourteenDays = pd.read_csv('14s.csv')["Units Ordered"].to_numpy()
UnitsOrderedThirtyDays = pd.read_csv('30s.csv')["Units Ordered"].to_numpy()
UnitsOrderedNinetyDays = pd.read_csv('90s.csv')["Units Ordered"].to_numpy()

verifySevenDaySum = 0
verifyFourteenDaySum = 0
verifyThirtyDaySum = 0
verifyNinetyDaySum = 0

for units in UnitsOrderedSevenDays:
    verifySevenDaySum = int(str(units).replace(",","")) + int(verifySevenDaySum)
if sevenDaySum == verifySevenDaySum:
    print("Verified successfully: 7 Day Sales.")
else:
    print("Sums doesn't match properly 7 day sales, Main Dashboard:"+str(sevenDaySum)+" , Individual file:" + str(verifySevenDaySum) + ".")


for units in UnitsOrderedFourteenDays:
    verifyFourteenDaySum = int(str(units).replace(",","")) + verifyFourteenDaySum
if fourteenDaySum == verifyFourteenDaySum:
    print("Verified successfully: 14 Day Sales.")
else:
    print("Sums doesn't match properly for 14 days sales, Main Dashboard:"+str(fourteenDaySum)+" , Individual file:" + str(verifyFourteenDaySum) + ".")


for units in UnitsOrderedThirtyDays:
    verifyThirtyDaySum = int(str(units).replace(",","")) + verifyThirtyDaySum
if thirtyDaySum == verifyThirtyDaySum:
    print("Verified successfully: 30 Day Sales.")
else:
    print("Sums doesn't match properly for 30 day sales, Main Dashboard:"+str(thirtyDaySum)+" , Individual file:" + str(verifyThirtyDaySum) + ".")


for units in UnitsOrderedNinetyDays:
    verifyNinetyDaySum = int(str(units).replace(",","")) + verifyNinetyDaySum
if ninetyDaySum == verifyNinetyDaySum:
    print("Verified successfully: 90 Day Sales.")
else:
    print("Sums doesn't match properly for 90 day sales, Main Dashboard:"+str(ninetyDaySum)+" , Individual file:" + str(verifyNinetyDaySum) + ".")