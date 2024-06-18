import json
import re
import pandas as pd

# with open('daily.json', 'r') as file:
#     data = file.read()

# corrected_data = re.sub(r'\\/', '/', data)

# json_data = json.loads(corrected_data)

with open("data2.json","r") as f:
    json_data = json.load(f)

data = json_data.get("Data").get("Data")
items = data.get("Items")

total_cars = sum([item.get("Quantity") for item  in items])

# print(total_cars)

# print(len(items))

# dff = pd.DataFrame(items)

# dff.to_excel("daily2.xlsx",index=False)

keys_mapping ={
    "4595":"Ultimate Wash Club",
    "4596":"Deluxe Wash Club",
    "4594":"V.I.P. Wash Club",
    "4":"Dash Wash",
    "401":"V.I.P. Wash Club" ,#App Wash Club Sign Ups
    "402":"Ultimate Wash Club", #WashClubReactivation
    "403":"Deluxe Wash Club", #Kiosk Sign Ups
    "1822":"V.I.P. Wash Club" #Kiosk Sign Ups 
}

ids_count = {}

item_ids = set()


vi_premier_wash_purchases = 0
delux_wash_wash_purchases =0
ultimate_wash_purchases = 0
dash_wash_purchases =0

total_wash_purchases = 0
for item in items :
    
    item_id = item.get("ItemId")
    flag = item.get("Flag")
    item_ids.add(item_id)
    ids_count[str(item_id)]=ids_count.get(str(item_id),0)+1

    if item_id==1 and not flag:
        vi_premier_wash_purchases+=1
    
    if item_id==2 and not flag:
        ultimate_wash_purchases+=1
    if item_id ==3 and not flag:
        delux_wash_wash_purchases+=1

    if item_id==4 and not flag:
        dash_wash_purchases+=1

total_wash_purchases=sum([vi_premier_wash_purchases,delux_wash_wash_purchases,ultimate_wash_purchases,dash_wash_purchases])

print(item_ids)
print("len:",len(item_ids))

print(ids_count)
print("vi premier wash purchases :",vi_premier_wash_purchases)
print("Delux wash purchaes :",delux_wash_wash_purchases)
print("Ultimate  wash pruchases :",ultimate_wash_purchases)
print("Dash wash  purchases :",dash_wash_purchases)
print("Total wash purchases :",total_wash_purchases)