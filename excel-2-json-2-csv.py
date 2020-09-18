# import excel2json
# excel2json.convert_from_file('programming-2.xlsx')
year = ['2014','2015','2016','2017','2018','2019','2020', ]

import json
import csv

count = 2014
itemlist = []
finalcounter = 1
result = {}

#fetch itemlist
with open('Items.json') as item:
    data = json.load(item)
    for i in data:
        itemlist.append(i["Itemname"])

        
#List of all items 
print(itemlist)

for i in year:
    print("##############", i)
    filename = i+'.json'
    with open(filename) as f:
        data = json.load(f)
        for item in itemlist:
            counter_in = 0
            counter_out = 0
            temp = {}
            for entry in data:
                if(item == entry['item_in']):
                    counter_in = counter_in + entry['qty_in']
                if(item == entry['item_out']):
                    counter_out = counter_out + entry['qty_out']
            result[finalcounter]={'year':i, 'item_code':item, 'item_in':counter_in, 'item_out':counter_out}
            finalcounter = finalcounter + 1
            
#print(result)

# for val in result:
#     print(val, result[val]['item_in'])


# write to CSV


with open('finalresult.csv', 'w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(["SN", "item_code", "year", "item_in", "item_out"])
    # writer.writerow([1, "Linus Torvalds", "Linux Kernel"])
    # writer.writerow([2, "Tim Berners-Lee", "World Wide Web"])
    # writer.writerow([3, "Guido van Rossum", "Python Programming"])

    for val in result:
        writer.writerow([val, result[val]['item_code'], result[val]['year'], result[val]['item_in'], result[val]['item_out']])
        


