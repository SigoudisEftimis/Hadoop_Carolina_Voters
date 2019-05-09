import csv
import openpyxl
import pandas as pd
import re 
 

#Constants indicating the colums to be deleted

class Constants :
    ONE  = 1  
    NINE = 9 
    SEVEN = 7
    FIFTY_THREE = 53 
    EIGHT = 8 
    FOUR = 4 



# set the paths from the files 

print("Starting........")

input_file = 'ncvoter_Statewide.txt'
output_file_xlsx = 'outputfile.xlsx'

wb = openpyxl.Workbook()
ws = wb.worksheets[0]


#open the file with the unprocessed data

with open(input_file, 'r' , encoding="utf8", errors='ignore') as data:
    reader = csv.reader(data, delimiter='\t')
    for row in reader:
        ws.append(row)

print("Just Finish opening the data file")

# delete the unecessary colums
        
ws.delete_cols(Constants.ONE,Constants.NINE)
ws.delete_cols(Constants.SEVEN,Constants.FIFTY_THREE)
ws.delete_cols(Constants.EIGHT,Constants.NINE)
ws.delete_cols(Constants.FOUR)
ws.delete_cols(Constants.FOUR)
wb.save(output_file_xlsx)


print("Colums removal just complete")

# open xlsx file and convert xlsx file to txt , save to text.txt

xlsx = pd.read_excel('outputfile.xlsx', sheet_name=0, index=False)
with open('text.txt','w') as outfile: xlsx.to_string(outfile)
with open('text.txt', 'r') as f: lines = f.readlines()
    
print("Complete convertion xlsx to txt")

# remove spaces

regex_mspace = re.compile(' +')
regex_space = re.compile(' ')
lines = [re.sub(regex_mspace,' ',line) for line in lines]
print("Space Removal Phase 1 complete")
lines = [re.sub(regex_space,',',line) for line in lines]
print("Space Removal Phase 2 complete")

# finally, write lines in the file

with open('text.txt', 'w') as f:
    f.writelines(lines)










