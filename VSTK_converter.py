from openpyxl import load_workbook
import subprocess
from time import sleep

wb = load_workbook("portrait.xlsx", data_only=True)
sh = wb["Sheet"]

#ZEROPOINT = sh.cell(1, 1)
#print(sh["n31"].value)
#print(sh.cell(17, 2).value)

def cell_no():
    for i in range(1, 100):
        if sh.cell(i, 1).value == "No.":
            return [i, 1]
POS_Cell_No = cell_no()
print("No. is at position " + str(POS_Cell_No))

def cell_part():
    for i in range(1, 100):
        if sh.cell(POS_Cell_No[0], i).value == "Part#\n":
            return [POS_Cell_No[0], i]
POS_Cell_Part = cell_part()
print("Part# is at position " + str(POS_Cell_Part))

def cell_first_characteristic():
    for i in range(1,100):
        finder = str(sh.cell(POS_Cell_No[0]+1, i).value)
        if any(c.isalpha() for c in finder) and finder != "None":
            return [POS_Cell_No[0]+1, i]
POS_Cell_First_Characteristic = cell_first_characteristic()
print("First Characteristic is at position " + str(POS_Cell_First_Characteristic))

#TODO  read provided string and find right result to output.
#      example output : (19.9 , 20.1) or (I.O)
def characteristic_decoder(string):
    pass


#TODO  lowers spread of provided tolerances to be more tight/realistic.
def tolerance_modifier(value1, value2):
    pass


#TODO  Enters results into cells
def result_inputer():
    pass





#wb.save('test.xlsx')
#sleep(0.5)
#subprocess.Popen(["test.xlsx"],shell=True)
