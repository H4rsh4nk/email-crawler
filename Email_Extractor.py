from File_Write import File_Write
from File_Read import File_Read
from Core import Core
from openpyxl import load_workbook





if __name__ == "__main__":
    _File_Write = File_Write()
    _File_Read = File_Read()
    _Core = Core()

    # print(" ___                       ___     ___  __        __  ___  __   __")
    # print("|__   |\/|  /\  | |    __ |__  \_/  |  |__)  /\  /  `  |  /  \ |__)")
    # print("|___  |  | /~~\ | |___    |___ / \  |  |  \ /~~\ \__,  |  \__/ |  \ \n")
    
    # try:
    wb = load_workbook("Book2.xlsx")
    sh1 = wb['Sheet1']
    row = sh1.max_row
    column = sh1.max_column
    for i in range(1 , row+1):
        if(sh1.cell(i,4).value == ""):
            continue
        elif( _Core._Core__Emails != ""):
            _Core.URL(sh1.cell(i,4).value)
            sh1.cell(i,5).value = _Core._Core__Emails
            print(str((i//row)*100) + "% " + "Entry : " + str(i) + "  |" + _Core._Core__Emails)
    # Write text to file
    wb.save("Output.xlsx")
    # _File_Write.File_W(set(_Core._Core__Emails))
# except Exception:
#     print("error")
