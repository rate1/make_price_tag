import pandas as pd
import xlwings as xw
from pathlib import Path

def write_result():
		PosterID = []
		Name = []
		Code = []
		Unit = []
		Cost = []

		try:
			data = pd.read_excel('./input.xls')
		except:
			print("Input file must be named input.xls")
			data = ''

		cols = len(data)
		print("Cols in file = ", cols)

		for x in range(cols):
			if data.Cost[x] != data.Cost[x]: 
				base_name = data.Name[x].strip()
			else:
				if data.PosterID_m[x] != data.PosterID_m[x]:
					Name.append(data.Name[x])
				else:
					new_name = base_name + ' ' + data.Name[x]
					Name.append(new_name)

				PosterID.append(data.PosterID[x])
				Code.append(data.Code[x])
				Unit.append(data.Unit[x])
				Cost.append(data.Cost[x])


		df = pd.DataFrame({'PosterID':PosterID,
						   'Name':Name,
						   'Code':Code,
						   'Unit':Unit,
						   'Cost':Cost})

		df.to_excel('./result.xls', index = False)

def del_cols_rows():
    vba_book = xw.Book("./macros.xlsm")
    vba_macro = vba_book.macro("Get_All_File_from_Folder")
    try:
        vba_macro()
    except:
        print("Error-xy-error")
    print("Empty rows and cols deleted")

def merge_files():
    min_excel_file_size = 1
    path = Path("./input")
    
    df = pd.concat([pd.read_excel(f)
                    for f in path.glob("*.xls")
                    if f.stat().st_size >= min_excel_file_size],
                    ignore_index = True)    
    
    df.to_excel('./input.xls', index = False)
    print("Files are merged")

def main():
    del_cols_rows()
    merge_files()
    #write_result()

if __name__ == '__main__':
    main()
