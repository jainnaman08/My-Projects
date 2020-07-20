import os
from win32com import client
root_path = input("Please input the root path: ")
for folder , sub_folders , files in os.walk(root_path):
    
    print("Currently looking at folder: "+ folder)
    print('\n')
    #print("THE SUBFOLDERS ARE: ")
    #for sub_fold in sub_folders:
        #print("\t Subfolder: "+sub_fold )
    
    #print('\n')
    
    print("Converted files: ")
    for f in files:
        try:
        	excel_path = folder+"\\" +f
        	pdf_path = folder+"\\"+f.split('.')[0]
        	excel = client.DispatchEx("Excel.Application")
        	excel.Visible = 0 #keeps the excel sheet closed 0 or False
        	#excel.DisplayAlerts = False #"Do you want to over write it?" Will not Pop up
        	wb = excel.Workbooks.Open(excel_path)
        	#ws = wb.Worksheets[0] #selecting the worksheet
        	#ws.PageSetup.Orientation = 2 #change orientation to landscape
          
        	#wb.ActiveSheet.ExportAsFixedFormat(0,pdf_path) #alternative way to save as pdf
        	wb.SaveAs(pdf_path, FileFormat = 57) 
        except Exception as e:
            print("Failed to convert")
            print(str(e))
        finally:
            wb.Close()
            excel.Quit()
        print(f)
    print('\n')
