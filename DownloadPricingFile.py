import xlrd
import os
import requests

def download(url: str, dest_folder: str, counter:int):
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)  # create folder if it does not exist


    filename = url.split('/')[-1].replace(" ", "_")  # be careful with file names
    file_path = os.path.join(dest_folder,  str(counter)+"_"+filename)

    r = requests.get(url, stream=True)
    if r.ok:
        print("saving to", os.path.abspath(file_path))
        with open(file_path, 'wb') as f:
            for chunk in r.iter_content(chunk_size=1024 * 8):
                if chunk:
                    f.write(chunk)
                    f.flush()
                    os.fsync(f.fileno())
    else:  # HTTP status code 4XX/5XX
        print("Download failed: status code {}\n{}".format(r.status_code, r.text))


#download("http://website.com/Motivation-Letter.docx", dest_folder="mydir")

loc = ("C:\\Users\\rkumar\\sheet.xls")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# For row 0 and column 0
sheet.cell_value(1, 8)
destinationfolder = 'C:\\Users\\rkumar\\auditprc\\'
for i in range(sheet.nrows):
    if i > 0:
        print(sheet.cell_value(i,8))
        download(sheet.cell_value(i,8),destinationfolder,i)
