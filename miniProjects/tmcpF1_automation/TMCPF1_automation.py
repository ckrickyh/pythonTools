# %%
import pyautogui as gui
import time
from datetime import datetime
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import cv2
import sys
import os
import gspread
from gspread_dataframe import set_with_dataframe
import socket
from config import GOOGLE_SHEET_ID, GOOGLE_SHEET_ID2, sysPassword
import keyboard
import threading

#from PIL import ImageGrab
#import pytesseract
#from paddleocr import PaddleOCR

# %%
print(f'Transforming Nature: The Revolution of HK Urban Forest Management')
#print(f'pyautogui == {gui.__version__}')
#print(f'pandas == {pd.__version__}')
#print(f'numpy == {np.__version__}')
#print(f'opencv-python == {cv2.__version__}')
#print(f'gspread == {gspread.__version__}')
#print(f'google-api-python-client == {google-api-python-client.__version__}')
#print(f'google-auth-httplib2 == {google-auth-httplib2.__version__}')

'''
pip install google-api-python-client google-auth-httplib2
pyautogui == 0.9.54
pandas == 2.3.1
numpy == 2.2.6
opencv-python == 4.10.0
auto-py-to-exe == 2.46.0
gspread == 6.2.1
'''

# %% [markdown]
# # exit program when pressing  'esc' - threading

# %%
stop_thread = False

def check_for_exit():
    global stop_thread
    while True:
        if keyboard.is_pressed('esc'):  # If 'esc' is pressed
            print("Exiting program...")
            stop_thread = True
            sys.exit()

# %%
# paddleOcr
#https://www.paddleocr.ai/latest/quick_start.html#1-paddlepaddle
#%pip install paddlepaddle==3.1.0 -i https://www.paddlepaddle.org.cn/packages/stable/cpu/
#%pip install paddlepaddle-gpu==3.1.0 -i https://www.paddlepaddle.org.cn/packages/stable/cu118/
#%pip uninstall paddlepaddle
#%pip install paddleocr
#%pip uninstall paddleocr

# %% [markdown]
# # pyautogui Detect image

# %%
def resourcePath(relativePath):
    try:
        #pyinstaller temporary folder
        base_path = sys._MEIPASS
    except AttributeError:
            base_path = os.path.abspath(".")
    return os.path.join(base_path, relativePath)

def detectImage(image):
  config_path = resourcePath(image)
  #config_path = image
  imageLocation = None
  maxTime = 60
  startTime = time.time()

  while time.time()-startTime < maxTime:
    print(f"Searching for image {image}...")
    try:
      imageLocation = gui.locateCenterOnScreen(config_path, confidence=0.9)
      if imageLocation:
        print("image found")
        x = imageLocation.x
        y = imageLocation.y
        time.sleep(0.5)
        break
    except:
      time.sleep(1)

  if imageLocation is None:
    print(f"notFound image within {maxTime}s, Stop Program")

    messagebox.showwarning("Warning", "over 60s, stop program!")
    tk.destroy()
    exit()

  return x, y

# %% [markdown]
# # Fundamental

# %%
# oop
class OriginalScreen:
    width = 2736
    height = 1824

    @staticmethod
    def getWidth():
        return OriginalScreen.width

    @staticmethod
    def getHeight():
        return OriginalScreen.height


class UserScreenSize:
    @staticmethod
    def getScreenSize():
        return gui.size()

    @staticmethod
    def getWidth():
        width, _ = UserScreenSize.getScreenSize()
        return width

    @staticmethod
    def getHeight():
        _, height = UserScreenSize.getScreenSize()
        return height

def relativeXY(x, y):
    relativeX = x/OriginalScreen.getWidth() # static method => OriginalScreen.width(), instantace method => originalScreen().width
    relativeY = y/OriginalScreen.getHeight()
    userX = int(relativeX * UserScreenSize.getWidth())
    userY = int(relativeY * UserScreenSize.getHeight())
    #print(f'x:{userX}, y:{userY}')
    return userX, userY

# relative XY
def click(x,y):
    gui.click(relativeXY(x,y))
    time.sleep(0.1)

# %% [markdown]
# # Home

# %%
def clickF1():
    click(1160,870)
    time.sleep(0.5)

def save():
    x, y = detectImage('saveBtn.PNG')
    click(x, y)
    time.sleep(0.5)
    x,y = detectImage('saveSuccessed.PNG')
    click(x, y)
    time.sleep(1)

def close():
    x, y = detectImage('closeBtn.PNG')
    gui.click(x, y)
    #click(2652, 62)
    x, y = detectImage('saveAndExit.PNG')
    click(x, y)
    #gui.click(1038,997)
    time.sleep(5)

# %% [markdown]
# # General Information

# %%
# class General Information

def clickNewF1():
    time.sleep(1)
    click(180, 140)
    x, y = detectImage('addBtn.PNG')
    click(x, y)
    click(200,400)
    x, y = detectImage('form1Title.PNG') #能否進入 General Information (save button)
    print("clickNewF1 Done")

def wF1RefNo(f1RefNo):
    click(447, 263)
    gui.hotkey('ctrl', 'a', 'del')
    gui.write(f1RefNo)
    print("f1RefNo Done")

def wFileRef(fileRef):
    if isinstance(fileRef, str):
        click(1330, 986)
        gui.write(fileRef)
        time.sleep(0.5)
        print("fileRef Done")

#special treatment (yyyy/dd/mm)
def wDateOfInspection(dateOfInspection):
    click(435, 1353)
    gui.write(dateOfInspection, interval=0.02)
    time.sleep(0.5)
    print("dateOfInspection Done")

def wLastInspectionDate(lastInspectionDate):
    if isinstance(lastInspectionDate, str):
        click(1790, 1361)
        time.sleep(0.5)
        click(1252, 1679)
        click(1330, 1363)
        gui.write(lastInspectionDate, interval=0.02)
        time.sleep(0.5)
        print("lastInspectionDate Done")

#(L, 3, 6, 12, 24 , 36, 48, 60, ad)
def wInspectionFreq(inspectionFreq):
    if isinstance(inspectionFreq, str):
        click(2269, 1358)
        gui.write('a')
        if inspectionFreq == '36':
            gui.write('3')
        gui.write(inspectionFreq[0])   #variable (text)
        gui.press('enter')
        time.sleep(0.5)
        print("inspectionFreq Done")

# %% [markdown]
# # Main General Info

# %%
#General Information
def generalInfo(f1RefNo, fileRef, dateOfInspection, lastInspectionDate, lnspectionFreq):
    clickNewF1()
    wF1RefNo(f1RefNo)
    wFileRef(fileRef)
    wDateOfInspection(dateOfInspection)
    wLastInspectionDate(lastInspectionDate)
    wInspectionFreq(lnspectionFreq)

# %% [markdown]
# # Location info

# %%
# Location info

# change to LocationInfo
def wLocationInfo():
    click(223, 47)
    gui.write('l')
    time.sleep(0.5)
    print("locationInfo Done")

def wMasterZoneRef(masterZoneRef):
    click(424, 253)
    gui.write(masterZoneRef)
    gui.press('enter')
    gui.press('enter')
    time.sleep(2)
    print("masterZoneRef Done")

def wSubZoneRef(subZoneRef):
    if isinstance(subZoneRef, str):
        click(1440,262)
        time.sleep(1)
        gui.write(subZoneRef)
        gui.press('enter')
        time.sleep(1)
        print("subZoneRef Done")

def wLampPost(lampPost):
    if isinstance(lampPost, str):
        click(2156, 749)
        gui.write(lampPost)
        time.sleep(0.5)
        print("lp Done")

def wLocationTypes(locationTypes):
    if isinstance(locationTypes, str):
        locLst = locationTypes.split(',')

        locTypeDict = {
        'ROADSIDE': {'x': 29, 'y': 1077},
        'PUBLICPARK': {'x': 895, 'y': 1077},
        'CENTRALDIVIDER': {'x': 2180, 'y': 1077},
        'HOUSING': {'x': 29, 'y': 1185},
        'UNLEASED': {'x': 895, 'y': 1185},
        'PLANTER': {'x': 2180, 'y': 1185},
        'GOVERNMENT': {'x': 29, 'y': 1300},
        'RECREATIONAL': {'x': 895, 'y': 1300},
        'TREEPIT': {'x': 2180, 'y': 1300},
        'SIMAR': {'x': 28, 'y': 1406},
        'OTHERS': {'x': 28, 'y': 1524},
        }

        for loc in locLst:
            if loc in locTypeDict:
                originalX= locTypeDict[loc]['x']
                originalY = locTypeDict[loc]['y']
                xfinal, yfinal = relativeXY(originalX, originalY)
                click(xfinal, yfinal)
        print("locationTypes Done")


def wLocationTypeInfo(locationTypes, masterZoneRef, otherLocRemarks):
    if isinstance(locationTypes, str):
        locLst = locationTypes.split(',')
        for loc in locLst:
            if (loc == 'SIMAR'):
                click(1057,1415)
                gui.write(masterZoneRef)
            if (loc == 'OTHERS') and isinstance(otherLocRemarks, str):
                click(1092,1521)
                gui.write(otherLocRemarks)
        print("locationTypeInfo Done")

# %% [markdown]
# # Main Location Info

# %%
def locationInfo(masterZoneRef, subZoneRef, lampPost, locationTypes, otherLocRemarks):
    wLocationInfo()
    wMasterZoneRef(masterZoneRef)  # e.g., 'Kowloon Bay, Wang Chin Street'
    wSubZoneRef(subZoneRef)          # e.g., '1 Wang Chin St, Kowloon Bay'
    wLampPost(lampPost)              # e.g., lampPost variable
    wLocationTypes(locationTypes)     # e.g., "ROADSIDE,OTHERS"
    wLocationTypeInfo(locationTypes, masterZoneRef, otherLocRemarks)  # e.g., "ROADSIDE,OTHERS", "11sw-cc/678", "hospital"
    save()
    close()

# %% [markdown]
# # Clean df

# %%
def cleanDf(dfForm1):
    # trimTreatment
    dfForm1 = dfForm1.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    # Clean cells that contain only whitespace
    dfForm1 = dfForm1.applymap(lambda x: np.nan if isinstance(x, str) and x.strip() == "" else x)
    # lowercase
    dfForm1['inspectionFreq'] = dfForm1['inspectionFreq'].str.lower()
    # uppercase
    dfForm1['locationTypes'] = dfForm1['locationTypes'].str.upper().str.replace(" ", "")
    # dateTypeTreatment
    dfForm1['dateOfInspection'] = pd.to_datetime(dfForm1['dateOfInspection'], errors='coerce')
    dfForm1['dateOfInspection'] = dfForm1['dateOfInspection'].dt.strftime('%Y/%d/%m') # change to object
    dfForm1['lastInspectionDate'] = pd.to_datetime(dfForm1['lastInspectionDate'], errors='coerce')
    dfForm1['lastInspectionDate'] = dfForm1['lastInspectionDate'].dt.strftime('%Y/%d/%m') # change to object

    return dfForm1

# %% [markdown]
# # Check Empty

# %%
def checkEmptyZone(dfForm1):
    #check null in masterZoneRef
    null_masterZoneRef = dfForm1['masterZoneRef'].isnull().any()
    empty_masterZoneRef = (dfForm1['masterZoneRef'] == '').any()
    #null_subZoneRef = dfForm1['subZoneRef'].isnull().any()
    #empty_subZoneRef = (dfForm1['subZoneRef'] == '').any()
    contains_issue = null_masterZoneRef or empty_masterZoneRef #or null_subZoneRef or empty_subZoneRef

    return contains_issue

# %% [markdown]
# # Login

# %%
def validatePassword(password):
    correct_password = getSpreadSheetPass()

    if password == correct_password:
        #messagebox.showinfo("Success", "Password is valid.")
        return True
    else:
        messagebox.showwarning("Warning", "Invalid password.")
        return False


def getSpreadSheetPass():
    # Construct the CSV export URL
    sheet_id = GOOGLE_SHEET_ID2
    worksheet_gid = '0' #'pass'
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={worksheet_gid}"

    try:
        dfPass = pd.read_csv(url)
        if dfPass.shape[0] == 0:
            correct_password = sysPassword
        else:
            correct_password = str(dfPass['Pass'][0])
    except:
        correct_password = sysPassword

    return correct_password

# %% [markdown]
# # Usage Stat

# %%
#deperciated
def ip():
  hostName = socket.gethostname()
  IPAddr = socket.gethostbyname(hostName)
  return hostName , IPAddr

# %%
def getSpreadSheetUsage():
    # Google Sheets ID and GID for the worksheet
    sheet_id = GOOGLE_SHEET_ID
    worksheet_gid = '0'  # Replace with the actual GID for the 'f1Usage' worksheet

    # Construct the CSV export URL
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={worksheet_gid}"

    # Load the data into a DataFrame
    dfF1Usage = pd.read_csv(url)
    usageVolume = len(dfF1Usage)

    return usageVolume

def writeDataToF1Temp(dfForm1):

    #https://www.google.com/search?q=how+to+use+google+sheet+api+in+python&oq=how+to+use+google+sheet+api+&gs_lcrp=EgZjaHJvbWUqCAgBEAAYFhgeMgoIABBFGBYYHhg5MggIARAAGBYYHjIICAIQABgWGB4yCAgDEAAYFhgeMggIBBAAGBYYHjIICAUQABgWGB4yCAgGEAAYFhgeMggIBxAAGBYYHjIICAgQABgWGB4yCAgJEAAYFhge0gEJMTA3MDBqMGo3qAIAsAIA&sourceid=chrome&ie=UTF-8#fpstate=ive&vld=cid:9b2c2b8a,vid:zCEJurLGFRk,st:0

    # Authenticate using the service account key file
    path = resourcePath('python-447003-6d8f2815ca54.json')
    gc = gspread.service_account(filename=path)
    sheet_id = GOOGLE_SHEET_ID
    worksheetName = 'f1Temp'
    spreadsheet = gc.open_by_key(sheet_id)
    worksheet = spreadsheet.worksheet(worksheetName)
    worksheet.clear()

    dfF1Temp = pd.DataFrame()
    dfF1Temp['f1RefNo'] = dfForm1['f1RefNo']
    dfF1Temp['masterZoneRef'] = dfForm1['masterZoneRef']
    import datetime
    dfF1Temp['time'] = datetime.datetime.now()
    set_with_dataframe(worksheet=worksheet, dataframe=dfF1Temp, include_index=False, include_column_header=True, resize=True)
    numForm1 = len(dfF1Temp)
    return dfF1Temp

def f1UsageUpdate(dfF1Temp):
    path = resourcePath('python-447003-6d8f2815ca54.json')
    gc = gspread.service_account(filename=path)
    sheet_id = GOOGLE_SHEET_ID
    worksheetName = 'f1Usage'
    spreadsheet = gc.open_by_key(sheet_id)
    worksheet = spreadsheet.worksheet(worksheetName)
    existing_data = worksheet.get_all_values()
    start_row = len(existing_data) + 1  # +1 to account for header
    # Write the DataFrame to the worksheet starting from the next empty row
    set_with_dataframe(worksheet=worksheet, dataframe=dfF1Temp, row=start_row, col=1, include_index=False, include_column_header=False)

    #updatedExisting_data = worksheet.get_all_values()
    #total_row = len(existing_data) + 1
    #print("DataFrame appended to f1Usage successfully!")

# %% [markdown]
# # Main Load Excel

# %%
## =========================== main ===========================
def loadTmcpExcel(filePath):

    startTime = time.time()

    #=============threading================
    global stop_thread
    keypress_thread = threading.Thread(target=check_for_exit, daemon=True).start()

    #============
    try:
        dfForm1 = pd.read_excel(filePath, sheet_name='generalInfo', \
                                dtype={'f1RefNo': str,'inspectionFreq': str, 'lampPost': str})
        # r'C:\Users\Adm\Desktop\parttime\tmcp\form1TMCP.xlsx'
        dfForm1 = cleanDf(dfForm1)
        contains_issue = checkEmptyZone(dfForm1)

        if (dfForm1.shape[0] == 0) or (contains_issue):
            # tinker
            messagebox.showwarning("Warning", "empty excel / missing masterZone / missing subZone !")
            #root.destroy()

        else:

            # looping
            for index, row in dfForm1.iterrows():
                #==============stop thread================
               
                if stop_thread:
                    print('stop pogram because of triagging esc')
                    return
            
                print(row['f1RefNo'])
                generalInfo(row['f1RefNo'], row['fileRef'], row['dateOfInspection'], row['lastInspectionDate'], row['inspectionFreq'])
                locationInfo(row['masterZoneRef'], row['subZoneRef'], row['lampPost'], row['locationTypes'], row['otherLocRemarks'])
                
                    
                #========================================


            dfF1Temp = writeDataToF1Temp(dfForm1)
            f1UsageUpdate(dfF1Temp)
            totalRunTime = time.time()-startTime
            messagebox.showwarning("Warning", f"completed. Total run time: {totalRunTime}")
            #root.destroy()
            exit()

    except Exception as e:
        print(f"An error occurred: {e}")

# %% [markdown]
# # Interface

# %%
def on_button_release(event):
    # Minimize or hide the interface after the button is released
    root.withdraw()  # Use iconify() to minimize instead

def submit():
    #filePath = filePath_entry.get()
    filePath = browseFile_entry.get()
    password = str(password_entry.get())

    #correct_password
    validateResult = validatePassword(password)

    current_date = datetime.now().strftime("%Y%m%d")
    target_date = "20261130"

    if (current_date <= target_date):
        if validateResult:
            #start program
            print("Start")
            loadTmcpExcel(filePath) # main
            root.destroy()
            #exit()
    else:
        messagebox.showerror("Warning", "Incorrect Password, retry")
        root.destroy()
        #exit()

def browse_file():
    filename = filedialog.askopenfilename()  # Open file dialog
    browseFile_entry.delete(0, tk.END)  # Clear current entry
    browseFile_entry.insert(0, filename)  # Insert the selected file path

# Create the main application window
root = tk.Tk()
root.title("FilePath and Password")

# Get screen dimensions
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Display screen dimensions
dimension_label = tk.Label(root, text=f"Screen dimensions: {screen_width} x {screen_height}")
dimension_label.pack(pady=5)

#tk.Label(root, text="Enter file full path:").pack(pady=5)
#filePath_entry = tk.Entry(root, width = 60)
#filePath_entry.pack(pady = 5)

# Create a browse view
tk.Label(root, text="Enter file full path:").pack(pady=5)
browseFile_entry = tk.Entry(root, width=60)
browseFile_entry.pack(pady=5)

# Create a browse button
browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.pack(pady=5)

tk.Label(root, text="Enter Passowrd:").pack(pady=5)
password_entry = tk.Entry(root, show="*", width = 60)
password_entry.pack(pady = 5)

# Create a button to start the process
submit_button = tk.Button(root, text="Submit", command = submit)
submit_button.pack(pady=20)

# Bind the button release event to the on_button_release function
submit_button.bind("<ButtonRelease-1>", on_button_release) # def on_button_release
usageVolume = getSpreadSheetUsage()
signature_label = tk.Label(root, text=f"Developed by Dev-Intelligent Arborist. Usage Volume: {usageVolume}")
signature_label.pack(pady=5)


# Start the GUI main loop
root.mainloop()

#加密
#https://comate.baidu.com/zh/page/9mb6z7ac6mi
#download upx win64
# cmd negative to the upx folder , then run below in cmd
#C:\Users\Adm\Desktop\parttime\upx-5.0.2-win64\upx-5.0.2-win64
#upx --best --lzma --force C:\Users\Adm\Desktop\parttime\tmcp\PCMT.exe


