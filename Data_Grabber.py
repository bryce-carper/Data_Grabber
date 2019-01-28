#############################################################################################
#
#   Mouse-and-key automation for pulling historical NunyaCorp logistics data
#
#############################################################################################


'''
Hey Coworker/Anyone Else Reading This:
This should do what you asked for, downloading and archiving all files over the course of a ~11 hour period, and then exiting all
programs involved before shutting down the workstation, all via simulated mouse-and-key usage.
To run this script, have it in a single JuPyteR Notebook cell in a Chrome window maximized on the right-hand side of your screen,
and a BigFirm Partner Analytics window maximized on the left. Leave the BigFirm Partner Analytics default page tab open, and have
the SuperBizServices tab for the logistics reporting app open, to the Logistics Detail menu, set to the earliest week for both
delivery rate and shipping manifest validity. Leave it on the DELIVERY RATE tab. Have nothing but Chrome, Anaconda, and JuPyteR
Notebook open. Set the shipment logistics field to "ALL SHIPMENTS", and make sure that initial weeks are set in both the DELIVERY
RATE and MANIFEST VALIDITY tabs. Also might as well clear the downloads folder before running it.
If for some reason this script is pressed into service in some bleak, distant future: Verify that clicks still occur where needed,
as the script is sensitive to changing out displays and it's also possible that the SuperBizServices logistics site has changed.
'''



import win32api as wapi
import win32con as wcon
import win32gui as wgui
import win32clipboard as clip
import openpyxl as pyxl
import pandas
import os
import time


# Paths to files, silly!
inPath = r'C:\Users\acoworker\Downloads\Data Download.xlsx'
outPath = r'C:\Users\acoworker\Desktop\Notes\Logistics Data\Historical Data\\'



# Cooordinates for the last few weeks in the weeks drop-down menu, the ones that can't be got by the weekUp() function.
# lastWeeksCoords = [[-1218, 625, -1320, 678], [-1218, 608, -1320, 661], [-1218, 591, -1320, 644], [-1218, 574, -1320, 627],
#                    [-1218, 557, -1320, 610], [-1218, 540, -1320, 593], [-1218, 523, -1320, 576], [-1218, 506, -1320, 559],
#                    [-1218, 489, -1320, 542], [-1218, 472, -1320, 525], [-1218, 455, -1320, 508], [-1218, 438, -1320, 491],
#                    [-1218, 421, -1320, 474], [-1218, 404, -1320, 457], [-1218, 387, -1320, 440], [-1218, 370, -1320, 423],
#                    [-1218, 353, -1320, 406]]

def grabMouse(delay=3): # Helpful tool in building this code. Returns the (x,y) coordinates of where the cursor is.
    time.sleep(delay)
    return wgui.GetCursorPos()

def click(x,y): # Simulate click at (x,y). From upper-left of primary monitor.
    #Secondary monitor to the left has negative x-values.
    time.sleep(.1)
    oldCoordinates = wgui.GetCursorPos()
    wapi.SetCursorPos((x,y))
    wapi.mouse_event(wcon.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    wapi.mouse_event(wcon.MOUSEEVENTF_LEFTUP,x,y,0,0)
    wapi.SetCursorPos(oldCoordinates)

def dragSelect(x1,y1,x2,y2): # Yeah, these are pretty self-explanatory.
    time.sleep(.2)
#     oldCoordinates = wgui.GetCursorPos()
    wapi.SetCursorPos((x1,y1))
    time.sleep(.2)
    wapi.mouse_event(wcon.MOUSEEVENTF_LEFTDOWN,x1,y1,0,0)
    time.sleep(.2)
    wapi.SetCursorPos((x2,y2))
    time.sleep(.2)
    wapi.mouse_event(wcon.MOUSEEVENTF_LEFTUP,x2,y2,0,0)
    time.sleep(.2)
#     wapi.SetCursorPos(oldCoordinates)

def ctrlC():
    time.sleep(.2)
    wapi.keybd_event(0xA2,0,wcon.KEYEVENTF_EXTENDEDKEY | 0,0) # Start holding left ctrl
    time.sleep(.2)
    # Simulating a held keypress is a little more involved. See 3rd argument.
    wapi.keybd_event(67,0,0,0) ## Simulating a held keypress is a little more involved. See 3rd argument.
    time.sleep(.2)
    wapi.keybd_event(0xA2,0,wcon.KEYEVENTF_EXTENDEDKEY | wcon.KEYEVENTF_KEYUP,0) # Let up left ctrl.
    time.sleep(.2)

def grabXlsx(pathName=inPath):
    time.sleep(3)
    for i in range(10):
        click(-1251,202) # Click SHIPMENT/ITEM DOWNLOAD link.
        time.sleep(1)
    time.sleep(.4)
    t = 0 # Set a counter to check if file is downloaded
    while t < 3600: # Wait up to an hour.
        t += 1
        time.sleep(1)
        if os.path.isfile(pathName): # Check if file exists.
            break
    t=0
    oldSize = -1
    strikes = 0
    while t < 30 and strikes < 3: # Start checking if the file's size is increasing. Stop when it isn't.
        t += 1
        time.sleep(.4)
        newSize = os.path.getsize(pathName)
        if newSize > oldSize and newSize > 0:
            oldSize = newSize
        else:
            strikes += 1
    time.sleep(3)
    click(-19, 1056) # Close downloads ribbon
    time.sleep(.4)
    click(-1346, 12) # Close download tab
    time.sleep(3)
    click(-1346, 12)
    time.sleep(3)
    click(-1346, 12)

def digestXlsx(pathIn=inPath, pathOut = outPath): # Turns .Xlsx into 3 .CSVs in the right directory, marked by BigFirm fiscal date.
    # Also deletes downloaded file, and returns the fiscal date from the delivery rate data file.
    wb = pyxl.load_workbook(pathIn) # Load .xlsx in OpenPyXl.

    df = pandas.DataFrame(wb['MANIFEST VALIDITY'].values) # Load shipping manifest validity worksheet in pandas.
    fiscDate=str(list(df[0])[1]) # Determine fiscal date for worksheet.
    spec = fiscDate
    df.to_csv(pathOut+'MANIFEST VALIDITY\\'+fiscDate+'-'+'ValidManifestRate'+'.csv',index=False,header=False) # Write out as CSV.

    df = pandas.DataFrame(wb['DELIVERY MET'].values) # Same thing for on time rate.
    fiscDate=str(list(df[0])[1])
    spec += ' ' + fiscDate
    df.to_csv(pathOut+'DELIVERY MET\\'+fiscDate+'-'+'DeliveryMetRate'+'.csv',index=False,header=False)

    df = pandas.DataFrame(wb['DELIVERY RATE'].values) # Same thing for delivery rate.
    fiscDate=str(list(df[0])[1])
    spec += ' ' + fiscDate
    df.to_csv(pathOut+'DEVIVERY RATE\\'+fiscDate+'-'+'DeliveryRate'+'.csv',index=False,header=False)

    print(spec)
    os.remove(pathIn) # Remove downloaded .Xlsx.
    return spec # Return the date from the delivery rate file.

def grabLoadText(): # grabs text around the loading notification for SuperBizServices, so we can see what it says.
    time.sleep(.2)
    dragSelect(-945,540,-878,558) # Select the central area.
    ctrlC() # Copy selection.
    time.sleep(.2)
    clip.OpenClipboard() # Read the clipboard's contents into clipText, and then close clipboard.
    clipText = clip.GetClipboardData()
    clip.CloseClipboard()
    return clipText # Return clopboard contents.

def waitForSuperBizServices(): # Decides when a SuperBizServices report is done loading.
    strikes = 0
    time.sleep(.4)
    text = ''
    text = grabLoadText() # See what the load text says.
    while 'Loading' in text or 'being generated' in text or strikes < 3: # Does it look like it's still running?
        time.sleep(3)
        clip.OpenClipboard()
        clip.EmptyClipboard()
        clip.SetClipboardText('')
        clip.CloseClipboard()
        text = grabLoadText()
        if ('Loading' not in text and 'being generated' not in text) or text == '': # If not, try a few more times.
            strikes += 1
    time.sleep(3) # Wait a bit more once you're sure.


# def weekUp(): # Adjusts one week up, provided the week appears at the bottom of the pull-down!
#     click(-1271, 337) # click week pull-down tab.
#     time.sleep(3)
#     click(-1430, 665) # Click one week up from the bottom.
#     waitForSuperBizServices() # Wait for report to catch up.
#     click(-1728, 177) # Change to the shipping manifest validity tab.
#     waitForSuperBizServices() # Wait on report.
#     click(-1293, 379) # Click the shipping manifest tab week pull-down menu.
#     time.sleep(3)
#     click(-1451, 702) # Click one week up from the bottom on THAT menu.
#     waitForSuperBizServices() # Still waiting.
#     click(-1882, 177) # Click back to the delivery rate tab.
#     waitForSuperBizServices()



def weekUp(tabs=18):
    for i in range(tabs):
        time.sleep(.3)
        wapi.keybd_event(9,0,0,0)
    time.sleep(.3)
    wapi.keybd_event(0x0D,0,0,0)
    time.sleep(.3)
    wapi.keybd_event(0x26,0,0,0)
    time.sleep(.3)
    wapi.keybd_event(0x0D,0,0,0)




time.sleep(3) # Pause at the start, to get ready.

strikes = 0
for i in range(300): # Keep pulling the next week from the earliest, break the loop if we hit the end.
    grabXlsx() # Grab this week's data.
    check = digestXlsx() # Digest it, check if we're at the end
    if check == '  ': # If we are, break the loop.
        break
    weekUp(9)
    waitForSuperBizServices()
    click(-1280,278) # Change to the shipping manifest validity tab.
    waitForSuperBizServices()
    click(-1269, 354)
    weekUp(9)
    waitForSuperBizServices()

    grabXlsx() # Grab this week's data.
    check = digestXlsx() # Digest it, check if we're at the end
    if check == '  ': # If we are, add a strike.
        strikes += 1
        if strikes == 5:
            break
    weekUp(9)
    waitForSuperBizServices()
    click(-1280,278) # Change to delivery rate tab.
    waitForSuperBizServices()
    click(-1269, 354)
    weekUp(8)
    waitForSuperBizServices()



time.sleep(15)
click(-27,7) # Close the left-most Chrome window.
time.sleep(1)
click(522,1058) # Bring up the Anaconda Navigator start menu.
time.sleep(1)
click(1597,97) # Close Anaconda Navigator
time.sleep(1)
click(1166,493) # Yes, we REALLY want to leave Conda.
click(1166,493)
click(1166,493)
time.sleep(1)
click(1891,7) # Start closing out of JuPyter NB
time.sleep(1)
click(30,1056) # Click the Windows button
time.sleep(1)
click(295,1008) # Issue a shutdown command.
time.sleep(.05) # Move quickly to exit JuPyter NB
for i in range(10):
    click(1027, 182)
time.sleep(.05)
