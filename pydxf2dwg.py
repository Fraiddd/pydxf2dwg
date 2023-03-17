# coding: utf-8
# Python 3.11.1
'''
    pydxf2dwg 1.0

    convert dxf to dwg.

    No install

    Requirements: Autocad
    External Modules: pyautocad

    Only for Windows
    Tested on Windows 10 and Autocad 2015

    :No copyright: (!) 2023 by Frédéric Coulon.
    :No license: Do with it what you want.
'''
from pyautocad import Autocad
from tkinter import Tk, filedialog, messagebox as mb
import time
import win32gui
import win32con

# Connect to Autocad.
acad = Autocad()

def gethandlewin(winame):
    rslts = []
    ret = None
    win32gui.EnumWindows(lambda h, liste: liste.append(h), rslts)
    for handle in rslts:
        res = win32gui.GetWindowText(handle)
        if winame in res:
            ret = handle
    return ret

def pydxf2dwg(acad):
    doc = acad.ActiveDocument
    #  Start message.
    acad.prompt('\pydxf2dwg connected\n')
    # Files explorer
    root = Tk()
    # Hides the root window.
    root.withdraw()
    file_path = filedialog.askopenfilenames(initialdir = 'c:/',
                                        title = 'Select DXF files',
                                        filetypes = [('DXF files', '*.dxf *.DXF')])
    # If files were found.
    if file_path:
        # Hide Autocad (faster)
        hwin = gethandlewin('Autodesk AutoCAD')
        win32gui.ShowWindow(hwin, win32con.SW_HIDE)
        # Loop over files
        for fil in file_path:
            # Extract file path name.
            nf = '/'.join(fil.split('/')[-1:])
            # Figure of speech
            acdc = '(vla-get-documents (vlax-get-acad-object))'
            # Open a dxf
            doc.SendCommand('(vla-open ' + acdc + ' "' + fil + '")\n')
            time.sleep(2)
            # Activate 
            doc.SendCommand('(vla-activate (vla-item ' +
                            acdc + ' "' + nf + '"))\n')
            time.sleep(1)
            # Reconnect to Autocad in the new draw.
            acad = Autocad()            
            doc = acad.ActiveDocument
            # Save with .dxf name.
            doc.SaveAs(fil[:-3] + 'dwg')
            time.sleep(1)
            doc.close()
            time.sleep(1)
            # Reconnect to Autocad in the drawing1.
            acad = Autocad()            
            doc = acad.ActiveDocument
        win32gui.ShowWindow(hwin, win32con.SW_SHOW)
        win32gui.SetForegroundWindow(hwin)
        mb.showinfo(title='Done',
                    message= str(len(file_path)) + ' files processed.')
    else:
        mb.showerror(title='Error',
                    message='No file, or you lost your way ...')


# Autocad check.
if acad:
    pydxf2dwg(acad)
else:
    mb.showerror(title='Error',
        message='Autocad must be installed and strated, or unknown error')
