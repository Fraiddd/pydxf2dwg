# coding: utf-8

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
        hwin = gethandlewin('Autodesk AutoCAD')
        win32gui.ShowWindow(hwin, win32con.SW_HIDE)

        for fil in file_path:
            # Extract file path name.
            nf = '/'.join(fil.split('/')[-1:])
            acdc = '(vla-get-documents (vlax-get-acad-object))'
            # time.sleep(1)

            # Open a dxf
            doc.SendCommand('(vla-open ' + acdc + ' "' + fil + '")\n')
            time.sleep(2)
            # Activate 
            doc.SendCommand('(vla-activate (vla-item ' + acdc + ' "' + nf + '"))\n')
            time.sleep(1)
            # Reconnect to Autocad in the new draw.
            acad = Autocad()            
            doc = acad.ActiveDocument
            # Current layer "0"
            # doc.SetVariable('clayer', "0")
            # Insertion unit in meter.
            # doc.SetVariable('insunits', 6)
            # Focus
            # acad.app.ZoomExtents()
            # time.sleep(2)
            # Save with .dxf name.
            doc.SaveAs(fil[:-3] + 'dwg')
            time.sleep(1)
            doc.close()
            time.sleep(1)
            # Reconnect to Autocad in the drawing1.
            acad = Autocad()            
            doc = acad.ActiveDocument
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
