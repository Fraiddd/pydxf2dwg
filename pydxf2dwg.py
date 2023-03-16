# coding: utf-8

from pyautocad import Autocad
from tkinter import Tk, filedialog, messagebox
import time

# Connect to Autocad.
acad = Autocad()

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
        acad.prompt('In Progress ...\n')
        for fil in file_path:
            # Extract file path name.
            nf = '/'.join(fil.split('/')[-1:])
            time.sleep(2)
            # Open a dxf
            doc.SendCommand('(vla-open (vla-get-documents (vlax-get-acad-object)) "' + fil + '")\n')
            time.sleep(4)
            # Activate 
            doc.SendCommand('(vla-activate (vla-item (vla-get-documents (vlax-get-acad-object)) "' + nf + '"))\n')
            time.sleep(3)
            # Reconnect to Autocad in the new draw.
            acad = Autocad()            
            doc = acad.ActiveDocument
            # Current layer "0"
            doc.SetVariable('clayer', "0")
            # Insertion unit in meter.
            doc.SetVariable('insunits', 6)
            # Focus
            acad.app.ZoomExtents()
            # Save with .dxf name.
            time.sleep(2)
            doc.SaveAs(fil[:-3] + 'dwg')
            time.sleep(2)
            doc.close()
            time.sleep(4)
            # Reconnect to Autocad in the drawing1.
            acad = Autocad()            
            doc = acad.ActiveDocument
    else:
        messagebox.showerror(title='Error',
                    message='No file, or you lost your way ...')


# Autocad check.
if acad:
    pydxf2dwg(acad)
else:
    messagebox.showerror(title='Error',
                    message='Autocad must be installed and strated, or unknown error')
