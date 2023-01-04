# 04/01/2023 - Tijs Cools

from win32com.shell import shell, shellcon

BIF_NEWDIALOGSTYLE = 64 #extended dialog
BIF_UAHINT = 256 #usage hint

def chooseFolder(startpath=None, extended=False):
    """
    Typical Windows folderchooser-dialog. \n
    
    Options(all are optional): 
    - hwnd: handler of parentwindow(e.g. Tkinter: root.winfo_id())
    - startpath: basepath to start from. Desktop-root if no path is given.
    - extended: (True|False) resizable dialog with posibility to make a new folder.
    Return value: string containting choosen path or None when canceled by user
    """

    flags = shellcon.BIF_STATUSTEXT | shellcon.BIF_RETURNONLYFSDIRS
    if extended:
        flags = flags | BIF_NEWDIALOGSTYLE | BIF_UAHINT
    if startpath==None:
        pidl=None
    else:
        startpath=startpath.replace('/','\\')
        desktop = shell.SHGetDesktopFolder()
        cb, pidl, extra = desktop.ParseDisplayName(0, None, startpath)
    p_hwnd=pywintypes.HANDLE(hwnd)
    this=shell.SHBrowseForFolder(p_hwnd, # parent HWND
                            pidl, # root PIDL.
                            "Default of %s" % startpath ,# title
                            flags, #flags
                            )
    if this[1]==None:
        return None
    else:
        return(shell.SHGetPathFromIDListW(this[0]))
