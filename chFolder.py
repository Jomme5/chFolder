# 04/01/2023 - Tijs Cools

from win32com.shell import shell, shellcon

BIF_NEWDIALOGSTYLE = 64 #extended dialog
BIF_UAHINT = 256 #usage hint

def chooseFolder(startpath=None, extended=False):
    """
    Typical Windows folderchooser-dialog starting on given basepath(startpath=) or desktop-root if no path is given. 
    Return value: string containting choosen path or None when canceled by user \n
    Options: 
    - extended: resizable dialog with posibility to make a new folder.
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

    this=shell.SHBrowseForFolder(0, # parent HWND
                            pidl, # root PIDL.
                            "Default of %s" % startpath ,# title
                            flags, #flags
                            )
    if this[1]==None:
        return None
    else:
        return(shell.SHGetPathFromIDListW(this[0]))
