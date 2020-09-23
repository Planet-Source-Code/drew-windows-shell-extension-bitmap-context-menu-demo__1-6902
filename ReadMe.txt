ReadMe File: (In Notepad, turn WordWrap OFF)

How to Test this Project:
  1) Register BitmapShellMenu.dll.
  2) Import WinShellMenu.reg into the Registry.
  3) Right-Click on any .bmp file. (There are samples in this folder.)
     You should see a menu Item called 'Set Wallpaper'.

How Windows Context Menus (Shell Extensions) Work:
  The first time the 'right-click' menu pops up in the Windows Explorer, 
  Windows looks in the Registry for Shell Extension entries for the 
  specific file type selected(.bmp in our case).  If the file is a .bmp 
  file, Windows loads the object (in our case the BitmapShellMenu.dll) 
  associated with the ClassID entry found in the Registry for that file 
  type.  Therefore, Registry entries must be imported into the Registry 
  for your specific file type, so Windows knows how to call your Shell 
  Extension.

Import Registry Entries for this project:
  Included in this project is a file called: WinShellMenu.reg.  This file 
  needs to be imported into your Registry in order for this program to 
  work.  You can simply Double-Click on the file, or Run RegEdit.exe 
  and use the 'Registry' Menu's 'Import Registry File...' menu item 
  to import the file.


*************** Sample Contents of 'WinShellMenu.reg' file: ***************
REGEDIT4

[HKEY_CLASSES_ROOT\.bmp]
@="bmpfile"

[HKEY_CLASSES_ROOT\bmpFile]
@="Windows bmp - File"

[HKEY_CLASSES_ROOT\bmpFile\shellex\ContextMenuHandlers]
@="bmpFileMenu"

; ------ Modify the ClassID entry below -------------------------------------------
;  Replace {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx} with a valid ClassID entry...
; ---------------------------------------------------------------------------------
[HKEY_CLASSES_ROOT\bmpFile\shellex\ContextMenuHandlers\bmpFileMenu]
@="{810B2B4C-CF81-11D3-AF8B-005004322411}"

; ------ Modify the ClassID entry below -------------------------------------------
;  Replace {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx} with a valid ClassID entry...
; ---------------------------------------------------------------------------------
[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved]
"{810B2B4C-CF81-11D3-AF8B-005004322411}"="Shell Extension"
************************************************************************

Registry Entry Notes:

  1) [HKEY_CLASSES_ROOT\.bmp]
     This entry ensures that there is a .bmp entry on your HKEY_CLASSES_ROOT,
     and sets it's Default Value(@) to "bmpFile".  Windows uses "bmpFile" to
     lookup the next item(2).

  2) [HKEY_CLASSES_ROOT\bmpFile]
     This entry creates this "bmpFile" entry under [HKEY_CLASSES_ROOT], and 
     sets it Default Value(@) to "Windows bmp - File", which is just a 
     description of this item.  The next item(3) creates the actual Context Menu 
     pointer to the ClassID.

  3) [HKEY_CLASSES_ROOT\bmpFile\shellex\ContextMenuHandlers]
     This entry creates the Context Menu Shell Extension Handler item, and sets
     it's Default Value(@) to "bmpFileMenu".  Where "bmpFileMenu" is the name of
     the Context Menu Handler to be called for this file type.

  4) [HKEY_CLASSES_ROOT\bmpFile\shellex\ContextMenuHandlers\bmpFileMenu]
     This entry identifies the ClassID to be called to Handle this Context Menu.
     This is the ClassID of your object(.dll).

  5) [HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved]
     This entry tells Windows that this ClassID is Approved to run in the
     'CurrentVersion' of Windows.  "Shell Extension" is just a description of this item.


Menu Item Bitmap:
  It seems that the Menu Item Bitmap should be Width: 13 x Height: 15.  When the 
  Menu Image is created it is resized/stretched.  This size seems to work best, by
  not causing the image to be distorted.


Compiling the Project:
  When you re-compile this project, the .dll may still be loaded in memory.  If you
  are running Windows NT, you can try and unload the DLL using the Task Manager.  In
  any case, if you cannot recompile this project because the .dll is loaded in memory,
  you may need to End Task on Windows(Explorer.exe), or even reboot your system.

  This Project maintains 'Binary Compatability' to retain it's ClassID (CLSID).  If
  you don't maintain 'Binary Compatability', your ClassID will change every time you
  re-compile, thus, causing you to have to modify your WinShellMenu.reg file, and 
  re-import it into the Registry.


CUSTOMIZING THIS PROJECT for your purposes:
  First:
    COPY/BACKUP this Project!

  Second:
    Open the Project Properties, and change the Project Name to your Project's Name.
    This will break the 'Binary Compatability', but that's ok.  Once you compile your
    Project for the first time, make sure that you reset the project to use 'Binary
    Compatability' again.

  Third:
    Search through the Project for this text string: CHANGE HERE.  
    Then, modify the the Project Class Name constant(kPROGRAM_CLASS_NAME) to match your 
    new Project Name.  Also, modify the Menu Text constant(kMENU_ITEM_TEXT) to what 
    you want the Menu Item to say.  Next, change the code to Execute your code, or 
    shell out to your program.
  
  Fourth:
    Open the WinShellMenu.res Repository file, and add your 13x15 bitmap image with the
    name 101.

  Fifth:
    The Original project has 2 Reference to 2 type libraries in it's folder.  The files
    are: IctxMenu.tlb and IDataObj.tlb.  If you have copied this project to a new folder
    to modify it, you should probably remove the reference to the type libraries in 
    the other folder, and add them to point to your new folder.  It should work if you
    don't do this, but just to be safe, and point to the current location, it's a good
    idea.

  Last:
    Compile your program.  Then, Search the Registry for your ProjectName.ClassName.  
    Once you find the entry, copy the ClassID and use it to modify your WinShellMenu.reg 
    file, and import this file into the Registry.

  Now TEST YOUR PROGRAM!


That's it!  Good Luck...




