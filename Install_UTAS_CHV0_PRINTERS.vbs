''  Print Server Link to Desktop And Start Menu 
''  by Edward L. Thomas                                                     
''  Email: Edward@ThomasITServices.com Phone: 503-409-8918
''  Created:    6/30/2016                                                                           
''  Last Modified:  6/30/2016                                                                              
''  Last Modified By: Edward Thomas                                                            
''  Programming Language: VBScript
''  Location: "\\netapp3\aero_iet\software\Applications\UTAS\Printers\Install_UTAS_CHV0_PRINTERS.vbs"                                                                                                             

Option Explicit

'''''Start Up'''''
If WScript.Arguments.Named.Exists("elevated") = False Then
    'Launch the script again as administrator
    Dim StartMeUP : Set StartMeUP = createObject("Shell.Application")
    Call StartMeUP.ShellExecute("wscript.exe", chr(34) + WScript.ScriptFullName + chr(34) + " /elevated", "", "runas", 1)
                
Else

    '''' Setting Variables ''''

    Dim ProgramName: ProgramName = "Print Server Link to Desktop And Start Menu"
    Dim popOnTop : popOnTop = 4096

    Dim LinkName : LinkName = "Add Printers UTAS Chula Vista Site"
    Dim PrintServerPath : PrintServerPath = "\\chv0fp01"
    Dim IconLocation : IconLocation = "%SystemRoot%\system32\imageres.dll,46"


    Dim objShell : Set objShell = createobject("Wscript.shell")
    Dim StartMenu: StartMenu = objShell.SpecialFolders("AllUsersStartMenu") + "\Programs\"
    Dim DesktopPath: DesktopPath = objShell.SpecialFolders("AllUsersDesktop")  + "\"

    Dim oEnv : Set oEnv = objShell.Environment("PROCESS")

    '''''Process Block ''''

    Call AddPrintShortCut(StartMenu,LinkName,PrintServerPath,IconLocation)
    Call AddPrintShortCut(DesktopPath,LinkName,PrintServerPath,IconLocation)



    ''''Function Block '''''

    Function AddPrintShortCut(Location,LinkName,TargetPath,IconLocation)
        Dim Shortcut : Set Shortcut = objShell.CreateShortcut(Location + LinkName + ".lnk")
        Shortcut.TargetPath = TargetPath
        Shortcut.WorkingDirectory = TargetPath
        Shortcut.IconLocation = IconLocation
        Shortcut.Save()
        objShell.popup "Making was copied to " + Location,2,ProgramName,popOnTop
    End Function


    ''''End'''
    Wscript.Quit

End IF

