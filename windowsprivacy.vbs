' The MIT License (MIT)
' 
' Copyright (c) 2015 Boris Meyer
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.


' Constants
Const ForReading   = 1
Const ForWriting   = 2
Const ForAppending = 8
Const Chars = "abcdefghijklmnopqrstuvwxyz"

' Registry Constants
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA      = &h80000006

' Globals
Dim fso, wmi, ws, os, sys
Dim dictUpdates, dictTasks, dictServices
Dim boolUpdatesDisabled, boolStartedLocal, strTEMPFolderPath, strINIFilePath, strPath, boolTEMPLocation, boolUserInteraction, boolAdmin, intOSArch

' Objects
Set ws = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set os = CreateObject("Shell.Application")
Set sys = CreateObject("Microsoft.Update.SystemInfo")

' Dictionaries
Set dictUpdates = CreateObject("Scripting.Dictionary")
Set dictInstalledUpdates = CreateObject("Scripting.Dictionary")
Set dictTasks = CreateObject("Scripting.Dictionary")
Set dictServices = CreateObject("Scripting.Dictionary")
Set dictAcceptedRegistryTypes = CreateObject("Scripting.Dictionary")

' Returns True if KB ID is found installed
Function SearchUpdate(kb)
  Set colItems = wmi.ExecQuery("Select * from Win32_QuickFixEngineering where HotFixID = 'kb" & kb & "'")
  For Each colItem In colItems
    If InStr(1, LCase(colItem.HotFixId), LCase(kb), 1) > 0 Then
      SearchUpdate = True
      Exit Function
    End If 
  Next
  SearchUpdate = False
End Function 


' Removes a single Update
Sub RemoveUpdate(kb)
  ' If found
  If SearchUpdate(kb) = True Then 
    ws.Exec("wusa.exe /KB:" & kb & " /uninstall /quiet /norestart")
    Do 
      WScript.Sleep 3000
    Loop Until CheckRunningProcess("wusa.exe") = False
    
    'Rescan
    If SearchUpdate(kb) = False Then
      Logger("... uninstalled")
    Else
      Logger("... failed to uninstall!")
    End If
  End If
End Sub 


' Returns True if process found
Function CheckRunningProcess(strProcess)
  For Each Process in GetObject("winmgmts://.").InstancesOf("Win32_Process")
    If Process.Name = strProcess Then
      CheckRunningProcess = True
      Exit Function 
    End If
  Next
  CheckRunningProcess = False
End Function


' Foreach shit update in installed updates: Call remover
Sub RemoveUpdates()
  
  Logger(vbCrLf)
  Logger("+-----------------------+")
  Logger("¦  Removing Updates...  ¦")
  Logger("+-----------------------+")
  Logger(vbCrLf)
  
  ' Key, Values
  strKey = dictUpdates.Keys
  strValue = dictUpdates.Items
  
  ' Foreach
  For i = 0 To dictUpdates.Count-1
    If dictInstalledUpdates.Exists("kb" & strKey(i)) Then
      Logger("KB" & strKey(i) & " Removing " & strValue(i))
      Call RemoveUpdate(strKey(i))
    Else
      Logger("KB" & strKey(i) & " is not installed")
    End If
  Next

End Sub


' Read INI File.
Sub ReadINIFile(mySection)
  
  strSection  = Trim(mySection)
  
  ' INI Exists
  If fso.FileExists(strINIFilePath) Then
  
  ' Open INI
    Set objIniFile = fso.OpenTextFile(strINIFilePath, ForReading, False)
    
    ' Until EOF
    Do While objIniFile.AtEndOfStream = False
      
      ' The Line
      strLine = Trim(objIniFile.ReadLine)
      
      ' Check if section is found in the current line
      If LCase(strLine) = "[" & LCase(strSection) & "]" Then
        
        ' Next line?
        strLine = Trim(objIniFile.ReadLine)
        
        ' Parse lines until the next section is reached
        Do While Left(strLine, 1) <> "["
          
          ' Contain data
          If Len(strLine) > 0 Then
            
            ' Updates
            If strSection = "Updates" Then
            
              ' Key/Value
              If InStr(1, strLine, "=", 1) > 0 Then
              arrData = Split(strLine, "=")
              
              strKB = Trim(arrData(0))
              strInfo = Trim(arrData(1))
          
              If Left(LCase(strKB), 2) = "kb" Then
                strKB = Mid(strKB, 3, Len(strKB))
              End If
              
              If dictUpdates.Exists(strKB) Then
                Logger("Warning: You have defined update """ & strKB & """ in INI file more than once...")
              Else
                dictUpdates.Add strKB, strInfo
              End If
              
             End If
            
            ' Tasks
            ElseIf strSection = "Tasks" Then
            
        If dictTasks.Exists(strLine) Then
          Logger("Warning: You have defined task """ & strLine & """ in INI file more than once...")
        Else
          dictTasks.Add strLine, ""
        End If

           
            ' Services
            ElseIf strSection = "Services" Then
            
              If InStr(1, strLine, "=", 1) > 0 Then
                arrData = Split(strLine, "=")
                strService = Trim(arrData(0))
                strAction = Trim(arrData(1))

        If dictServices.Exists(strService) Then
          Logger("Warning: You have defined service """ & strService & """ in INI file more than once...")
        Else
          dictServices.Add strService, strAction
        End If
                        
                
              End If
            
            ' Registry
            ElseIf strSection = "Registry" Then
              
              If Left(strLine, 5) = "HKEY_" Then
                
                arrData = Split(strLine, "||")
                
                arrRootKey = Split(arrData(0), "\") ' HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa
                strRoot = UCase(arrRootKey(0)) ' HKEY_LOCAL_MACHINE
                strKey = arrData(0) ' Copy Array
                strKey = Replace(strKey, strRoot & "\", "") ' SYSTEM\CurrentControlSet\Control\Lsa
          
                strValueName = arrData(1)   ' LmCompatibilityLevel
                strType = arrData(2)    ' DWORD
                strValue = arrData(3)    ' 00000002
                
                ' Deletion
                If strType = "DELETE" Then

                  ' Hässlich...
            Select Case strRoot
              Case "HKEY_LOCAL_MACHINE" Call RegistryDelete(HKEY_LOCAL_MACHINE, strKey, strValueName)
              Case "HKEY_CURRENT_USER" Call RegistryDelete(HKEY_CURRENT_USER, strKey, strValueName)
              Case "HKEY_CLASSES_ROOT" Call RegistryDelete(HKEY_CLASSES_ROOT, strKey, strValueName)
              Case "HKEY_USERS" Call RegistryDelete(HKEY_USERS, strKey, strValueName)
              Case Else Logger("Invalid Root " & strRoot & " to perform registry operation.")
            End Select

                Else
                
                  ' Accepted Types
                  If dictAcceptedRegistryTypes.Exists(strType) Then
                    
                    ' Check if Key exists
                    arrKey = Split(strKey, "\")
                    tmpKey = ""
                    
                    For i=0 To UBound(arrKey)
                On Error Resume Next
                  tmpKey = tmpKey & "\" & arrKey(i)
                  ws.RegRead(strRoot & tmpKey & "\")
                  
                  If Err.Number <> 0 Then
                  
                    ' Root from Constants
                    Select Case strRoot
                      Case "HKEY_LOCAL_MACHINE" Call RegistryWrite("KEY", HKEY_LOCAL_MACHINE, strKey, "", "")
                      Case "HKEY_CURRENT_USER" Call RegistryWrite("KEY", HKEY_CURRENT_USER, strKey, "", "")
                      Case "HKEY_CLASSES_ROOT" Call RegistryWrite("KEY", HKEY_CLASSES_ROOT, strKey, "", "")
                      Case "HKEY_USERS" Call RegistryWrite("KEY", HKEY_USERS, strKey, "", "")
                      Case Else Logger("Invalid Root " & strRoot & " to perform registry operation.")
                    End Select
                  
                  End If
                On Error Goto 0
                    Next
                    
                    ' Root from Constants
              Select Case strRoot
                Case "HKEY_LOCAL_MACHINE"
                  Call RegistryWrite(strType, HKEY_LOCAL_MACHINE, strKey, strValueName, strValue)
                Case "HKEY_CURRENT_USER"
                  Call RegistryWrite(strType, HKEY_CURRENT_USER, strKey, strValueName, strValue)
                Case "HKEY_CLASSES_ROOT"
                  Call RegistryWrite(strType, HKEY_CLASSES_ROOT, strKey, strValueName, strValue)
                Case "HKEY_USERS"
                  Call RegistryWrite(strType, HKEY_USERS, strKey, strValueName, strValue)
                Case Else Logger("Invalid Root " & strRoot & " to perform registry operation.")
              End Select
  
                  Else
                    Logger("Invalid Type " & strType & " to perform registry operation.")
                  End If

                End If

              End If ' End HKEY_
              
            End If ' End strSection
            
          End If ' End Line contains data
          
          
          ' Abort if the end of the INI file is reached
          If objIniFile.AtEndOfStream Then Exit Do
          
          ' Continue with next line
          strLine = Trim(objIniFile.ReadLine)
      
        Loop
      
      Exit Do
      
      End If ' End if [SECTION]
    
    Loop
    
    objIniFile.Close
    
  Else
    Logger("Missing INI file """ & strINIFilePath & """. Exiting...")
    Call Exiter
  End If ' End if INI Exists
  
End Sub


' Disable given Task
Sub DisableTask(strTask)
  Logger("Disabling Task """ & Replace(strTask, "\Microsoft\Windows\", "") & """")

  strCommand = "schtasks /Change /TN """ & strTask & """ /DISABLE"
  strResult = ws.Run("%comspec% /c " & strCommand, 0, True)
  
  If strResult = 1 Then
    Logger("Task """ & strTask & """ not found...")
  ElseIf strResult <> 0 Then
    Logger("Error (" & strResult & ") while disabling Task """ & strTask & """")
  End If
End Sub


' Disable Tasks
Sub DisableTasks()
  Logger(vbCrLf)
  Logger("+-----------------------+")
  Logger("¦  Disabling Tasks...   ¦")
  Logger("+-----------------------+")
  Logger(vbCrLf)
  
  strKey = dictTasks.Keys
  For i = 0 To dictUpdates.Count-1
    Call DisableTask(strKey(i))
  Next
End Sub


' Returns Folder Location of Script
Sub GetScriptLocation()
  ' Set if started localy
  If Not Left(WScript.ScriptFullName, 2) = "\\" Then
  
  ' What if mounted?
  Set drv = fso.GetDrive(Left(WScript.ScriptFullName, 1))
  If Not drv.DriveType = 2 Then
    ' Not a "Fixed" drive, not localy started
    ' 0 = "Unknown"
    ' 1 = "Removable"
    ' 2 = "Fixed"
    ' 3 = "Network"
    ' 4 = "CD-ROM"
    ' 5 = "RAM Disk"
    ' Mounted network location
    boolStartedLocal = False
  Else
    ' Fixed drive
    boolStartedLocal = True
  End If
  End If
  
  ' Set when local location temp
  If strTEMPFolderPath = Replace(WScript.ScriptFullName, "\" & WScript.ScriptName, "") Then
    boolTEMPLocation = True
  End If
  
  ' Global strPath
  strPath = Replace(WScript.ScriptFullName, "\" & WScript.ScriptName, "") 
End Sub


' Filename with extension INI
Function CreateScriptINIFileName()
  CreateScriptINIFileName = Replace(WScript.ScriptName, "." & fso.GetExtensionName(WScript.ScriptName), "") & ".ini"
End Function


' Hide Windows Updates which are listed in dictionary dictUpdates
Sub HideUpdates()
  Logger(vbCrLf)
  Logger("+-----------------------+")
  Logger("¦   Hiding Updates...   ¦")
  Logger("+-----------------------+")
  Logger(vbCrLf)
  Logger("Searching for pending updates. This may take a while...")
  
  Set updateSession = CreateObject("Microsoft.Update.Session")
  Set updateSearcher = updateSession.CreateUpdateSearcher()
  
  updateSearcher.ServerSelection = 2 ' ssWindowsUpdate
  updateSearcher.Online = True ' bypass WSUS server
  
  Set searchResult = updateSearcher.Search("IsInstalled=0")
  
  If searchResult.Updates.Count <> 0 Then
    
    Logger("Found " & CStr(searchResult.Updates.Count) & " updates...")
    
    For i = 0 To searchResult.Updates.Count - 1
        
        Set update = searchResult.Updates.Item(i)
          
        For Each kb in update.KBArticleIDs
          
          If dictUpdates.Exists(kb) Then
            If update.IsHidden = False Then
              Logger("Hiding update " & update.Title)
              update.IsHidden = True
            Else
              Logger("Already hidden: " & update.Title)
            End If            
          End If
          
        Next

    Next

  Else
    Logger("No pending updates found.")
  End If
  
End Sub


' Handles services as defined in INI file
Sub HandleServices()
  Logger(vbCrLf)
  Logger("+-----------------------+")
  Logger("¦ Configuring Services  ¦")
  Logger("+-----------------------+")
  Logger(vbCrLf)
  
  ' Key, Values
  strService = dictServices.Keys
  strAction = dictServices.Items
  
  ' Foreach service
  For i = 0 To dictServices.Count-1
    
    ' Get Service
    Set objServices = wmi.ExecQuery("Select * From Win32_Service Where Name='" & strService(i) & "'")
    
    If objServices.Count > 0 Then
      
      For Each objService In objServices
        
        ' Disable
        If strAction(i) = "disable" Then
          If LCase(objService.StartMode) = "disabled" Then
            Logger("Service " & objService.DisplayName & " already disabled")
          Else
            Logger("Disabling service " & objService.DisplayName)
            objService.StopService()
            objService.Change  , , , , "Disabled" 'Automatic, Manual, Disabled, Stopped
            Logger("Startmode: " & objService.StartMode)
          End If
        End If
        
        ' Delete
        If strAction(i) = "delete" Then
          Logger("Deleting service " & objService.DisplayName)
            objService.StopService()
            objService.Delete()
        End If
      
      Next
      
    Else
      Logger("Service " & strService(i) & " not found")
    End If
    
  Next
End Sub


' Sets Service to a startmode.
Sub SetService(strService, strActiontype)

  ' strService = Service Name
  ' strActiontype = What to set the service to. (Automatic, Manual, Disabled, Stopped)
  
  Set objServiceStatus = wmi.ExecQuery("Select * From Win32_Service Where Name='" & strService & "'")

  For Each objService in objServiceStatus
    Call Logger("Setting service mode of " & objService.DisplayName & " to " & strActiontype)
    objService.Change  , , , , strActiontype
    
    If LCase(strActiontype) = "disabled" Or LCase(strActiontype) = "stopped" Then
      Call Logger("Stopping service " & objService.DisplayName)
      objService.StopService()
    Else
      Call Logger("Starting service " & objService.DisplayName)
      objService.StartService()
    End If
    
  Next
  
End Sub


' Creates a dictionary of installed updates
Sub InstalledUpdates()
  Logger("Creating dictionary of installed updates...")
  
  Set colQuickFixes = wmi.ExecQuery("Select * from Win32_QuickFixEngineering")
  
  For Each objItem in colQuickFixes
    On Error Resume Next
    dictInstalledUpdates.Add LCase(objItem.HotFixID), objItem.Description
    On Error Goto 0
  Next
End Sub


' Check if update service available and NOT disabled
Sub CheckUpdateService()
  Set objServiceStatus = wmi.ExecQuery("Select * From Win32_Service Where Name='wuauserv'")
  
  If objServiceStatus.Count > 0 Then
    For Each objService In objServiceStatus
      If LCase(objService.StartMode) = "disabled" Then
        boolUpdatesDisabled = True
        Logger("Enabling Windows Update service to scan for updates to hide...")
        Call SetService("wuauserv", "Automatic")
      End If
    Next
  End If
End Sub


' If Windows Update has been disabled by User restore his settings
Sub RevertUpdateService
  If boolUpdatesDisabled Then
    Logger("Setting Windows Update service to disabled.")
    Call SetService("wuauserv", "Disabled")
  End If
End Sub


' Log to console/file
Sub Logger(strData)
  If boolUserInteraction = True then
    WScript.Echo " " & strData
  End If
End Sub


' In case started over windows with wscript, enforce restart with console
' If started from temp folder delete files.
Sub Exiter()

  ' Wait on exit to show summary
  If boolUserInteraction = True Then
    WScript.StdIn.Read(1)
  End If

  ' If temp, cleanup
  If boolTEMPLocation = True Then
    On Error Resume Next
      ' Delete Temp .vbs
      fso.DeleteFile WScript.ScriptFullName, True
      ' Delete Temp .ini
      fso.DeleteFile strINIFilePath, True
    On Error Goto 0
  End If
  
  ' Finally exit
  WScript.Quit
End Sub


' Wait for user input on exit
Sub ExitInfo()
  Logger(vbCrLf)
  Logger("+-----------------------+")
  Logger("¦         Done!         ¦")
  Logger("+-----------------------+")
  Logger(vbCrLf)
  Logger("Note that you should restart your system after")
  Logger("using this tool. Please do that by yourself.")
  Logger(vbCrLf)
  Logger("Press any key to close this console...")
  Call Exiter
End Sub


' Abort if update service needs reboot
Sub UpdateServiceRequiresReboot()
  If sys.RebootRequired = True Then
    Logger(vbCrLf)
    Logger("+-----------------------+")
    Logger("¦      Sorry :'(        ¦")
    Logger("+-----------------------+")
    Logger(vbCrLf)
    Logger("Windows Update service has an reboot pending!")
    Logger(vbCrLf)
    Logger("Since we want a clean system, these updates need to be")
    Logger("configured completely by performing an reboot.")
    Logger(vbCrLf)
    Logger("Please reboot this system and start this script again.")
    Logger("...Exiting.")
    Call Exiter
  End If
End Sub


' Handles services as defined in INI file
Sub HandleRegistry()
  Logger(vbCrLf)
  Logger("+-----------------------+")
  Logger("¦       Registry        ¦")
  Logger("+-----------------------+")
  Logger(vbCrLf)
  
  Call ReadINIFile("Registry")
End Sub

' Handle startup
Sub CheckLocalStartup()
  
  ' Should be solid for ([console|windows][uac|no uac]) 
  
  ' No Administrative permissions, restart elevated
  If Not boolAdmin = True Then
    Call ExecuteElevated(WScript.ScriptFullName)
  End If
  
  ' If Network Location
  If Not boolStartedLocal Then
    
    ' Network sources, when not in trusted network zone, may show and cause
    ' strange behaviour. Because of that I'll copy this script to temporary folder
    ' with a random name to avoid injecting over malicious INI
    ' ... or maybe I'm a bit too paranoid. However, safer
    
    ' Get a random name (8 Chars)
    strRandomName = RandomName()
    
    ' Until valid name
    Do While True
      If  fso.FileExists(strTEMPFolderPath + "\" + strRandomName + ".vbs") Or _ 
        fso.FileExists(strTEMPFolderPath + "\" + strRandomName + ".ini") Then
        
        strRandomName = RandomName()
        ' We are not aggressive.
        WScript.Sleep 50
      Else
        ' Done. We have a free name
        Exit Do
      End If
    Loop
    
    ' No 8 Chars? Something wrong.
    If Len(strRandomName) <> 8 Then
      boolCritical = True
    End If

    ' In case of copy errors
    On Error Resume Next
      ' Copy VBS to Temp
      fso.CopyFile strPath & "\" & WScript.ScriptName, strTEMPFolderPath & "\" & strRandomName & ".vbs", True
      ' Copy INI to Temp
      fso.CopyFile strINIFilePath, strTEMPFolderPath & "\" & strRandomName & ".ini", True
      
      If Err.Number <> 0 Then
        boolCritical = True
      End If
    On Error Goto 0

    ' Errors?
    If boolCritical = True Then
      Logger("Please copy this script localy as something went wrong")
      Logger("by copying or creating a temporary name for this script.")
      Logger("... Exiting.")
      Call Exiter
    Else
      ' Launch from local location
      Call ExecuteElevated(strTEMPFolderPath & "\" & strRandomName & ".vbs")
      End If
      
  End If ' End if network location


  ' If not cscript
    If Not LCase(Right(WScript.FullName, 12)) = "\cscript.exe" Then
      Call ExecuteElevated(WScript.ScriptFullName)
    End If
    
End Sub


' Writes Registry Values
Function RegistryWrite(strRegType, strRootKey, strKey, strValueName, strValue)
  
  If Not strRegType = "KEY" Then
    If Len(strValueName) = 0 Then
      Logger("Setting empty registry default at """ & strKey & """")  
    Else
      Logger("Setting registry value for """ & strValueName & """ to """ & strValue & """")
    End If
  End If
  
  Set objContext = CreateObject("WbemScripting.SWbemNamedValueSet")
  objContext.Add "__ProviderArchitecture", intOSArch
  
  Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
  Set objRegistry = objLocator.ConnectServer("", "root\default", "", "", , , , objContext).Get("StdRegProv")
  
  If Ucase(strRegType) = "KEY" Then
    intReturnCode = objRegistry.CreateKey(strRootKey, strKey)
  End If
  
  If Ucase(strRegType) = "REG_DWORD" Then
    intReturnCode = objRegistry.SetDWORDValue(strRootKey, strKey, strValueName, strValue)
  End If
  
  If Ucase(strRegType) = "REG_SZ" Then
    intReturnCode = objRegistry.SetStringValue(strRootKey, strKey, strValueName, strValue)
  End If
  
  If Ucase(strRegType) = "REG_EXPAND_SZ" Then
    intReturnCode = objRegistry.SetExpandedStringValue(strRootKey, strKey, strValueName, strValue)
  End If
  
  ' Failed
  If Not intReturnCode = 0 Then
    Logger("Error setting key """ & strKey & "\" & strValueName & """ Error: " & intReturnCode)
    RegistryWrite = False
  Else
    'Passed
    RegistryWrite = True
  End If
End Function


'Reads different types of registry Values
Function RegistryRead(strRootkey, strKey, strValueName, strRegType)
    ' Where strRegType is:
    ' GetStringValue
    ' GetDWORDValue
    ' GetExpandedStringValue 
    ' GetQWORDValue  
    ' GetBinaryValue   
  
  ' TODO: make objects global
    Dim oInParams, oOutParams

    Set objContext = CreateObject("WbemScripting.SWbemNamedValueSet")
    objContext.Add "__ProviderArchitecture", intOSArch

    Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
    Set objRegistry = objLocator.ConnectServer("", "root\default", "", "", , , , objContext).Get("StdRegProv")

    Set oInParams = objRegistry.Methods_(strRegType).InParameters
    oInParams.hDefKey = strRootkey
    oInParams.sSubKeyName = strKey
    oInParams.sValueName = strValueName

    Set oOutParams = objRegistry.ExecMethod_(strRegType, oInParams, , objContext)
    
  ' Return Values
    If strRegType = "GetDWORDValue" Or strRegType = "GetQWORDValue" Or strRegType = "GetBinaryValue" Then
    RegistryRead = oOutParams.uValue
    Else
    RegistryRead = oOutParams.sValue
  End If
  
End Function


'Deletes Registry Values
Function RegistryDelete(strRootkey, strKey, strValueName)
  
  Logger("Deleting Registry value for """ & strValueName & """")
    
  Set objContext = CreateObject("WbemScripting.SWbemNamedValueSet")
  objContext.Add "__ProviderArchitecture", intOSArch
  
  Set objLocator = CreateObject("WbemScripting.SWbemLocator")
  Set objRegistry = objLocator.ConnectServer("", "root\default", "", "", , , , objContext).Get("StdRegProv")
  
  intReturn = objRegistry.deletevalue(strRootkey, strKey, strValueName)
  
  If (intReturn <> 0) Or (Err.Number <> 0) Then
    Logger("Error deleting key " & strKey & " Value Name " & strValueName & " Return code: " & CStr(intReturn) & " Error Description: " & Err.Description)
    RegistryDelete = False
  Else
    RegistryDelete = True
  End If
End Function
    

' Creates a random 8 char name
Function RandomName()
  strRandomName = ""
  
  Randomize()
  
  For i = 1 To 8
    strChar = Mid(Chars, Fix(26 * Rnd())+1, 1)
    strRandomName = strRandomName & strChar
  Next

  RandomName = strRandomName
End Function


' Exit if INI file missing
Sub CheckINIFileExists()
  If Not fso.FileExists(strINIFilePath) Then
    Logger("Sorry, no INI file found at """ & strINIFilePath & """...")
    Logger("... Exiting.")
    Call Exiter
  End If
End Sub


Sub Welcome()
  Logger(vbCrLf)
  Logger("+-----------------------+")
  Logger("¦       Welcome!        ¦")
  Logger("+-----------------------+")
  Logger(vbCrLf)
End Sub


' Do we have Kekse?
Sub GetPermissions()
  ' Try to read
  On Error Resume Next
    ws.RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
    intErr = Err.Number
  On Error Goto 0
    
  If intErr = 0 Then
    boolAdmin = True
  End If
End Sub


' I want Kekse!
Sub ExecuteElevated(strScript)
  os.ShellExecute "cscript.exe", "//nologo """ & strScript & """", , "runas", 1
  WScript.Quit
End Sub


' Some vars...
Sub Config()
  boolUserInteraction = True
  intOSArch = GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth
  strTEMPFolderPath = ws.ExpandEnvironmentStrings("%TEMP%")
  strINIFileName = CreateScriptINIFileName()
  strINIFilePath = strPath & "\" & strINIFileName
  
  dictAcceptedRegistryTypes.Add "REG_SZ", ""
  dictAcceptedRegistryTypes.Add "REG_DWORD", ""
  dictAcceptedRegistryTypes.Add "EXPAND_SZ", ""
  dictAcceptedRegistryTypes.Add "KEY", ""
 
End Sub


' Main Routine
Sub INIT()

  ' Set boolAdmin if administrative permissions
  Call GetPermissions
  
  ' Sets script path
  Call GetScriptLocation
  
  ' Some variables, settings and checks
  Call Config
  
  ' If tool has been started from network location (Hello Admins)
  ' then this one will copy the script to users local temp folder
  ' and restart from that location. Check code for details if interested
  Call CheckLocalStartup
    
  ' No squishy things. If the update service has an pending reboot, then
  ' these updates need to be configured completely by performing an
  ' reboot. Yeah, well, Windows. After reboot this script may be started again
  Call UpdateServiceRequiresReboot
  
  ' Show Welcome Info
  Call Welcome
  
  ' Read Updates section, key=value.
  ' Hotfix with or without "KB" doesn't matter
  Call ReadINIFile("Updates")
  
  ' Create dictionary of tasks to handle from INI file
  Call ReadINIFile("Tasks")
  
  ' Format: service=mode.
  ' Modes: disable|delete. Any other mode as stated is beeing ignored
  ' Please note, that for example "remoteregistry" in an company environment
  ' is widely used by antivirus programs, so you might want to skip this
  ' service and not disable it.
  Call ReadINIFile("Services")
  
  ' Creates a dictionary of installed updates.
  ' Much more faster than firing up an uninstall of all updates which may not even
  ' be installed. Only installed spying updates will be removed
  Call InstalledUpdates

  ' For each spying update in installed updates dictionary: remove
  Call RemoveUpdates
  
  ' Foreach task in INI file, disable
  Call DisableTasks
  
  ' Handle services as stated in INI file.
  ' Only disable|delete modes will be processed
  Call HandleServices

  ' Write or delete Registry entry
  Call HandleRegistry
    
  ' Some users disable windows update completely. In order to flag updates as "hidden",
  ' the Windows Update service needs to be available to search them. If the service is
  ' deactivated, it will enable the service
  Call CheckUpdateService
  
  ' Uses Windows Update service to scan all missing
  ' updates and flag the spying ones as hidden
  Call HideUpdates
  
  ' If Windows Update service has been disabled by the user before,
  ' revert his settings back to disabled.
  Call RevertUpdateService
  
  ' Exit informations
  Call ExitInfo
  
End Sub

' Let's go!
Call INIT()
