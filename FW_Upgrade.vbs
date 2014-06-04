Dim IPAddr, oX, oX2, cmdTmp1, cmdTmp2, IsStart, DelayTime

Dim FileName
Dim FromPath
Dim PortNo

Function GetExt(fileName)
    Dim pos, m_name, xtn
    pos    = InstrRev(fileName,".")
    xtn    = Mid(fileName,pos+1)
    m_name = Mid(fileName,1,pos-1)
    GetExt = xtn
End Function

Function GetLastPKGFile( path )
    Dim fso, file, recentDate, recentFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set recentFile = Nothing
    For Each file in fso.GetFolder( path ).Files
        if ( "pkg" = GetExt(file) ) then
            If (recentFile is Nothing) Then
                Set recentFile = file
            ElseIf (file.DateLastModified > recentFile.DateLastModified) Then
                Set recentFile = file
            End If
        End If
    Next
    GetLastPKGFile = recentFile
End Function

Class vbsFileUpload
    Public c_strDestURL
    Public c_strFileName
    Public c_strFieldName
    Public c_strBoundary
    Public c_strContentType ' text/plain or image/pjpeg and so on "application/upload"
    Public c_strResponseText
    Public c_boolPrepared
    Public c_strErrMsg

    Public Sub Class_Initialize()
        c_strDestURL     = ""
        c_strFileName    = ""
        c_strContentType = "application/upload"
        c_strFieldName   = "file"
        c_strBoundary    = "---------------------------7da1c52160186"
        c_boolPrepared   = false
    End Sub

    Public Sub Class_Terminate
    End Sub

    Public Function vbsUpload
        CheckRequirements()
        If  c_boolPrepared Then
            UploadFile c_strDestURL, c_strFileName, c_strFieldName
        Else
            WScript.Echo c_strErrMsg
        End If
    End Function

    Private Function CheckRequirements
        Dim objFSO
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        If Not objFSO.FileExists(c_strFileName) Then
            c_strErrMsg = c_strErrMsg & vbCrLf & "ÇlÇïÆé”Z”I.."
            MsgBox c_strFileName + " is not exit!! " + c_strErrMsg
        Else
            On Error Resume Next
            CreateObject "MSXML2.XMLHTTP"
            If Not Err = 0 Then
                c_strErrMsg = c_strErrMsg & vbCrLf & Err.Descriptiof
            Else
                c_boolPrepared = True
            End If
        End If
    End Function

    Private Function UploadFile(DestURL, FileName, FieldName)
        Dim FileContents, FormData,Boundary
        Boundary     = c_strBoundary
        FileContents = GetFile(FileName)
        FormData     = BuildFormData(FileContents, Boundary, FileName, FieldName)
        WinHTTPPostRequest DestURL, FormData, Boundary
    End Function

    Private Function WinHTTPPostRequest(URL, FormData, Boundary)
        Dim xmlhttp
        Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
        On Error Resume Next
        xmlhttp.Open "POST", URL, false
        xmlhttp.setRequestHeader "Accept","text/html, application/xhtml+xml, */*"
        xmlhttp.setRequestHeader "Pragma", "no-cache"
        xmlhttp.setRequestHeader "Content-Type", "multipart/form-data; boundary=" + Boundary
        xmlhttp.setRequestHeader "Authorization","Basic cm9vdDpyb290"
        xmlhttp.setRequestHeader "DNT", "1"
        xmlhttp.send FormData
        c_strResponseText = xmlhttp.responseText
        Set xmlhttp = Nothing
    End Function

    Private Function BuildFormData(FileContents, Boundary, FileName, FieldName)
        Dim FormData, Pre, Po, ContentType
        ContentType = c_strContentType

        'The two parts around file contents In the multipart-form data.
        Pre = "--" + Boundary + vbCrLf + mpFields(FieldName, FileName, ContentType)
        Po = vbCrLf + "--" + Boundary + "--" + vbCrLf

        'Build form data using recordset binary field
        Const adLongVarBinary = 205
        Dim RS: Set RS = CreateObject("ADODB.Recordset")
        RS.Fields.Append "b", adLongVarBinary, Len(Pre) + LenB(FileContents) + Len(Po)
        RS.Open
        RS.AddNew
        Dim LenData
        'Convert Pre string value To a binary data
        LenData = Len(Pre)
        RS("b").AppendChunk (StringToMB(Pre) & ChrB(0))
        Pre = RS("b").GetChunk(LenData)
        RS("b") = ""

        'Convert Po string value To a binary data
        LenData = Len(Po)
        RS("b").AppendChunk (StringToMB(Po) & ChrB(0))
        Po = RS("b").GetChunk(LenData)
        RS("b") = ""

        'Join Pre + FileContents + Po binary data
        RS("b").AppendChunk (Pre)
        RS("b").AppendChunk (FileContents)
        RS("b").AppendChunk (Po)
        RS.Update
        FormData = RS("b")
        RS.Close
        BuildFormData = FormData
    End Function

    'Converts OLE string To multibyte string
    Private Function StringToMB(S)
        Dim I, B
        For I = 1 To Len(S)
            B = B & ChrB(Asc(Mid(S, I, 1)))
        Next
        StringToMB = B
    End Function

    Private Function mpFields(FieldName, FileName, ContentType)
        Dim MPTemplate 'template For multipart header
        MPTemplate = "Content-Disposition: form-data; name=""{field}"";" + _
        " filename=""{file}""" + vbCrLf + _
        "Content-Type: {ct}" + vbCrLf + vbCrLf
        Dim Out
        Out = Replace(MPTemplate, "{field}", FieldName)
        ''Out = Replace(Out, "{file}", FileName)
        Out = Replace(Out, "{file}", me.c_strFieldName)
        mpFields = Replace(Out, "{ct}", ContentType)
    End Function

    Private Function GetFile(FileName)
        Dim Stream: Set Stream = CreateObject("ADODB.Stream")
        Stream.Type = 1 'Binary
        Stream.Open
        Stream.LoadFromFile FileName
        GetFile = Stream.Read
        Stream.Close
    End Function
End Class

sub Main
    if WScript.Arguments.Count = 0 Then
        IPAddr = "http://" + "192.168.20.49"
        PortNo = 80
        FileName = "N8072_V1.09_STD-1_20140604-101509.pkg"
        FromPath = "X:\hisi3511\release\"
    else
        IPAddr = "http://" + Wscript.Arguments.Item(0)
        PortNo = Wscript.Arguments.Item(1)
        FromPath = Wscript.Arguments.Item(2)
        if ( 4=WScript.Arguments.Count ) then
            FileName = Wscript.Arguments.Item(3)
        else
            FileName = replace(GetLastPKGFile(FromPath),FromPath,"")
        end if
    end If
    IPAddr = IPAddr + ":" + CStr(PortNo)

    Dim myUpload
    Set myUpload = New vbsFileUpload
    myUpload.c_strDestURL     = IPAddr + "/firmwareupgrade.cgi"
    myUpload.c_strFileName    = FromPath + FileName
    myUpload.c_strFieldName   = FileName
    myUpload.c_strContentType = "application/octet-stream"
    myUpload.vbsUpload()
    WScript.Echo myUpload.c_strResponseText
end sub

Main

