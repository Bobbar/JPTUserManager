Attribute VB_Name = "MyMod"
Option Explicit
Public Const strServerAddress As String = "ohbre-pwadmin01"
Public Const strUsername      As String = "TicketApp"
Public Const strPassword      As String = "yb4w4"
Public strSQLDriver           As String
Global cn_global              As New ADODB.Connection
Const HKEY_LOCAL_MACHINE = &H80000002
Private Declare Function RegOpenKeyEx _
                Lib "advapi32.dll" _
                Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       ByVal ulOptions As Long, _
                                       ByVal samDesired As Long, _
                                       phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumValue _
                Lib "advapi32.dll" _
                Alias "RegEnumValueA" (ByVal hKey As Long, _
                                       ByVal dwIndex As Long, _
                                       ByVal lpValueName As String, _
                                       lpcbValueName As Long, _
                                       ByVal lpReserved As Long, _
                                       lpType As Long, _
                                       lpData As Any, _
                                       lpcbData As Long) As Long
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (dest As Any, _
                                       Source As Any, _
                                       ByVal numBytes As Long)
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_MULTI_SZ = 7
Const ERROR_MORE_DATA = 234
Const KEY_READ = &H20019 ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
Public Const EmailDomain As String = "@worthingtonindustries.com"
Public Const group       As String = "Domain Users"
Public Const domain      As String = "corp.worthingtonindustries.local"
Public Type UserInfo
    Fullname As String
    Username As String
    Email As String
End Type
Public UserData()     As UserInfo
Public intSelReportID As Integer
Public Function NextGroup() As Integer
    GetReportGroups
    NextGroup = Form1.cmbReportGroup.ListCount
End Function
Public Sub GetReportGroups()
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT DISTINCT (reportsgroups_0.idGroupID) AS 'idGroupID' FROM ticketdb.reportsgroups reportsgroups_0 ORDER BY reportsgroups_0.idGroupID"
    Set rs = cn_global.Execute(strSQL1)
    Form1.cmbReportGroup.Clear
    Form1.cmbReportGroups.Clear
    Form1.cmbReportGroup.AddItem "0"
    Form1.cmbReportGroups.AddItem "0"
    Do Until rs.EOF
        With rs
            Form1.cmbReportGroup.AddItem !idGroupID '.Fields(.AbsolutePosition)
            Form1.cmbReportGroups.AddItem !idGroupID
            .MoveNext
        End With
    Loop
End Sub
Public Function FilterString() As String
    Dim tmpString As String
    With Form1
        If .chkControls.Value = 1 Then
            tmpString = tmpString + "1,"
        Else
            tmpString = tmpString + "0,"
        End If
        If .chkIM.Value = 1 Then
            tmpString = tmpString + "1,"
        Else
            tmpString = tmpString + "0,"
        End If
        If .chkNuclear.Value = 1 Then
            tmpString = tmpString + "1,"
        Else
            tmpString = tmpString + "0,"
        End If
        If .chkRockyMtn.Value = 1 Then
            tmpString = tmpString + "1,"
        Else
            tmpString = tmpString + "0,"
        End If
        If .chkSteelFab.Value = 1 Then
            tmpString = tmpString + "1,"
        Else
            tmpString = tmpString + "0,"
        End If
        If .chkWooster.Value = 1 Then
            tmpString = tmpString + "1"
        Else
            tmpString = tmpString + "0"
        End If
    End With
    FilterString = tmpString
End Function
Public Function FindUserNameExact(strUsername As String) As UserInfo
    Dim i As Integer
    FindUserNameExact.Email = ""
    FindUserNameExact.Fullname = ""
    FindUserNameExact.Username = ""
    For i = 0 To UBound(UserData)
        If Trim$(UCase$(strUsername)) = Trim$(UCase$(UserData(i).Username)) Then
            FindUserNameExact.Email = UserData(i).Email
            FindUserNameExact.Fullname = UserData(i).Fullname
            FindUserNameExact.Username = UserData(i).Username
        End If
    Next
End Function
Public Function FindUserNamePartial(strPartialUsername As String) As UserInfo
    Dim i As Integer
    FindUserNamePartial.Email = ""
    FindUserNamePartial.Fullname = ""
    FindUserNamePartial.Username = ""
    For i = 0 To UBound(UserData)
        If InStr(1, UCase$(UserData(i).Username), UCase$(strPartialUsername), vbTextCompare) <> 0 Then
            FindUserNamePartial.Email = UserData(i).Email
            FindUserNamePartial.Fullname = UserData(i).Fullname
            FindUserNamePartial.Username = UserData(i).Username
        End If
    Next
End Function
Public Function FindFullName(strPartialFullName As String) As UserInfo
    Dim i As Integer
    FindFullName.Email = ""
    FindFullName.Fullname = ""
    FindFullName.Username = ""
    For i = 0 To UBound(UserData)
        If InStr(1, UCase$(UserData(i).Fullname), UCase$(strPartialFullName), vbTextCompare) <> 0 Then
            FindFullName.Email = UserData(i).Email
            FindFullName.Fullname = UserData(i).Fullname
            FindFullName.Username = UserData(i).Username
        End If
    Next
End Function
Public Function FindEmail(strPartialEmail As String) As UserInfo
    Dim i As Integer
    FindEmail.Email = ""
    FindEmail.Fullname = ""
    FindEmail.Username = ""
    For i = 0 To UBound(UserData)
        If InStr(1, UCase$(UserData(i).Email), UCase$(strPartialEmail), vbTextCompare) <> 0 Then
            FindEmail.Email = UserData(i).Email
            FindEmail.Fullname = UserData(i).Fullname
            FindEmail.Username = UserData(i).Username
        End If
    Next
End Function
Public Sub FindMySQLDriver()
    GetODBCDrivers
    Dim i           As Integer
    Dim strPossis() As String
    Dim blah
    ReDim strPossis(0)
    For i = 1 To GetODBCDrivers.Count
        If InStr(1, GetODBCDrivers.Item(i), "MySQL") Then
            strPossis(UBound(strPossis)) = GetODBCDrivers.Item(i)
            ReDim Preserve strPossis(UBound(strPossis) + 1)
        End If
    Next i
    If UBound(strPossis) > 1 Then
        blah = MsgBox("Multiple MySQL Drivers detected!", vbExclamation + vbOKOnly, "Gasp!")
        strSQLDriver = strPossis(0)
    Else
        strSQLDriver = strPossis(0)
    End If
End Sub
Function GetODBCDrivers() As Collection
    Dim res    As Collection
    Dim values As Variant
    ' initialize the result
    Set GetODBCDrivers = New Collection
    ' the names of all the ODBC drivers are kept as values
    ' under a registry key
    ' the EnumRegistryValue returns a collection
    For Each values In EnumRegistryValues(HKEY_LOCAL_MACHINE, "Software\ODBC\ODBCINST.INI\ODBC Drivers")
        ' each element is a two-item array:
        ' values(0) is the name, values(1) is the data
        If StrComp(values(1), "Installed", 1) = 0 Then
            ' if installed, add to the result collection
            GetODBCDrivers.Add values(0), values(0)
        End If
    Next
End Function
Public Function BoltoInt(strBool As String) As Integer
    If UCase$(Trim$(strBool)) = "TRUE" Then
        BoltoInt = 1
    ElseIf UCase$(Trim$(strBool)) = "FALSE" Then
        BoltoInt = 0
    End If
End Function
Function EnumRegistryValues(ByVal hKey As Long, ByVal KeyName As String) As Collection
    Dim handle            As Long
    Dim Index             As Long
    Dim valueType         As Long
    Dim Name              As String
    Dim nameLen           As Long
    Dim resLong           As Long
    Dim resString         As String
    Dim dataLen           As Long
    Dim valueInfo(0 To 1) As Variant
    Dim retVal            As Long
    ' initialize the result
    Set EnumRegistryValues = New Collection
    ' Open the key, exit if not found.
    If Len(KeyName) Then
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
        ' in all cases, subsequent functions use hKey
        hKey = handle
    End If
    Do
        ' this is the max length for a key name
        nameLen = 260
        Name = Space$(nameLen)
        ' prepare the receiving buffer for the value
        dataLen = 4096
        ReDim resBinary(0 To dataLen - 1) As Byte
        ' read the value's name and data
        ' exit the loop if not found
        retVal = RegEnumValue(hKey, Index, Name, nameLen, ByVal 0&, valueType, resBinary(0), dataLen)
        ' enlarge the buffer if you need more space
        If retVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To dataLen - 1) As Byte
            retVal = RegEnumValue(hKey, Index, Name, nameLen, ByVal 0&, valueType, resBinary(0), dataLen)
        End If
        ' exit the loop if any other error (typically, no more values)
        If retVal Then Exit Do
        ' retrieve the value's name
        valueInfo(0) = Left$(Name, nameLen)
        ' return a value corresponding to the value type
        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                valueInfo(1) = resLong
            Case REG_SZ, REG_EXPAND_SZ
                ' copy everything but the trailing null char
                resString = Space$(dataLen - 1)
                CopyMemory ByVal resString, resBinary(0), dataLen - 1
                valueInfo(1) = resString
            Case REG_BINARY
                ' shrink the buffer if necessary
                If dataLen < UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To dataLen - 1) As Byte
                End If
                valueInfo(1) = resBinary()
            Case REG_MULTI_SZ
                ' copy everything but the 2 trailing null chars
                resString = Space$(dataLen - 2)
                CopyMemory ByVal resString, resBinary(0), dataLen - 2
                valueInfo(1) = resString
            Case Else
                ' Unsupported value type - do nothing
        End Select
        ' add the array to the result collection
        ' the element's key is the value's name
        EnumRegistryValues.Add valueInfo, valueInfo(0)
        Index = Index + 1
    Loop
    ' Close the key, if it was actually opened
    If handle Then RegCloseKey handle
End Function
Public Function IsInAD(Username As String) As Boolean
    IsInAD = False
    If FindUserNameExact(Username).Username <> "" Then IsInAD = True
End Function
Public Function GetUsersInfo()
    On Error GoTo errs
    Dim intEntry     As Integer
    Dim user         As IADsUser
    Dim BremanGroup  As IADsContainer
    Dim WoosterGroup As IADsContainer
    Set BremanGroup = GetObject("LDAP://ohbre-pwdc01.corp.worthingtonindustries.local:389/OU=CYL-Bremen-OH-USA-Users,OU=CYL-Bremen-OH-USA,OU=Cylinders,DC=corp,DC=worthingtonindustries,DC=local")
    Set WoosterGroup = GetObject("LDAP://ohbre-pwdc01.corp.worthingtonindustries.local:389/OU=CYL-Wooster-OH-USA-Users,OU=CYL-Wooster-OH-USA,OU=Cylinders,DC=corp,DC=worthingtonindustries,DC=local")
    BremanGroup.Filter = Array("user")
    WoosterGroup.Filter = Array("user")
    intEntry = 0
    ReDim UserData(0)
    For Each user In BremanGroup
        UserData(intEntry).Fullname = user.Fullname
        UserData(intEntry).Username = user.Get("sAMAccountName")
        UserData(intEntry).Email = user.Get("mail")
        intEntry = intEntry + 1
        ReDim Preserve UserData(intEntry)
    Next user
    For Each user In WoosterGroup
        UserData(intEntry).Fullname = user.Fullname
        UserData(intEntry).Username = user.Get("sAMAccountName")
        UserData(intEntry).Email = user.Get("mail")
        intEntry = intEntry + 1
        ReDim Preserve UserData(intEntry)
    Next user
    '    Dim i As Integer
    '    For i = 0 To UBound(UserData)
    '        Debug.Print UserData(i).Username & " - " & UserData(i).Fullname & " - " & UserData(i).Email
    '    Next
errs:
    If Err.Number = -2147463155 Then Resume Next
End Function
Public Function GetUsersInfo2()
    On Error Resume Next
    Dim intEntry        As Integer
    Dim strSubGroupName As String
    Dim CylGroup        As IADsContainer
    Dim StlGroup        As IADsContainer
    Dim SubGroup        As IADsContainer
    Dim SubGroupName    As Object
    Dim user            As IADsUser
    Set CylGroup = GetObject("LDAP://corp.worthingtonindustries.local:389/OU=Cylinders,DC=corp,DC=worthingtonindustries,DC=local")
    Set StlGroup = GetObject("LDAP://corp.worthingtonindustries.local:389/OU=Steel,DC=corp,DC=worthingtonindustries,DC=local")
    intEntry = 0
    ReDim UserData(0)
    For Each SubGroupName In CylGroup
        strSubGroupName = SubGroupName.Name
        frmWait.lblStatus.Caption = strSubGroupName
        DoEvents
        Set SubGroup = GetObject("LDAP://corp.worthingtonindustries.local:389/" & strSubGroupName & "-Users," & strSubGroupName & ",OU=Cylinders,DC=corp,DC=worthingtonindustries,DC=local")
        SubGroup.Filter = Array("user")
        For Each user In SubGroup
            '  Debug.Print user.Fullname & " - " & user.Get("mail")
            UserData(intEntry).Fullname = user.Fullname
            UserData(intEntry).Username = user.Get("sAMAccountName")
            UserData(intEntry).Email = user.Get("mail")
            intEntry = intEntry + 1
            ReDim Preserve UserData(intEntry)
        Next user
        Set SubGroup = Nothing
    Next SubGroupName
    For Each SubGroupName In StlGroup
        strSubGroupName = SubGroupName.Name
        frmWait.lblStatus.Caption = strSubGroupName
        DoEvents
        Set SubGroup = GetObject("LDAP://corp.worthingtonindustries.local:389/" & strSubGroupName & "-Users," & strSubGroupName & ",OU=Steel,DC=corp,DC=worthingtonindustries,DC=local")
        SubGroup.Filter = Array("user")
        For Each user In SubGroup
            ' Debug.Print user.Fullname & " - " & user.Get("mail")
            UserData(intEntry).Fullname = user.Fullname
            UserData(intEntry).Username = user.Get("sAMAccountName")
            UserData(intEntry).Email = user.Get("mail")
            intEntry = intEntry + 1
            ReDim Preserve UserData(intEntry)
        Next user
        Set SubGroup = Nothing
    Next SubGroupName
End Function
Public Function IsInDB(Username As String) As Boolean
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM ticketdb.users where idUsers = '" & Trim$(Username) & "'  ORDER BY idFullName"
    Set rs = cn_global.Execute(strSQL1)
    If rs.RecordCount > 0 Then
        IsInDB = True
    ElseIf rs.RecordCount = 0 Then
        IsInDB = False
    End If
    Set rs = Nothing
End Function
Public Sub AddToUserList(strUsername As String, _
                         strFullname As String, _
                         strEmail As String, _
                         intAdmin As Integer, _
                         intReport As Integer)
    Dim blah
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    strSQL1 = "INSERT INTO users (idFullName,idUsers,idEmail,idAdmins,idJPTReport)" & " VALUES ('" & strFullname & "','" & strUsername & "','" & strEmail & "','" & intAdmin & "','" & intReport & "')"
    Set rs = cn_global.Execute(strSQL1)
    Set rs = Nothing
    blah = MsgBox("User added to database!", vbInformation + vbOKOnly, "Success")
End Sub
