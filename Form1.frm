VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JPT User Manager"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   11205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   360
      Left            =   60
      TabIndex        =   12
      Top             =   840
      Width           =   990
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   360
      Left            =   10140
      TabIndex        =   10
      Top             =   900
      Width           =   990
   End
   Begin VB.TextBox txtEmail 
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      Top             =   360
      Width           =   2595
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   480
      Left            =   4380
      TabIndex        =   6
      Top             =   900
      Width           =   1530
   End
   Begin VB.CheckBox chkReport 
      Caption         =   "Gets Report?"
      Height          =   255
      Left            =   8160
      TabIndex        =   5
      Top             =   420
      Width           =   1515
   End
   Begin VB.CheckBox chkIsAdmin 
      Caption         =   "Admin?"
      Height          =   255
      Left            =   8160
      TabIndex        =   4
      Top             =   180
      Width           =   915
   End
   Begin VB.TextBox txtFullname 
      Height          =   315
      Left            =   3180
      TabIndex        =   2
      Top             =   360
      Width           =   1875
   End
   Begin VB.TextBox txtUserName 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexUserList 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7435
      _Version        =   393216
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coded By: Bobby Lovell"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Left            =   0
      TabIndex        =   11
      Top             =   1320
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email (@Worthingtonindustries.com):"
      Height          =   195
      Left            =   5280
      TabIndex        =   9
      Top             =   120
      Width           =   2670
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name:"
      Height          =   195
      Left            =   3180
      TabIndex        =   8
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   195
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   780
   End
   Begin VB.Menu mnuActions 
      Caption         =   "Actions"
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete User"
      End
      Begin VB.Menu mnuNewUser 
         Caption         =   "New User"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub GetUser(strUsername As String)
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM ticketdb.users where idUsers = '" & Trim$(strUsername) & "'  ORDER BY idFullName"
    Set rs = cn_global.Execute(strSQL1)
    With rs
        txtUserName.Text = !idUsers
        txtFullname.Text = !idFullname
        txtEmail.Text = !idEmail
        chkIsAdmin.Value = !idAdmins
        chkReport.Value = !idJPTReport
    End With
    Set rs = Nothing
    cmdUpdate.Enabled = True
    txtUserName.Enabled = False
End Sub
Private Sub ClearAll()
    txtUserName.Text = ""
    txtFullname.Text = ""
    txtEmail.Text = ""
    chkIsAdmin.Value = False
    chkReport.Value = False
    cmdUpdate.Enabled = False
    txtUserName.Enabled = True
End Sub
Private Sub cmdClear_Click()
    ClearAll
End Sub

Private Sub cmdCheck_Click()
    If IsInAD(LCase$(txtUserName.Text)) Then
        MsgBox "User is in directory!"
    Else
        MsgBox "User is NOT in directory!"
    End If
End Sub
Public Sub GetUsers()
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM ticketdb.users ORDER BY idFullName"
    Set rs = cn_global.Execute(strSQL1)
    flexUserList.Clear
    flexUserList.Rows = 2
    flexUserList.FixedCols = 1
    flexUserList.FixedRows = 1
    flexUserList.Rows = rs.RecordCount + 1
    flexUserList.Cols = 8
    flexUserList.TextMatrix(0, 1) = "Full Name"
    flexUserList.TextMatrix(0, 2) = "Username"
    flexUserList.TextMatrix(0, 3) = "Email"
    flexUserList.TextMatrix(0, 4) = "Is Admin"
    flexUserList.TextMatrix(0, 5) = "Gets Report"
    flexUserList.TextMatrix(0, 6) = "Last Log-In"
    flexUserList.TextMatrix(0, 7) = "Log-Ins"
    Do Until rs.EOF
        With rs
            flexUserList.TextMatrix(.AbsolutePosition, 1) = !idFullname
            flexUserList.TextMatrix(.AbsolutePosition, 2) = !idUsers
            flexUserList.TextMatrix(.AbsolutePosition, 3) = !idEmail
            flexUserList.TextMatrix(.AbsolutePosition, 4) = CBool(!idAdmins)
            flexUserList.TextMatrix(.AbsolutePosition, 5) = CBool(!idJPTReport)
            flexUserList.TextMatrix(.AbsolutePosition, 6) = IIf(IsNull(!idLastLogIn), "Never", !idLastLogIn)
            flexUserList.TextMatrix(.AbsolutePosition, 7) = !idLogIns
            .MoveNext
        End With
    Loop
    SizeTheSheet flexUserList
End Sub
Public Sub SizeTheSheet(TargetGrid As MSHFlexGrid)
    On Error Resume Next
    Dim z, Y As Integer
    z = 1
    Y = 600
    TargetGrid.ScrollBars = flexScrollBarNone
    Dim col(), i, b As Integer
    ReDim col(TargetGrid.Cols)
    For i = 0 To TargetGrid.Rows - 1
        For b = 0 To TargetGrid.Cols - 1
            If TextWidth(TargetGrid.TextMatrix(i, b)) > col(b) Then col(b) = TextWidth(TargetGrid.TextMatrix(i, b))
        Next b
    Next i
    For b = 0 To TargetGrid.Cols - 1
        If b = 4 Then
            TargetGrid.ColWidth(b) = (col(b) * z) + Y
        Else
            TargetGrid.ColWidth(b) = (col(b) * z) + Y
        End If
        TargetGrid.ColAlignment(b) = flexAlignLeftCenter
    Next b
    TargetGrid.ScrollBars = flexScrollBarBoth
    TargetGrid.ColWidth(0) = 0
End Sub

Private Sub cmdRefresh_Click()
GetUsers
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo errs
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    Dim blah
    Dim Splitstr() As String
    strSQL1 = "SELECT * From users Where idUsers = '" & txtUserName.Text & "'"
    cn_global.CursorLocation = adUseClient
    If Trim$(txtUserName.Text) = "" Or Trim$(txtFullname.Text) = "" Or Trim$(txtEmail.Text) = "" Then
        blah = MsgBox("One or more fields is blank. Please fill all fields.", vbOKOnly + vbInformation, "Something's missing...")
        Exit Sub
    End If
    Splitstr = Split(txtEmail.Text, "@")
    rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
    With rs
        !idFullname = Trim$(txtFullname.Text)
        !idEmail = Trim$(Splitstr(0) & EmailDomain)
        !idAdmins = chkIsAdmin.Value
        !idJPTReport = chkReport.Value
        .Update
    End With
    Set rs = Nothing
    ClearAll
    GetUsers
    blah = MsgBox("User Updated!", vbOKOnly + vbInformation, "Success")
    Exit Sub
errs:
    If Err.Number = -2147217864 Then
        blah = MsgBox("Nothing to update.", vbOKOnly + vbExclamation, "No Changes...")
        GetUser txtUserName.Text
    End If
End Sub

Private Sub flexUserList_DblClick()
    GetUser flexUserList.TextMatrix(flexUserList.RowSel, 2)
End Sub

Private Sub Form_Load()
    FindMySQLDriver
    cn_global.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    GetUsers
End Sub
Public Sub DeleteUser(strUsername As String)
    Dim blah
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * From users Where idUsers = '" & strUsername & "'"
    rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
    With rs
        .Delete
        .Update
    End With
    ClearAll
    GetUsers
    blah = MsgBox(UCase$(strUsername) & " has been deleted.", vbOKOnly + vbInformation, "Success")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    cn_global.Close
    Unload Me
    End
End Sub

Private Sub mnuDelete_Click()
    Dim blah
    If txtUserName.Text <> "" Then
        blah = MsgBox("Are you sure you want to delete user: " & UCase$(txtUserName.Text) & "?", vbOKCancel + vbCritical, "Delete User")
        If blah = vbOK Then
            DeleteUser txtUserName.Text
        Else
            ClearAll
            Exit Sub
        End If
    Else
        blah = MsgBox("Please select a user first.", vbOKOnly + vbExclamation, "Nothing Selected")
    End If
End Sub

Private Sub mnuNewUser_Click()
    frmWait.Show
    DoEvents
    'GetUsersInfo
    GetUsersInfo2
    frmWait.Hide
    frmAddNew.Show
End Sub
