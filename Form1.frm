VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JPT User Manager"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   13350
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
   ScaleHeight     =   7005
   ScaleWidth      =   13350
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   13155
      _ExtentX        =   23204
      _ExtentY        =   12091
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Users"
      TabPicture(0)   =   "Form1.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "flexUserList"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdRefresh"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdClear"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtEmail"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdUpdate"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkIsAdmin"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtFullname"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtUserName"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmbReportGroup"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Reports"
      TabPicture(1)   =   "Form1.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkHasRun"
      Tab(1).Control(1)=   "cmdClearReport"
      Tab(1).Control(2)=   "cmdUpdateReport"
      Tab(1).Control(3)=   "chkNewReport"
      Tab(1).Control(4)=   "txtEndDate"
      Tab(1).Control(5)=   "txtStartDate"
      Tab(1).Control(6)=   "txtRunHour"
      Tab(1).Control(7)=   "cmbRunDay"
      Tab(1).Control(8)=   "txtReportName"
      Tab(1).Control(9)=   "chkSteelFab"
      Tab(1).Control(10)=   "chkControls"
      Tab(1).Control(11)=   "chkIM"
      Tab(1).Control(12)=   "chkNuclear"
      Tab(1).Control(13)=   "chkWooster"
      Tab(1).Control(14)=   "chkRockyMtn"
      Tab(1).Control(15)=   "flexReports"
      Tab(1).Control(16)=   "Label10"
      Tab(1).Control(17)=   "Label9"
      Tab(1).Control(18)=   "Label8"
      Tab(1).Control(19)=   "Label7"
      Tab(1).Control(20)=   "Label6"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Report Groups"
      TabPicture(2)   =   "Form1.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label11"
      Tab(2).Control(1)=   "Label12"
      Tab(2).Control(2)=   "flexGroups"
      Tab(2).Control(3)=   "cmbReportGroups"
      Tab(2).Control(4)=   "cmbReports"
      Tab(2).Control(5)=   "cmdNewGroup"
      Tab(2).Control(6)=   "cmdDeleteGroup"
      Tab(2).Control(7)=   "cmdRemoveReport"
      Tab(2).ControlCount=   8
      Begin VB.CommandButton cmdRemoveReport 
         Caption         =   "Remove"
         Height          =   300
         Left            =   -63180
         TabIndex        =   43
         Top             =   1500
         Width           =   1170
      End
      Begin VB.CheckBox chkHasRun 
         Caption         =   "Has Run?"
         Height          =   255
         Left            =   -64680
         TabIndex        =   42
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdDeleteGroup 
         Caption         =   "Delete Group"
         Height          =   360
         Left            =   -70500
         TabIndex        =   41
         Top             =   1080
         Width           =   1170
      End
      Begin VB.CommandButton cmdNewGroup 
         Caption         =   "New Group"
         Height          =   360
         Left            =   -70500
         TabIndex        =   40
         Top             =   600
         Width           =   1170
      End
      Begin VB.ComboBox cmbReports 
         Height          =   315
         Left            =   -67500
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1020
         Width           =   2115
      End
      Begin VB.ComboBox cmbReportGroups 
         Height          =   315
         Left            =   -72180
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1020
         Width           =   1575
      End
      Begin VB.CommandButton cmdClearReport 
         Caption         =   "Clear"
         Height          =   360
         Left            =   -63000
         TabIndex        =   34
         Top             =   1380
         Width           =   990
      End
      Begin VB.CommandButton cmdUpdateReport 
         Caption         =   "Update"
         Height          =   540
         Left            =   -69180
         TabIndex        =   33
         Top             =   1140
         Width           =   1530
      End
      Begin VB.CheckBox chkNewReport 
         Caption         =   "New?"
         Height          =   195
         Left            =   -74820
         TabIndex        =   32
         Top             =   780
         Width           =   735
      End
      Begin VB.TextBox txtEndDate 
         Height          =   315
         Left            =   -67080
         TabIndex        =   30
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtStartDate 
         Height          =   315
         Left            =   -68580
         TabIndex        =   28
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtRunHour 
         Height          =   315
         Left            =   -69600
         TabIndex        =   26
         Text            =   "0"
         Top             =   720
         Width           =   795
      End
      Begin VB.ComboBox cmbRunDay 
         Height          =   315
         Left            =   -71280
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtReportName 
         Height          =   315
         Left            =   -74040
         TabIndex        =   21
         Top             =   720
         Width           =   2595
      End
      Begin VB.ComboBox cmbReportGroup 
         Height          =   315
         Left            =   8580
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   720
         Width           =   1995
      End
      Begin VB.CheckBox chkSteelFab 
         Caption         =   "SteelFab"
         Height          =   255
         Left            =   -63660
         TabIndex        =   13
         Top             =   840
         Width           =   1035
      End
      Begin VB.CheckBox chkControls 
         Caption         =   "Controls"
         Height          =   255
         Left            =   -65580
         TabIndex        =   18
         Top             =   840
         Width           =   915
      End
      Begin VB.CheckBox chkIM 
         Caption         =   "IM"
         Height          =   255
         Left            =   -65340
         TabIndex        =   17
         Top             =   600
         Width           =   555
      End
      Begin VB.CheckBox chkNuclear 
         Caption         =   "Nuclear"
         Height          =   255
         Left            =   -64740
         TabIndex        =   16
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chkWooster 
         Caption         =   "Wooster"
         Height          =   255
         Left            =   -63840
         TabIndex        =   15
         Top             =   600
         Width           =   915
      End
      Begin VB.CheckBox chkRockyMtn 
         Caption         =   "RockyMtn"
         Height          =   255
         Left            =   -64680
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtUserName 
         Height          =   315
         Left            =   2220
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtFullname 
         Height          =   315
         Left            =   3600
         TabIndex        =   6
         Top             =   720
         Width           =   1875
      End
      Begin VB.CheckBox chkIsAdmin 
         Caption         =   "Admin?"
         Height          =   255
         Left            =   10860
         TabIndex        =   5
         Top             =   720
         Width           =   915
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   480
         Left            =   5580
         TabIndex        =   4
         Top             =   1260
         Width           =   1530
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Left            =   5700
         TabIndex        =   3
         Top             =   720
         Width           =   2595
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   360
         Left            =   12000
         TabIndex        =   2
         Top             =   1500
         Width           =   990
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   990
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexUserList 
         Height          =   4815
         Left            =   180
         TabIndex        =   8
         Top             =   1920
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   8493
         _Version        =   393216
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexReports 
         Height          =   4755
         Left            =   -74820
         TabIndex        =   24
         Top             =   1920
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   8387
         _Version        =   393216
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexGroups 
         Height          =   4815
         Left            =   -74820
         TabIndex        =   35
         Top             =   1860
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   8493
         _Version        =   393216
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add Reports"
         Height          =   195
         Left            =   -67500
         TabIndex        =   39
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Group"
         Height          =   195
         Left            =   -72180
         TabIndex        =   37
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         Height          =   195
         Left            =   -67080
         TabIndex        =   31
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         Height          =   195
         Left            =   -68580
         TabIndex        =   29
         Top             =   480
         Width           =   1050
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Run Hour"
         Height          =   195
         Left            =   -69540
         TabIndex        =   27
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Run Day"
         Height          =   195
         Left            =   -71220
         TabIndex        =   25
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Name"
         Height          =   195
         Left            =   -74040
         TabIndex        =   22
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Group"
         Height          =   195
         Left            =   8580
         TabIndex        =   20
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         Height          =   195
         Left            =   2220
         TabIndex        =   12
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name:"
         Height          =   195
         Left            =   3600
         TabIndex        =   11
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email (@Worthingtonindustries.com):"
         Height          =   195
         Left            =   5700
         TabIndex        =   10
         Top             =   480
         Width           =   2670
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
         Left            =   60
         TabIndex        =   9
         Top             =   1680
         Width           =   1530
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "Actions"
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete User"
      End
      Begin VB.Menu mnuNewUser 
         Caption         =   "New User"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReports 
         Caption         =   "Reports"
         Begin VB.Menu mnuDeleteReport 
            Caption         =   "Delete Report"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub DeleteReportFromGroup(GroupID As Integer, _
                                 ReportID As Integer, _
                                 EntryID As Integer)
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM reportsgroups WHERE idGroupID = '" & GroupID & "' AND idReportID = '" & ReportID & "' AND idEntryID = '" & EntryID & "'"
    rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
    With rs
        .Delete
    End With
    'rs.Update
End Sub
Public Sub DeleteGroup(GroupID As Integer)
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM reportsgroups WHERE idGroupID = '" & GroupID & "'"
    rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
    Do Until rs.EOF
        With rs
            .Delete
            .MoveNext
        End With
    Loop
    'rs.Update
End Sub
Public Sub GetGroups(GroupID As Integer)
    SetupRunDayCombo
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT reportsgroups_0.idGroupID, reportsgroups_0.idReportID, reportsgroups_0.idEntryID, reports_0.idReportName, reports_0.idReportID, reports_0.idRunDay, reports_0.idRunTime, reports_0.idStartDate, reports_0.idEndDate, reports_0.idCompanyFilter, reports_0.idHasRun" & " FROM ticketdb.reports reports_0, ticketdb.reportsgroups reportsgroups_0" & " WHERE reports_0.idReportID = reportsgroups_0.idReportID AND ((reportsgroups_0.idGroupID=" & GroupID & "))"
    Set rs = cn_global.Execute(strSQL1)
    flexGroups.Clear
    flexGroups.Rows = 2
    flexGroups.FixedCols = 1
    flexGroups.FixedRows = 1
    flexGroups.Rows = rs.RecordCount + 1
    flexGroups.Cols = 9
    flexGroups.TextMatrix(0, 1) = "Report Name"
    flexGroups.TextMatrix(0, 2) = "ReportID"
    flexGroups.TextMatrix(0, 3) = "Run Day"
    flexGroups.TextMatrix(0, 4) = "Run Hour"
    flexGroups.TextMatrix(0, 5) = "Start Date"
    flexGroups.TextMatrix(0, 6) = "End Date"
    flexGroups.TextMatrix(0, 7) = "Company Filter"
    flexGroups.TextMatrix(0, 8) = "Entry ID"
    flexGroups.ColWidth(8) = 0
    Do Until rs.EOF
        With rs
            flexGroups.TextMatrix(.AbsolutePosition, 1) = !idReportName
            flexGroups.TextMatrix(.AbsolutePosition, 2) = !idReportID
            flexGroups.TextMatrix(.AbsolutePosition, 3) = !idRunDay
            flexGroups.TextMatrix(.AbsolutePosition, 4) = !idRunTime
            flexGroups.TextMatrix(.AbsolutePosition, 5) = !idStartDate
            flexGroups.TextMatrix(.AbsolutePosition, 6) = !idEndDate
            flexGroups.TextMatrix(.AbsolutePosition, 7) = !idCompanyFilter
            flexGroups.TextMatrix(.AbsolutePosition, 8) = !idEntryID
            .MoveNext
        End With
    Loop
    SizeTheSheet flexGroups
End Sub
Public Sub GetUser(strUsername As String)
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM ticketdb.users where idUsers = '" & Trim$(strUsername) & "'  ORDER BY idFullName"
    Set rs = cn_global.Execute(strSQL1)
    With rs
        txtUsername.Text = !idUsers
        txtFullName.Text = !idFullname
        txtEmail.Text = !idEmail
        chkIsAdmin.Value = !idAdmins
        '        chkReport.Value = !idJPTReport
        '        chkDailyReport.Value = !idJPTDailyReport
        cmbReportGroup.Text = !idGroupID
        If CInt(!idJPTDailyReport) = 1 Then
            EnableFilters
            SetFilters !idCompanyFilters
        Else
            'DisableFilters
        End If
    End With
    Set rs = Nothing
    cmdUpdate.Enabled = True
    txtUsername.Enabled = False
End Sub
Public Sub GetReport(strReportName As String, intReportID As Integer)
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    Dim RunVal As Integer
    ClearFilters
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM ticketdb.reports where idReportName = '" & Trim$(strReportName) & "' AND idReportID = '" & intReportID & "'"
    Set rs = cn_global.Execute(strSQL1)
    With rs
        txtReportName.Text = !idReportName
        cmbRunDay.ListIndex = CInt(!idRunDay) - 1
        txtRunHour.Text = !idRunTime
        txtStartDate.Text = !idStartDate
        txtEndDate.Text = !idEndDate
        SetFilters !idCompanyFilter
        intSelReportID = CInt(!idReportID)
        RunVal = !idHasRun
        chkHasRun.Value = IIf(RunVal = -1, 1, 0)
        
    End With
    Set rs = Nothing
    cmdUpdateReport.Enabled = True
    EnableFilters
End Sub
Public Sub SetFilters(Filters As String)
    Dim FilterArray() As String
    FilterArray = Split(Filters, ",")
    If FilterArray(0) = 1 Then chkControls.Value = 1
    If FilterArray(1) = 1 Then chkIM.Value = 1
    If FilterArray(2) = 1 Then chkNuclear.Value = 1
    If FilterArray(3) = 1 Then chkRockyMtn.Value = 1
    If FilterArray(4) = 1 Then chkSteelFab.Value = 1
    If FilterArray(5) = 1 Then chkWooster.Value = 1
End Sub
Public Sub DisableFilters()
    ClearFilters
    chkControls.Enabled = False
    chkIM.Enabled = False
    chkNuclear.Enabled = False
    chkRockyMtn.Enabled = False
    chkSteelFab.Enabled = False
    chkWooster.Enabled = False
End Sub
Public Sub EnableFilters()
    chkControls.Enabled = True
    chkIM.Enabled = True
    chkNuclear.Enabled = True
    chkRockyMtn.Enabled = True
    chkSteelFab.Enabled = True
    chkWooster.Enabled = True
End Sub
Public Sub ClearFilters()
    chkControls.Value = 0
    chkIM.Value = 0
    chkNuclear.Value = 0
    chkRockyMtn.Value = 0
    chkSteelFab.Value = 0
    chkWooster.Value = 0
End Sub
Private Sub ClearAll()
    txtUsername.Text = ""
    txtFullName.Text = ""
    txtEmail.Text = ""
    chkIsAdmin.Value = False
    'chkReport.Value = False
    cmdUpdate.Enabled = False
    txtUsername.Enabled = True
    'chkDailyReport.Value = False
    cmbReportGroup.Text = 0
    'DisableFilters
End Sub
Private Sub ClearAllGroups()
    'cmbReports.ListIndex = 0
    GetReportsList
    flexGroups.Rows = 0
    flexGroups.Cols = 0
    flexGroups.Clear
End Sub
Private Sub ClearAllReport()
    txtReportName.Text = ""
    cmbRunDay.ListIndex = 0
    txtRunHour.Text = ""
    txtStartDate.Text = ""
    txtEndDate.Text = ""
    ClearFilters
    chkNewReport.Value = 0
    intSelReportID = 0
End Sub
Private Sub cmbReportGroups_Click()
    GetGroups cmbReportGroups.Text
    GetReportsList
End Sub
Private Sub cmbReports_Click()
    Dim SplitStr() As String
    Dim rs         As New ADODB.Recordset
    Dim strSQL1    As String
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM reportsgroups"
    rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
    With rs
        .AddNew
        !idGroupID = cmbReportGroups.Text
        SplitStr = Split(cmbReports.Text, "-")
        !idReportID = CInt(Trim(SplitStr(1))) 'hackity hack
        .Update
    End With
    'GetReportGroups
    ClearAllGroups
    GetGroups cmbReportGroups.Text
    MsgBox "Added to group."
End Sub
Private Sub cmdClear_Click()
    ClearAll
End Sub
Private Sub cmdCheck_Click()
    If IsInAD(LCase$(txtUsername.Text)) Then
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
    flexUserList.TextMatrix(0, 5) = "Report Group"
    'flexUserList.TextMatrix(0, 6) = "Gets Daily Report"
    flexUserList.TextMatrix(0, 6) = "Last Log-In"
    flexUserList.TextMatrix(0, 7) = "Log-Ins"
    Do Until rs.EOF
        With rs
            flexUserList.TextMatrix(.AbsolutePosition, 1) = !idFullname
            flexUserList.TextMatrix(.AbsolutePosition, 2) = !idUsers
            flexUserList.TextMatrix(.AbsolutePosition, 3) = !idEmail
            flexUserList.TextMatrix(.AbsolutePosition, 4) = CBool(!idAdmins)
            flexUserList.TextMatrix(.AbsolutePosition, 5) = !idGroupID
            'flexUserList.TextMatrix(.AbsolutePosition, 6) = CBool(!idJPTDailyReport)
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
Private Sub cmdClearReport_Click()
    ClearAllReport
End Sub
Private Sub cmdDeleteGroup_Click()
    DeleteGroup cmbReportGroups.Text
    MsgBox "Group: " & cmbReportGroups.Text & " deleted."
    GetReportGroups
    ClearAllGroups
End Sub
Private Sub cmdNewGroup_Click()
    Debug.Print NextGroup
    cmbReportGroups.AddItem NextGroup
    cmbReportGroups.ListIndex = cmbReportGroups.ListCount - 1
End Sub
Private Sub cmdRefresh_Click()
    GetUsers
End Sub

Private Sub cmdRemoveReport_Click()
DeleteReportFromGroup cmbReportGroups.Text, flexGroups.TextMatrix(flexGroups.RowSel, 2), flexGroups.TextMatrix(flexGroups.RowSel, 8)
 ClearAllGroups
    GetGroups cmbReportGroups.Text
    MsgBox "Report removed."
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo errs
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    Dim blah
    Dim SplitStr() As String
    strSQL1 = "SELECT * From users Where idUsers = '" & txtUsername.Text & "'"
    cn_global.CursorLocation = adUseClient
    If Trim$(txtUsername.Text) = "" Or Trim$(txtFullName.Text) = "" Or Trim$(txtEmail.Text) = "" Then
        blah = MsgBox("One or more fields is blank. Please fill all fields.", vbOKOnly + vbInformation, "Something's missing...")
        Exit Sub
    End If
    SplitStr = Split(txtEmail.Text, "@")
    rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
    With rs
        !idFullname = Trim$(txtFullName.Text)
        !idEmail = Trim$(SplitStr(0) & EmailDomain)
        !idAdmins = chkIsAdmin.Value
        '!idJPTReport = chkReport.Value
        '!idJPTDailyReport = chkDailyReport.Value
        '!idCompanyFilters = FilterString
        !idGroupID = cmbReportGroup.Text
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
        GetUser txtUsername.Text
    End If
End Sub
Private Sub cmdUpdateReport_Click()
    On Error GoTo errs
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    Dim blah
    Dim SplitStr() As String
    strSQL1 = "SELECT * From reports Where idReportName = '" & txtReportName.Text & "'"
    cn_global.CursorLocation = adUseClient
    rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
    With rs
        If chkNewReport.Value = 1 Then .AddNew
        !idReportName = txtReportName.Text
        !idRunDay = cmbRunDay.ListIndex + 1
        !idRunTime = txtRunHour.Text
        !idStartDate = txtStartDate.Text
        !idEndDate = txtEndDate.Text
        !idCompanyFilter = FilterString
        !idHasRun = IIf(chkHasRun.Value = 1, -1, 0)
        
        .Update
    End With
    blah = MsgBox("Report Updated!", vbOKOnly + vbInformation, "Success")
    ClearAllReport
    GetReports
    GetReportsList
    Exit Sub
errs:
    If Err.Number = -2147217864 Then
        blah = MsgBox("Nothing to update.", vbOKOnly + vbExclamation, "No Changes...")
        GetUser txtUsername.Text
    End If
End Sub
Private Sub flexReports_DblClick()
    GetReport flexReports.TextMatrix(flexReports.RowSel, 1), flexReports.TextMatrix(flexReports.RowSel, 2)
End Sub
Private Sub flexUserList_DblClick()
    GetReportGroups
    GetUser flexUserList.TextMatrix(flexUserList.RowSel, 2)
End Sub
Private Sub Form_Load()
    FindMySQLDriver
    cn_global.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    GetUsers
    GetReports
    GetReportGroups
    'DisableFilters
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
Public Sub DeleteReport(ReportID As Integer)
    Dim blah
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * From reports Where idReportID = '" & ReportID & "'"
    rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
    With rs
        .Delete
        .Update
    End With
    blah = MsgBox("ReportID: " & ReportID & " has been deleted.", vbOKOnly + vbInformation, "Success")
    ClearAllReport
    GetReports
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    cn_global.Close
    Unload Me
    End
End Sub
Private Sub mnuDelete_Click()
    Dim blah
    If txtUsername.Text <> "" Then
        blah = MsgBox("Are you sure you want to delete user: " & UCase$(txtUsername.Text) & "?", vbOKCancel + vbCritical, "Delete User")
        If blah = vbOK Then
            DeleteUser txtUsername.Text
        Else
            ClearAll
            Exit Sub
        End If
    Else
        blah = MsgBox("Please select a user first.", vbOKOnly + vbExclamation, "Nothing Selected")
    End If
End Sub
Private Sub mnuDeleteReport_Click()
    Dim blah
    If intSelReportID > 0 Then
        blah = MsgBox("Are you sure you want to delete report: " & txtReportName.Text & "?", vbOKCancel + vbCritical, "Delete Report")
        If blah = vbOK Then
            DeleteReport intSelReportID
        Else
            ClearAllReport
            Exit Sub
        End If
    Else
        blah = MsgBox("Please select a report first.", vbOKOnly + vbExclamation, "Nothing Selected")
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
Private Sub SSTab1_Click(PreviousTab As Integer)
    Select Case SSTab1.Tab
        Case 0
            '            mnuDelete.Visible = True
            '            mnuNewUser.Visible = True
            '            mnuDeleteReport.Visible = False
            '            mnuDeleteGroup.Visible = False
        Case 1
            '            mnuDelete.Visible = False
            '            mnuNewUser.Visible = False
            '            mnuDeleteReport.Visible = True
            '            mnuDeleteGroup.Visible = False
            GetReports
        Case 2
            '            mnuDelete.Visible = False
            '            mnuNewUser.Visible = False
            '            mnuDeleteReport.Visible = False
            '            mnuDeleteGroup.Visible = True
    End Select
End Sub
Public Sub GetReports()
    SetupRunDayCombo
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM reports ORDER BY idReportID"
    Set rs = cn_global.Execute(strSQL1)
    flexReports.Clear
    flexReports.Rows = 2
    flexReports.FixedCols = 1
    flexReports.FixedRows = 1
    flexReports.Rows = rs.RecordCount + 1
    flexReports.Cols = 9
    flexReports.TextMatrix(0, 1) = "Report Name"
    flexReports.TextMatrix(0, 2) = "ReportID"
    flexReports.TextMatrix(0, 3) = "Run Day"
    flexReports.TextMatrix(0, 4) = "Run Hour"
    flexReports.TextMatrix(0, 5) = "Start Date"
    flexReports.TextMatrix(0, 6) = "End Date"
    flexReports.TextMatrix(0, 7) = "Company Filter"
    flexReports.TextMatrix(0, 8) = "Has Run"
    Do Until rs.EOF
        With rs
            flexReports.TextMatrix(.AbsolutePosition, 1) = !idReportName
            flexReports.TextMatrix(.AbsolutePosition, 2) = !idReportID
            flexReports.TextMatrix(.AbsolutePosition, 3) = !idRunDay
            flexReports.TextMatrix(.AbsolutePosition, 4) = !idRunTime
            flexReports.TextMatrix(.AbsolutePosition, 5) = !idStartDate
            flexReports.TextMatrix(.AbsolutePosition, 6) = !idEndDate
            flexReports.TextMatrix(.AbsolutePosition, 7) = !idCompanyFilter
            flexReports.TextMatrix(.AbsolutePosition, 8) = CBool(!idHasRun)
            .MoveNext
        End With
    Loop
    SizeTheSheet flexReports
End Sub
Public Sub GetReportsList()
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM reports ORDER BY idReportID"
    Set rs = cn_global.Execute(strSQL1)
    cmbReports.Clear
    cmbReports.AddItem ""
    Do Until rs.EOF
        With rs
            cmbReports.AddItem !idReportName & " - " & !idReportID
            .MoveNext
        End With
    Loop
End Sub
Public Sub SetupRunDayCombo()
    cmbRunDay.Clear
    cmbRunDay.AddItem "1-Sunday" ', 1
    cmbRunDay.AddItem "2-Monday" ', 2
    cmbRunDay.AddItem "3-Tuesday" ', 3
    cmbRunDay.AddItem "4-Wednesday" ', 4
    cmbRunDay.AddItem "5-Thursday" ', 5
    cmbRunDay.AddItem "6-Friday" ', 6
    cmbRunDay.AddItem "7-Saturday" ', 7
    cmbRunDay.AddItem "8-Every Weekday" ', 8
End Sub
