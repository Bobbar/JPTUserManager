VERSION 5.00
Begin VB.Form frmAddNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New User"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4530
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddNew.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   360
      Left            =   2580
      TabIndex        =   10
      Top             =   2520
      Width           =   990
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Enabled         =   0   'False
      Height          =   360
      Left            =   720
      TabIndex        =   9
      Top             =   2520
      Width           =   1290
   End
   Begin VB.CheckBox chkReport 
      Caption         =   "Gets Report"
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Top             =   2040
      Width           =   1395
   End
   Begin VB.CheckBox chkIsAdmin 
      Caption         =   "Is Admin"
      Height          =   315
      Left            =   780
      TabIndex        =   3
      Top             =   2040
      Width           =   1035
   End
   Begin VB.TextBox txtEmail 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   1620
      Width           =   3915
   End
   Begin VB.TextBox txtUsername 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   420
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   645
      Left            =   2820
      TabIndex        =   5
      Top             =   480
      Width           =   1290
   End
   Begin VB.TextBox txtFullName 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1020
      Width           =   2355
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email (@worthingtonindustries.com):"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1380
      Width           =   2640
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name:  (Last, First)"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   780
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   180
      Width           =   780
   End
End
Attribute VB_Name = "frmAddNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
    AddToUserList txtUsername.Text, txtFullName.Text, txtEmail.Text, chkIsAdmin.Value, chkReport.Value
    ClearAll
    Form1.GetUsers
End Sub

Private Sub cmdClear_Click()
    ClearAll
End Sub

Private Sub cmdSearch_Click()
    If txtFullName.Text <> "" Then
        If FindFullName(txtFullName.Text).Username <> "" Then
            txtUsername.Text = FindFullName(txtFullName.Text).Username
            txtFullName.Text = FindFullName(txtFullName.Text).Fullname
            txtEmail.Text = FindFullName(txtFullName.Text).Email
            cmdAdd.Enabled = True
            txtUsername.Enabled = False
            txtFullName.Enabled = False
            txtEmail.Enabled = False
        End If
    End If
    If txtUsername.Text <> "" Then
        If FindUserNamePartial(txtUsername.Text).Username <> "" Then
            txtUsername.Text = FindUserNamePartial(txtUsername.Text).Username
            txtFullName.Text = FindUserNamePartial(txtUsername.Text).Fullname
            txtEmail.Text = FindUserNamePartial(txtUsername.Text).Email
            cmdAdd.Enabled = True
            txtUsername.Enabled = False
            txtFullName.Enabled = False
            txtEmail.Enabled = False
        End If
    End If
    If txtEmail.Text <> "" Then
        If FindEmail(txtEmail.Text).Username <> "" Then
            txtUsername.Text = FindEmail(txtEmail.Text).Username
            txtFullName.Text = FindEmail(txtEmail.Text).Fullname
            txtEmail.Text = FindEmail(txtEmail.Text).Email
            cmdAdd.Enabled = True
            txtUsername.Enabled = False
            txtFullName.Enabled = False
            txtEmail.Enabled = False
        End If
    End If
    If IsInDB(txtUsername.Text) Then
        blah = MsgBox("Already in userlist!", vbCritical + vbOKOnly, "Error")
        ClearAll
    End If
End Sub
Private Sub ClearAll()
    txtUsername.Text = ""
    txtFullName.Text = ""
    txtEmail.Text = ""
    chkIsAdmin.Value = 0
    chkReport.Value = 0
    txtUsername.Enabled = True
    txtFullName.Enabled = True
    txtEmail.Enabled = True
    cmdAdd.Enabled = False
End Sub
