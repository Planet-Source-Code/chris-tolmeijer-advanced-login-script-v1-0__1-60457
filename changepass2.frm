VERSION 5.00
Begin VB.Form changepass2 
   Caption         =   "New Pass"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2280
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   2280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      DataField       =   "Veld2"
      DataSource      =   "Data1"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Mijn eerste progammatjes\Advanced login\user.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Settings"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Confirm New Password"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "New Password"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Changepass2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
    MsgBox "You must fill in a correct password"
Else
If Text1.Text = Text2.Text Then
Data1.Recordset.Edit
    Changepass2.Visible = False
    Menu.Visible = True
    MsgBox "The password has been changed!"
Else
    MsgBox "The 2 Passwords are not the same!"

End If
End If
End Sub

Private Sub Command2_Click()
Menu.Visible = True
Changepass2.Visible = False

End Sub

