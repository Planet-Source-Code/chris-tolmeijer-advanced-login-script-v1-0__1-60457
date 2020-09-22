VERSION 5.00
Begin VB.Form changepass 
   Caption         =   "Change Password"
   ClientHeight    =   990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   ScaleHeight     =   990
   ScaleWidth      =   2535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
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
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Settings"
      Top             =   2880
      Width           =   1140
   End
   Begin VB.Label Label2 
      Caption         =   "Fill in your old password!"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      DataField       =   "Veld2"
      DataSource      =   "Data1"
      Height          =   975
      Left            =   600
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "changepass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
    MsgBox "You must fill in your old password!"
    
Else

If Text1.Text = Label1.Caption Then
    changepass.Visible = False
    Changepass2.Visible = True
Else
    MsgBox "The password is incorrect"
    
End If
End If
End Sub

Private Sub Command2_Click()
changepass.Visible = False
Menu.Visible = True

End Sub
