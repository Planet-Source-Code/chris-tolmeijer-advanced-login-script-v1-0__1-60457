VERSION 5.00
Begin VB.Form login 
   Caption         =   "Login"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   2535
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Mijn eerste progammatjes\Advanced login\user.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Settings"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "password"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "username"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label2 
      DataField       =   "Veld2"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label1 
      DataField       =   "Veld1"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
    MsgBox "You must fill in an username"

Else
If Text2.Text = "" Then
    MsgBox "You must fill in a password"

Else
If Text1.Text <> Label1.Caption Then
    MsgBox "the username is Incorrect!"
Else
If Text2.Text <> Label2.Caption Then
    MsgBox "the password is incorrect"
    
Else
If Text1.Text = Label1.Caption Then
If Text2.Text = Label2.Caption Then
    login.Visible = False
    Menu.Visible = True
    
End If
End If
End If
End If
End If
End If

End Sub
