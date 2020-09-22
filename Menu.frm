VERSION 5.00
Begin VB.Form Menu 
   Caption         =   "Main Menu"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2295
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   2295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Logout"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete Account"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change Username"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Password"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Mijn eerste progammatjes\Advanced login\user.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Settings"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      DataField       =   "Veld1"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome,"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Menu.Visible = False
changepass.Visible = True

End Sub

Private Sub Command2_Click()
changename.Visible = True
Menu.Visible = False

End Sub

Private Sub Command3_Click()
Menu.Visible = False
delete.Visible = True

End Sub

Private Sub Command4_Click()
Menu.Visible = False
lookup.Visible = True

End Sub
