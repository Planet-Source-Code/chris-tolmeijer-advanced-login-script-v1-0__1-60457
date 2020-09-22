VERSION 5.00
Begin VB.Form changename2 
   Caption         =   "Change Name"
   ClientHeight    =   990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2340
   LinkTopic       =   "Form1"
   ScaleHeight     =   990
   ScaleWidth      =   2340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Mijn eerste progammatjes\Advanced login\user.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Settings"
      Top             =   2760
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      DataField       =   "Veld1"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Now change your username: "
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "changename2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
    MsgBox "You must fill in an username!"

Else
Data1.Recordset.Edit
    changename2.Visible = False
    Menu.Visible = True
    MsgBox "Your username has been changed!"
    
End If

End Sub
