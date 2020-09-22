VERSION 5.00
Begin VB.Form delete 
   Caption         =   "WARNING!!!"
   ClientHeight    =   1155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   1155
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      DataField       =   "Veld2"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      DataField       =   "Veld1"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Mijn eerste progammatjes\Advanced login\user.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Settings"
      Top             =   3840
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "This will delete your account for ever! There is no undo feature!!!! Are you shure you want to delete your account?"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "delete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Data1.Recordset.delete

delete.Visible = False
lookup.Visible = True

MsgBox "Your acocunt is deleted..."


End Sub

Private Sub Command2_Click()
delete.Visible = False
Menu.Visible = True

End Sub

