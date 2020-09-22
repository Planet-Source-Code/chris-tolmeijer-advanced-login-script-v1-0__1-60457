VERSION 5.00
Begin VB.Form lookup 
   Caption         =   "Start"
   ClientHeight    =   705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2205
   LinkTopic       =   "Form1"
   ScaleHeight     =   705
   ScaleWidth      =   2205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click to start"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      DataField       =   "Veld1"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Text            =   "empty"
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Mijn eerste progammatjes\Advanced login\user.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Settings"
      Top             =   2160
      Width           =   2340
   End
   Begin VB.Label lble 
      Caption         =   "©Copyright Chris Tolmeijer"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "lookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()


End Sub

Private Sub Command1_Click()
If Text1.Text = "" Then

lookup.Visible = False
Firsttime.Visible = True

Else

lookup.Visible = False
login.Visible = True

End If
End Sub
