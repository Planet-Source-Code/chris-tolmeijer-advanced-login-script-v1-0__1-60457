VERSION 5.00
Begin VB.Form Firsttime 
   Caption         =   "Your First Time"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   3810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "I accept all terms and agreements"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2040
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      DataField       =   "Veld2"
      DataSource      =   "Data1"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      DataField       =   "Veld1"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   3495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Mijn eerste progammatjes\Advanced login\user.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Settings"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label4 
      Caption         =   "It is better to close the program after you have registered"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Confirm password:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label info 
      Caption         =   "This is your first time you use this program you must set an username and password first!"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Firsttime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
    MsgBox "You must fill in an username!"

Else

If Text2.Text = "" Then
    MsgBox "You must fill in a password!"

Else

If Text3.Text = Text2.Text Then


    Data1.Recordset.Update
    Firsttime.Visible = False
    login.Visible = True

Else
    MsgBox "Youre passwords are not the same!"
    
End If
End If
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Command3_Click()
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Data1.Recordset.AddNew


End Sub
