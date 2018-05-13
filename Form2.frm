VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000004&
   Caption         =   "Manage"
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6990
   LinkTopic       =   "Form2"
   ScaleHeight     =   5520
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Home"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   840
      TabIndex        =   3
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   2
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000E&
      Caption         =   "Add Entity"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   1
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "Entities:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i As Integer
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
Unload Me
Form1.Show
End Sub

Private Sub Command3_Click()
Unload Me
Form1.Show
End Sub

Private Sub Command2_Click()
Form3.Show
End Sub


Private Sub Command5_Click()
prod = List1.List(List1.ListIndex)
Unload Me
Form4.Show
End Sub

Private Sub Form_Load()
Command5.Enabled = False
db.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\invent_db.mdb;Persist Security Info=False")
sql = "Select * from table_invent;"
Set rs = db.Execute(sql)
i = 0
If (rs.BOF = False) Then
    rs.MoveFirst
    Do Until (rs.EOF = True)
    List1.AddItem rs.Fields(1)
    i = i + 1
    rs.MoveNext
    Loop
Else
    MsgBox "No more records !!"
End If
Set rs = Nothing
db.Close
End Sub

Private Sub List1_Click()
Command5.Enabled = True
End Sub
