VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7470
   LinkTopic       =   "Form4"
   ScaleHeight     =   3495
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Home"
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Back"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete items"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   10
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add items"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   8
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Enter number of items here:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label Label8 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label7 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Quantity:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Price:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Product Name:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Product ID:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strDate As Date
Dim mul As Single



Private Sub Command1_Click()
If Not IsNumeric(Text1.Text) Then
    MsgBox "Wrong input", vbCritical, "error"
    Exit Sub
End If

db.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\invent_db.mdb;Persist Security Info=False")

sql = "update table_invent set Quantity = Quantity + " & Text1.Text & " where Product like '" & prod & "';"
Set rs = db.Execute(sql)
strDate = Format(Now, "dd/mm/yyyy")
mul = (Val(Label7.Caption) * Val(Text1.Text))
sql = "INSERT INTO Trans (Quantity, ProdID, Product, Trans_date, Amount, Status) VALUES ('" & Val(Text1.Text) & "', " & Label5.Caption & ", '" & Label6.Caption & "', '" & strDate & "', " & mul & ", 'IN');"
Set rs = db.Execute(sql)

sql = "Select Quantity from table_invent where Product like '" & prod & "';"
Set rs = db.Execute(sql)
Label8.Caption = rs.Fields(0)

MsgBox "Update successful", vbInformation, "success"
Text1.Text = ""

db.Close
End Sub

Private Sub Command2_Click()

If Not IsNumeric(Text1.Text) Then
    MsgBox "Wrong input", vbCritical, "error"
    Exit Sub
End If

db.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\invent_db.mdb;Persist Security Info=False")

sql = "Select Quantity from table_invent where Product like '" & prod & "';"
Set rs = db.Execute(sql)

If rs.Fields(0) - Val(Text1.Text) < 0 Then
    MsgBox "No more items", vbCritical, "error"
    Exit Sub
End If

sql = "update table_invent set Quantity = Quantity - " & Val(Text1.Text) & " where Product like '" & prod & "';"
Set rs = db.Execute(sql)

strDate = Format(Now, "dd/mm/yyyy")
mul = (Val(Label7.Caption) * Val(Text1.Text))
sql = "INSERT INTO Trans (Quantity, ProdID, Product, Trans_date, Amount, Status) VALUES ('" & Val(Text1.Text) & "', " & Label5.Caption & ", '" & Label6.Caption & "', '" & strDate & "', " & mul & ", 'OUT');"
Set rs = db.Execute(sql)

sql = "Select Quantity from table_invent where Product like '" & prod & "';"
Set rs = db.Execute(sql)
Label8.Caption = rs.Fields(0)

MsgBox "Update successful", vbInformation, "success"
Text1.Text = ""

db.Close
End Sub

Private Sub Command3_Click()
Unload Me
Form2.Show

End Sub

Private Sub Command4_Click()
Unload Me
Form1.Show

End Sub

Private Sub Form_Load()
db.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\invent_db.mdb;Persist Security Info=False")

sql = "Select * from table_invent where Product like '" & prod & "';"
Set rs = db.Execute(sql)
Label5.Caption = rs.Fields(0)
Label6.Caption = rs.Fields(1)
Label7.Caption = rs.Fields(3)
Label8.Caption = rs.Fields(2)

db.Close

End Sub

