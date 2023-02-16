VERSION 5.00
Begin VB.Form Form16 
   Caption         =   "Student Information (Nine Human)"
   ClientHeight    =   7770
   ClientLeft      =   1020
   ClientTop       =   495
   ClientWidth     =   11085
   LinkTopic       =   "Form16"
   Picture         =   "Form16.frx":0000
   ScaleHeight     =   7770
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   9840
      TabIndex        =   27
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   13
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   12
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3720
      TabIndex        =   11
      Text            =   "Select ID"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      Height          =   405
      Left            =   7800
      TabIndex        =   10
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   7800
      TabIndex        =   7
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   7800
      TabIndex        =   6
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Height          =   405
      Left            =   7800
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Height          =   405
      Left            =   3720
      TabIndex        =   4
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   7800
      TabIndex        =   0
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Student Information(Nine Human)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   26
      Top             =   720
      Width           =   6135
   End
   Begin VB.Label Label2 
      Caption         =   " Student's ID"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   25
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label12 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   24
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label11 
      Caption         =   "Mother's Name"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   23
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "Father's Name"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   22
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Age Today"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   21
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Birth_Date"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   20
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Roll Number"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   19
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   18
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Gurdian's Name"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   17
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Student's Name"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Contact Number"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   15
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label Label13 
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   14
      Top             =   2520
      Width           =   1935
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim cnn2 As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Sub Fill_TextBox()
With cnn2

.Provider = "Microsoft.Jet.OLEDB.4.0"
.ConnectionString = "C:\Project\TITASH GAS HIGH SCHOOL.mdb"
.Open
End With
Dim id, monthtx, yeartx, monthtj As String
id = Combo1.Text
Dim month1, month2, monthtd, monthbd, monthjob As Integer
Dim datebd, datejob As Date
With rs2
.Open "select * from NineHuman Where NineHuman.StudentID='" & id & "'", cnn, adOpenKeyset, adLockOptimistic
End With
With rs2
Text1.Text = .Fields(1)
Text2.Text = .Fields(2)
Text3.Text = .Fields(7)
Text5.Text = .Fields(3)
Text6.Text = .Fields(4)
Text7.Text = .Fields(5)
Text8.Text = .Fields(6)
Text9.Text = .Fields(8)
Text10.Text = .Fields(9)
Text11.Text = .Fields(10)
monthtx = Text3.Text
datebd = CDate(monthtx)

yeartd = Year(Date)
yearbd = Year(datebd)
year1 = yeartd - yearbd

monthbd = Month(datebd)
monthtd = Month(Date)
month1 = monthtd - monthbd
If month1 < 0 Then
monthtd = monthtd + 12
month1 = monthtd - monthb
year1 = year1 - 1
End If
Text4.Text = CStr(year1) + " years " + CStr(month1) + "months"


End With
cnn2.Close
rs2.Close
End Sub
Private Sub Command1_Click()
Unload Me
Form3.Show

End Sub
Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form16.Hide
Form1.Show
End Sub

Private Sub Form_Load()
With cnn
.Provider = "Microsoft.Jet.OLEDB.4.0"
.ConnectionString = "C:\Project\TITASH GAS HIGH SCHOOL.mdb"
.Open
End With
With rs
.Open "Select NineHuman.StudentID from NineHuman", cnn, adOpenKeyset, adLockOptimistic
End With
With rs
Dim str As String
Dim record, i, field As Integer
record = rs.RecordCount
For i = 0 To record - 1

str = .Fields(0)
Combo1.AddItem str
rs.MoveNext
Next
End With
End Sub
Private Sub Combo1_Click()
Select Case Combo1.ListIndex
Case 0
Call Fill_TextBox
Case 1
Call Fill_TextBox
Case 2
Call Fill_TextBox
Case 3
Call Fill_TextBox
Case 4
Call Fill_TextBox
End Select
End Sub
Private Sub Command8_Click()
If Text19.Text = Text18.Text Then

Label21.Visible = False
Text18.Visible = False
Command8.Visible = False
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text8.Visible = True
Text9.Visible = True
Text10.Visible = True
Text11.Visible = True
End If
End Sub


