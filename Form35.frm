VERSION 5.00
Begin VB.Form Form35 
   Caption         =   "Teacher's Information"
   ClientHeight    =   8055
   ClientLeft      =   1560
   ClientTop       =   495
   ClientWidth     =   9885
   LinkTopic       =   "Form35"
   Picture         =   "Form35.frx":0000
   ScaleHeight     =   8055
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   2280
      TabIndex        =   33
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   8520
      TabIndex        =   31
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   2280
      TabIndex        =   30
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   2280
      TabIndex        =   29
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   7200
      TabIndex        =   28
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8520
      TabIndex        =   26
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   375
      Left            =   8520
      TabIndex        =   25
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   7200
      TabIndex        =   24
      Top             =   7200
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   7200
      TabIndex        =   23
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   7200
      TabIndex        =   22
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   7200
      TabIndex        =   21
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   7200
      TabIndex        =   20
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   7200
      TabIndex        =   19
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2280
      TabIndex        =   18
      Top             =   6600
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2280
      TabIndex        =   17
      Top             =   6000
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Top             =   5400
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2280
      TabIndex        =   15
      Top             =   4200
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2280
      TabIndex        =   14
      Text            =   "Select ID"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "Contact Number"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   32
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label Label15 
      Caption         =   "Gender"
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
      Left            =   5640
      TabIndex        =   27
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "Religion"
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
      Left            =   5640
      TabIndex        =   13
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "Teacher's Name"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label12 
      Caption         =   "Teaching Subject"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   11
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Job Period"
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
      Left            =   5640
      TabIndex        =   10
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Joining Date"
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
      Left            =   5640
      TabIndex        =   9
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Post"
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
      Left            =   5640
      TabIndex        =   8
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Merital Status"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Age Today"
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
      Left            =   480
      TabIndex        =   6
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Date of Birth"
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
      Left            =   480
      TabIndex        =   5
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Address"
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
      Left            =   480
      TabIndex        =   4
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Mother's Name"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Father's Name"
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
      Left            =   480
      TabIndex        =   2
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Teacher's ID"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Teacher's Information"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "Form35"
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
.Open "select * from Teacher Where Teacher.TeacherID='" & id & "'", cnn, adOpenKeyset, adLockOptimistic
End With
With rs2
Text12.Text = .Fields(1)
Text1.Text = .Fields(2)
Text13.Text = .Fields(3)
Text2.Text = .Fields(4)
Text3.Text = .Fields(5)
Text11.Text = .Fields(11)
Text5.Text = .Fields(6)
Text6.Text = .Fields(7)
Text7.Text = .Fields(8)
Text8.Text = .Fields(9)
Text10.Text = .Fields(10)
Text14.Text = .Fields(12)

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

monthtj = Text8.Text
datejob = CDate(monthtj)
yearjob = Year(datejob)
year2 = yeartd - yearjob
monthjob = Month(datejob)
monthtd = Month(Date)
month2 = monthtd - monthjob
If month2 < 0 Then
monthtd = monthtd + 12
month2 = monthtd - monthjob
year2 = year2 - 1
End If
Text9.Text = CStr(year2) + "years" + CStr(month2) + "months"
End With
cnn2.Close
rs2.Close
End Sub
Private Sub Command1_Click()
Unload Me
Form2.Show

End Sub
Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form35.Hide
Form1.Show
End Sub

Private Sub Form_Load()
With cnn
.Provider = "Microsoft.Jet.OLEDB.4.0"
.ConnectionString = "C:\Project\TITASH GAS HIGH SCHOOL.mdb"
.Open
End With
With rs
.Open "Select Teacher.TeacherID from Teacher", cnn, adOpenKeyset, adLockOptimistic
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
Text12.Visible = True
Text1.Visible = True
Text13.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Text14.Visible = True
Text11.Visible = True
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text8.Visible = True
Text9.Visible = True
Text10.Visible = True
End If
End Sub

