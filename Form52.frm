VERSION 5.00
Begin VB.Form Form52 
   Caption         =   "Result's of Pree-Test(Science)"
   ClientHeight    =   8265
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12045
   LinkTopic       =   "Form52"
   ScaleHeight     =   8265
   ScaleWidth      =   12045
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text34 
      Height          =   375
      Left            =   10800
      TabIndex        =   43
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox Text33 
      Height          =   375
      Left            =   8400
      TabIndex        =   42
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox Text32 
      Height          =   375
      Left            =   8400
      TabIndex        =   41
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox Text31 
      Height          =   375
      Left            =   8400
      TabIndex        =   40
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox Text30 
      Height          =   375
      Left            =   8400
      TabIndex        =   39
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text29 
      Height          =   375
      Left            =   5880
      TabIndex        =   38
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox Text28 
      Height          =   375
      Left            =   5880
      TabIndex        =   37
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox Text27 
      Height          =   375
      Left            =   5880
      TabIndex        =   36
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox Text26 
      Height          =   375
      Left            =   5880
      TabIndex        =   35
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text25 
      Height          =   375
      Left            =   3360
      TabIndex        =   34
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox Text24 
      Height          =   375
      Left            =   3360
      TabIndex        =   33
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox Text23 
      Height          =   375
      Left            =   3360
      TabIndex        =   32
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox Text22 
      Height          =   375
      Left            =   3360
      TabIndex        =   31
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text21 
      Height          =   375
      Left            =   5160
      TabIndex        =   30
      Top             =   7440
      Width           =   735
   End
   Begin VB.TextBox Text20 
      Height          =   375
      Left            =   2280
      TabIndex        =   29
      Top             =   7440
      Width           =   2655
   End
   Begin VB.TextBox Text19 
      Height          =   375
      Left            =   360
      TabIndex        =   28
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox Text18 
      Height          =   375
      Left            =   360
      TabIndex        =   27
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   26
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   360
      TabIndex        =   25
      Top             =   7440
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   6720
      TabIndex        =   24
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   6720
      TabIndex        =   23
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   6720
      TabIndex        =   22
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   4200
      TabIndex        =   21
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   4200
      TabIndex        =   20
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox Text111 
      Height          =   375
      Left            =   9240
      TabIndex        =   19
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   6720
      TabIndex        =   18
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   4200
      TabIndex        =   17
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2040
      TabIndex        =   14
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   2640
      Width           =   1575
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
      Left            =   360
      TabIndex        =   9
      Text            =   "Select ID"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   9240
      TabIndex        =   8
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text35 
      Height          =   375
      Left            =   10800
      TabIndex        =   7
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Total"
      Height          =   375
      Left            =   10560
      TabIndex        =   4
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Fomat to GPA"
      Height          =   375
      Left            =   10560
      TabIndex        =   3
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   10560
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label28 
      Caption         =   "GPA"
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
      Left            =   10800
      TabIndex        =   70
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label27 
      Caption         =   "GPA"
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
      Left            =   8400
      TabIndex        =   69
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label26 
      Caption         =   "GPA"
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
      Left            =   5880
      TabIndex        =   68
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label25 
      Caption         =   "GPA"
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
      Left            =   3360
      TabIndex        =   67
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label24 
      Caption         =   "Total GPA"
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
      Left            =   5160
      TabIndex        =   66
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label23 
      Caption         =   "Total number pass/Fail"
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
      Left            =   2280
      TabIndex        =   65
      Top             =   6960
      Width           =   2655
   End
   Begin VB.Label Label22 
      Caption         =   "Optional"
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
      Left            =   360
      TabIndex        =   64
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label21 
      Caption         =   "Compulsory"
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
      Left            =   360
      TabIndex        =   63
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label20 
      Caption         =   "Total number"
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
      Left            =   360
      TabIndex        =   62
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label Label19 
      Caption         =   "Agriculture"
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
      Left            =   6720
      TabIndex        =   61
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Bengali1st"
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
      Left            =   2040
      TabIndex        =   60
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Bengali2nd"
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
      Left            =   2040
      TabIndex        =   59
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "Computer"
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
      Left            =   6720
      TabIndex        =   58
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label15 
      Caption         =   "Higher Math"
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
      Left            =   9240
      TabIndex        =   57
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "Biology"
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
      Left            =   6720
      TabIndex        =   56
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label13 
      Caption         =   "Chamistry"
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
      Left            =   4200
      TabIndex        =   55
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "Physics"
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
      Left            =   4200
      TabIndex        =   54
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Socoal Science"
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
      Left            =   4200
      TabIndex        =   53
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Genaral Math"
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
      Left            =   6720
      TabIndex        =   52
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Religion"
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
      Left            =   4200
      TabIndex        =   51
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "English2nd"
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
      Left            =   2040
      TabIndex        =   50
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "English1st"
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
      Left            =   2040
      TabIndex        =   49
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   48
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label17 
      Caption         =   "Roll Number"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   47
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Student's Name"
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
      Left            =   360
      TabIndex        =   46
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Student'ID"
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
      Left            =   360
      TabIndex        =   45
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label29 
      Caption         =   "Home Economics"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   44
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Result's of Pree-Test(Science)"
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
      Left            =   3600
      TabIndex        =   6
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label2 
      Caption         =   "Please Select Student ID to see Result"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1080
      Width           =   4095
   End
End
Attribute VB_Name = "Form52"
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
Dim id As Integer
id = Combo1.Text
With rs2
.Open "select * from ResultofPreeTestScience where ResultofPreeTestScience.StudentID='" & id & "'", cnn, adOpenKeyset, adLockOptimistic
End With
With rs2
Text1.Text = .Fields(4)
Text2.Text = .Fields(5)
Text3.Text = .Fields(6)
Text4.Text = .Fields(7)
Text5.Text = .Fields(8)
Text6.Text = .Fields(9)
Text7.Text = .Fields(10)
Text8.Text = .Fields(11)
Text9.Text = .Fields(12)
Text10.Text = .Fields(13)
Text11.Text = .Fields(14)
Text12.Text = .Fields(15)
Text111.Text = .Fields(16)
Text13.Text = .Fields(1)
Text14.Text = .Fields(3)
Text15.Text = .Fields(2)
Text18.Text = .Fields(18)
Text19.Text = .Fields(19)
Text16.Text = .Fields(17)
End With

cnn2.Close
rs2.Close
End Sub

Private Sub Command1_Click()
Form46.Hide
Form36.Show
End Sub
Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form46.Hide
Form1.Show
End Sub

Private Sub Command5_Click()
 Text17.Text = (Text1.Text - -Text2.Text - -Text3.Text - -Text4.Text - -Text5.Text - -Text6.Text - -Text7.Text - -Text8.Text - -Text9.Text - -Text10.Text - -Text11.Text - -Text12.Text - -Text111.Text - -Text16.Text)
Text21.Text = (Text22.Text - -Text23.Text - -Text24.Text - -Text25.Text - -Text26.Text - -Text27.Text - -Text28.Text - -Text29.Text - -Text30.Text - -Text31.Text - -Text32.Text - -Text33.Text - -Text34.Text - -Text35.Text) / 11


End Sub

Private Sub Form_Load()
With cnn
.Provider = "Microsoft.Jet.OLEDB.4.0"
.ConnectionString = "C:\Project\TITASH GAS HIGH SCHOOL.mdb"
.Open
End With
With rs
.Open "Select ResultofPreeTestScience.StudentID from ResultofPreeTestScience where ResultofPreeTestScience.Class = 'PreeTest(Science)'", cnn, adOpenKeyset, adLockOptimistic
End With
With rs
Dim str As String
Dim record, i, field As Integer
record = rs.RecordCount
For i = 0 To record - 1
field = .Fields(0)
str = CStr(field)
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
Case 5
Call Fill_TextBox
Case 6
Call Fill_TextBox
Case 7
Call Fill_TextBox
Case 8
Call Fill_TextBox
Case 9
Call Fill_TextBox
End Select
End Sub

Private Sub Command4_Click()
Command5.Visible = True
If (Text1.Text) >= 80 Then
Text22.Text = 5
End If
If (Text1.Text) <= 79 Then
Text22.Text = 4
End If

If (Text1.Text) <= 69 Then
Text22.Text = 3.5
End If

If (Text1.Text) <= 59 Then
Text22.Text = 3
End If

If (Text1.Text) <= 49 Then
Text22.Text = 2
End If

If (Text1.Text) <= 39 Then
Text22.Text = 1
End If

If (Text1.Text) <= 32 Then
Text22.Text = 0
End If
If (Text2.Text) >= 80 Then
Text23.Text = 5
End If
If (Text2.Text) <= 79 Then
Text23.Text = 4
End If

If (Text2.Text) <= 69 Then
Text23.Text = 3.5
End If

If (Text2.Text) <= 59 Then
Text23.Text = 3
End If

If (Text2.Text) <= 49 Then
Text23.Text = 2
End If

If (Text2.Text) <= 39 Then
Text23.Text = 1
End If

If (Text2.Text) <= 32 Then
Text23.Text = 0
End If
If (Text3.Text) >= 80 Then
Text24.Text = 5
End If
If (Text3.Text) <= 79 Then
Text24.Text = 4
End If

If (Text3.Text) <= 69 Then
Text24.Text = 3.5
End If

If (Text3.Text) <= 59 Then
Text24.Text = 3
End If

If (Text3.Text) <= 49 Then
Text24.Text = 2
End If

If (Text3.Text) <= 39 Then
Text24.Text = 1
End If

If (Text3.Text) <= 32 Then
Text24.Text = 0
End If
If (Text4.Text) >= 80 Then
Text25.Text = 5
End If
If (Text4.Text) <= 79 Then
Text25.Text = 4
End If

If (Text4.Text) <= 69 Then
Text25.Text = 3.5
End If

If (Text4.Text) <= 59 Then
Text25.Text = 3
End If

If (Text4.Text) <= 49 Then
Text25.Text = 2
End If

If (Text4.Text) <= 39 Then
Text25.Text = 1
End If

If (Text4.Text) <= 32 Then
Text25.Text = 0
End If

If (Text5.Text) >= 80 Then
Text26.Text = 5
End If
If (Text5.Text) <= 79 Then
Text26.Text = 4
End If

If (Text5.Text) <= 69 Then
Text26.Text = 3.5
End If

If (Text5.Text) <= 59 Then
Text26.Text = 3
End If

If (Text5.Text) <= 49 Then
Text26.Text = 2
End If

If (Text5.Text) <= 39 Then
Text26.Text = 1
End If

If (Text5.Text) <= 32 Then
Text26.Text = 0
End If

If (Text6.Text) >= 80 Then
Text27.Text = 5
End If
If (Text6.Text) <= 79 Then
Text27.Text = 4
End If

If (Text6.Text) <= 69 Then
Text27.Text = 3.5
End If

If (Text6.Text) <= 59 Then
Text27.Text = 3
End If

If (Text6.Text) <= 49 Then
Text27.Text = 2
End If

If (Text6.Text) <= 39 Then
Text27.Text = 1
End If

If (Text6.Text) <= 32 Then
Text27.Text = 0
End If

If (Text7.Text) >= 80 Then
Text28.Text = 5
End If
If (Text7.Text) <= 79 Then
Text28.Text = 4
End If

If (Text7.Text) <= 69 Then
Text28.Text = 3.5
End If

If (Text7.Text) <= 59 Then
Text28.Text = 3
End If

If (Text7.Text) <= 49 Then
Text28.Text = 2
End If

If (Text7.Text) <= 39 Then
Text28.Text = 1
End If

If (Text7.Text) <= 32 Then
Text28.Text = 0
End If
If (Text8.Text) >= 80 Then
Text29.Text = 5
End If
If (Text8.Text) <= 79 Then
Text29.Text = 4
End If

If (Text8.Text) <= 69 Then
Text29.Text = 3.5
End If

If (Text8.Text) <= 59 Then
Text29.Text = 3
End If

If (Text8.Text) <= 49 Then
Text29.Text = 2
End If

If (Text8.Text) <= 39 Then
Text29.Text = 1
End If

If (Text8.Text) <= 32 Then
Text29.Text = 0
End If
If (Text9.Text) >= 80 Then
Text30.Text = 5
End If
If (Text9.Text) <= 79 Then
Text30.Text = 4
End If

If (Text9.Text) <= 69 Then
Text30.Text = 3.5
End If

If (Text9.Text) <= 59 Then
Text30.Text = 3
End If

If (Text9.Text) <= 49 Then
Text30.Text = 2
End If

If (Text9.Text) <= 39 Then
Text30.Text = 1
End If

If (Text9.Text) <= 32 Then
Text30.Text = 0
End If
If (Text10.Text) >= 80 Then
Text31.Text = 5
End If
If (Text10.Text) <= 79 Then
Text31.Text = 4
End If

If (Text10.Text) <= 69 Then
Text31.Text = 3.5
End If

If (Text10.Text) <= 59 Then
Text31.Text = 3
End If

If (Text10.Text) <= 49 Then
Text31.Text = 2
End If

If (Text10.Text) <= 39 Then
Text31.Text = 1
End If

If (Text10.Text) <= 32 Then
Text31.Text = 0
End If
If (Text11.Text) >= 80 Then
Text32.Text = 5
End If
If (Text11.Text) <= 79 Then
Text32.Text = 4
End If

If (Text11.Text) <= 69 Then
Text32.Text = 3.5
End If

If (Text11.Text) <= 59 Then
Text32.Text = 3
End If

If (Text11.Text) <= 49 Then
Text32.Text = 2
End If

If (Text11.Text) <= 39 Then
Text32.Text = 1
End If

If (Text11.Text) <= 32 Then
Text32.Text = 0
End If
If (Text12.Text) >= 80 Then
Text33.Text = 5
End If
If (Text12.Text) <= 79 Then
Text33.Text = 4
End If

If (Text12.Text) <= 69 Then
Text33.Text = 3.5
End If

If (Text12.Text) <= 59 Then
Text33.Text = 3
End If

If (Text12.Text) <= 49 Then
Text33.Text = 2
End If

If (Text12.Text) <= 39 Then
Text33.Text = 1
End If

If (Text12.Text) <= 32 Then
Text33.Text = 0
End If
If (Text111.Text) >= 80 Then
Text34.Text = 5
End If
If (Text111.Text) <= 79 Then
Text34.Text = 4
End If

If (Text111.Text) <= 69 Then
Text34.Text = 3.5
End If

If (Text111.Text) <= 59 Then
Text34.Text = 3
End If

If (Text111.Text) <= 49 Then
Text34.Text = 2
End If

If (Text111.Text) <= 39 Then
Text34.Text = 1
End If

If (Text111.Text) <= 32 Then
Text34.Text = 0
End If
If (Text16.Text) >= 80 Then
Text35.Text = 5
End If
If (Text16.Text) <= 79 Then
Text35.Text = 4
End If

If (Text16.Text) <= 69 Then
Text35.Text = 3.5
End If

If (Text16.Text) <= 59 Then
Text35.Text = 3
End If

If (Text16.Text) <= 49 Then
Text35.Text = 2
End If

If (Text16.Text) <= 39 Then
Text35.Text = 1
End If

If (Text16.Text) <= 32 Then
Text35.Text = 0
End If

End Sub



