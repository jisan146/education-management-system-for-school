VERSION 5.00
Begin VB.Form Form40 
   Caption         =   "Result's of Three"
   ClientHeight    =   7845
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11355
   LinkTopic       =   "Form40"
   ScaleHeight     =   7845
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text014 
      Height          =   375
      Left            =   7440
      TabIndex        =   42
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox Text013 
      Height          =   375
      Left            =   7440
      TabIndex        =   41
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   7440
      TabIndex        =   39
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   5160
      TabIndex        =   38
      Top             =   5520
      Width           =   735
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
      Left            =   1320
      TabIndex        =   23
      Text            =   "Select ID"
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   1320
      TabIndex        =   22
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   1320
      TabIndex        =   21
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   1320
      TabIndex        =   20
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3840
      TabIndex        =   18
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6120
      TabIndex        =   17
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3840
      TabIndex        =   15
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   4080
      TabIndex        =   14
      Top             =   6960
      Width           =   2655
   End
   Begin VB.TextBox Text18 
      Height          =   375
      Left            =   6960
      TabIndex        =   13
      Top             =   6960
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   10080
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Total"
      Height          =   375
      Left            =   10080
      TabIndex        =   3
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Fomat to GPA"
      Height          =   375
      Left            =   10080
      TabIndex        =   2
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
      Left            =   10080
      TabIndex        =   1
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
      Left            =   10080
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "English 2"
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
      Left            =   3840
      TabIndex        =   43
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Bengali 2"
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
      Left            =   3840
      TabIndex        =   40
      Top             =   3000
      Width           =   1215
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
      Left            =   1320
      TabIndex        =   37
      Top             =   2040
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
      Left            =   1320
      TabIndex        =   36
      Top             =   3120
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
      Left            =   1320
      TabIndex        =   35
      Top             =   4080
      Width           =   1575
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
      Left            =   1320
      TabIndex        =   34
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "English 1"
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
      Left            =   3840
      TabIndex        =   33
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Math"
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
      Left            =   6120
      TabIndex        =   32
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Genaral Knowledge"
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
      Left            =   6120
      TabIndex        =   31
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Bengali 1"
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
      Left            =   3840
      TabIndex        =   30
      Top             =   2040
      Width           =   1215
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
      Left            =   2160
      TabIndex        =   29
      Top             =   6480
      Width           =   1575
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
      Left            =   4080
      TabIndex        =   28
      Top             =   6480
      Width           =   2655
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
      Left            =   6960
      TabIndex        =   27
      Top             =   6480
      Width           =   1335
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
      Left            =   5280
      TabIndex        =   26
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Social Science"
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
      Left            =   6240
      TabIndex        =   25
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label8 
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
      Left            =   7560
      TabIndex        =   24
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Result's of Three"
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
      Left            =   3240
      TabIndex        =   5
      Top             =   0
      Width           =   4335
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
      Left            =   3360
      TabIndex        =   4
      Top             =   840
      Width           =   4095
   End
End
Attribute VB_Name = "Form40"
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
.Open "select * from ResultofThree where ResultofThree.StudentID='" & id & "'", cnn, adOpenKeyset, adLockOptimistic
End With
With rs2
Text1.Text = .Fields(4)
Text2.Text = .Fields(5)
Text3.Text = .Fields(6)
Text4.Text = .Fields(7)
Text5.Text = .Fields(8)
Text6.Text = .Fields(9)

Text7.Text = .Fields(10)

Text13.Text = .Fields(1)
Text14.Text = .Fields(2)
Text15.Text = .Fields(3)

End With

cnn2.Close
rs2.Close
End Sub

Private Sub Command1_Click()
Form40.Hide
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
 Text16.Text = (Text1.Text - -Text2.Text - -Text3.Text - -Text4.Text - -Text5.Text - -Text6.Text - -Text7.Text)
Text18.Text = (Text8.Text - -Text9.Text - -Text10.Text - -Text11.Text - -Text12.Text - -Text013.Text - -Text014.Text) / 8


End Sub

Private Sub Form_Load()
With cnn
.Provider = "Microsoft.Jet.OLEDB.4.0"
.ConnectionString = "C:\Project\TITASH GAS HIGH SCHOOL.mdb"
.Open
End With
With rs
.Open "Select ResultofThree.StudentID from ResultofThree where ResultofThree.Class = 'Three'", cnn, adOpenKeyset, adLockOptimistic
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
Text8.Text = 5
End If
If (Text1.Text) <= 79 Then
Text8.Text = 4
End If

If (Text1.Text) <= 69 Then
Text8.Text = 3.5
End If

If (Text1.Text) <= 59 Then
Text8.Text = 3
End If

If (Text1.Text) <= 49 Then
Text8.Text = 2
End If

If (Text1.Text) <= 39 Then
Text8.Text = 1
End If

If (Text1.Text) <= 32 Then
Text8.Text = 0
End If
If (Text2.Text) >= 80 Then
Text9.Text = 5
End If
If (Text2.Text) <= 79 Then
Text9.Text = 4
End If

If (Text2.Text) <= 69 Then
Text9.Text = 3.5
End If

If (Text2.Text) <= 59 Then
Text9.Text = 3
End If

If (Text2.Text) <= 49 Then
Text9.Text = 2
End If

If (Text2.Text) <= 39 Then
Text9.Text = 1
End If

If (Text2.Text) <= 32 Then
Text9.Text = 0
End If
If (Text3.Text) >= 80 Then
Text10.Text = 5
End If
If (Text3.Text) <= 79 Then
Text10.Text = 4
End If

If (Text3.Text) <= 69 Then
Text10.Text = 3.5
End If

If (Text3.Text) <= 59 Then
Text10.Text = 3
End If

If (Text3.Text) <= 49 Then
Text10.Text = 2
End If

If (Text3.Text) <= 39 Then
Text10.Text = 1
End If

If (Text3.Text) <= 32 Then
Text10.Text = 0
End If
If (Text4.Text) >= 80 Then
Text11.Text = 5
End If
If (Text4.Text) <= 79 Then
Text11.Text = 4
End If

If (Text4.Text) <= 69 Then
Text11.Text = 3.5
End If

If (Text4.Text) <= 59 Then
Text11.Text = 3
End If

If (Text4.Text) <= 49 Then
Text11.Text = 2
End If

If (Text4.Text) <= 39 Then
Text11.Text = 1
End If

If (Text4.Text) <= 32 Then
Text11.Text = 0
End If
If (Text5.Text) >= 80 Then
Text12.Text = 5
End If
If (Text5.Text) <= 79 Then
Text12.Text = 4
End If

If (Text5.Text) <= 69 Then
Text12.Text = 3.5
End If

If (Text5.Text) <= 59 Then
Text12.Text = 3
End If

If (Text5.Text) <= 49 Then
Text12.Text = 2
End If

If (Text5.Text) <= 39 Then
Text12.Text = 1
End If

If (Text5.Text) <= 32 Then
Text12.Text = 0
End If
If (Text6.Text) >= 80 Then
Text013.Text = 5
End If
If (Text6.Text) <= 79 Then
Text013.Text = 4
End If

If (Text6.Text) <= 69 Then
Text013.Text = 3.5
End If

If (Text6.Text) <= 59 Then
Text013.Text = 3
End If

If (Text6.Text) <= 49 Then
Text013.Text = 2
End If

If (Text6.Text) <= 39 Then
Text013.Text = 1
End If

If (Text6.Text) <= 32 Then
Text013.Text = 0
End If
If (Text7.Text) >= 80 Then
Text014.Text = 5
End If
If (Text7.Text) <= 79 Then
Text014.Text = 4
End If

If (Text7.Text) <= 69 Then
Text014.Text = 3.5
End If

If (Text7.Text) <= 59 Then
Text014.Text = 3
End If

If (Text7.Text) <= 49 Then
Text014.Text = 2
End If

If (Text7.Text) <= 39 Then
Text014.Text = 1
End If

If (Text7.Text) <= 32 Then
Text014.Text = 0
End If







End Sub





