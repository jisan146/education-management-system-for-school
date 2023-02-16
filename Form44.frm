VERSION 5.00
Begin VB.Form Form44 
   Caption         =   "Result's of Seven"
   ClientHeight    =   7755
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form44"
   ScaleHeight     =   7755
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text07 
      Height          =   375
      Left            =   6840
      TabIndex        =   36
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox Text05 
      Height          =   375
      Left            =   6840
      TabIndex        =   35
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text06 
      Height          =   375
      Left            =   6840
      TabIndex        =   34
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox Text04 
      Height          =   375
      Left            =   6840
      TabIndex        =   33
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text03 
      Height          =   375
      Left            =   4320
      TabIndex        =   32
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox Text02 
      Height          =   375
      Left            =   4320
      TabIndex        =   31
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox Text01 
      Height          =   375
      Left            =   4320
      TabIndex        =   30
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text0 
      Height          =   375
      Left            =   4320
      TabIndex        =   29
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text19 
      Height          =   375
      Left            =   7320
      TabIndex        =   28
      Top             =   6480
      Width           =   735
   End
   Begin VB.TextBox Text20 
      Height          =   375
      Left            =   4440
      TabIndex        =   27
      Top             =   6480
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3000
      TabIndex        =   26
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   360
      TabIndex        =   25
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   5160
      TabIndex        =   24
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   5160
      TabIndex        =   23
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   5160
      TabIndex        =   22
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   5160
      TabIndex        =   21
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3000
      TabIndex        =   20
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3000
      TabIndex        =   19
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   3480
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
      TabIndex        =   14
      Text            =   "Select ID"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox Text09 
      Height          =   375
      Left            =   9360
      TabIndex        =   12
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text010 
      Height          =   375
      Left            =   9360
      TabIndex        =   11
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox Text08 
      Height          =   375
      Left            =   9360
      TabIndex        =   10
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   7680
      TabIndex        =   9
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   7680
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
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
      Left            =   9240
      TabIndex        =   4
      Top             =   120
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
      Left            =   9240
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Fomat to GPA"
      Height          =   375
      Left            =   9240
      TabIndex        =   2
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Total"
      Height          =   375
      Left            =   9240
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   9240
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
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
      Left            =   6840
      TabIndex        =   58
      Top             =   1920
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
      Left            =   4320
      TabIndex        =   57
      Top             =   1920
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
      Left            =   7320
      TabIndex        =   56
      Top             =   6000
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
      Left            =   4440
      TabIndex        =   55
      Top             =   6000
      Width           =   2655
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
      Left            =   2520
      TabIndex        =   54
      Top             =   6000
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
      Left            =   3000
      TabIndex        =   53
      Top             =   1920
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
      Left            =   3000
      TabIndex        =   52
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Genaral Science"
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
      TabIndex        =   51
      Top             =   4800
      Width           =   1815
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
      Left            =   5160
      TabIndex        =   50
      Top             =   3840
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
      Left            =   5160
      TabIndex        =   49
      Top             =   2880
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
      Left            =   5160
      TabIndex        =   48
      Top             =   1920
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
      Left            =   3000
      TabIndex        =   47
      Top             =   4800
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
      Left            =   3000
      TabIndex        =   46
      Top             =   3840
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
      TabIndex        =   45
      Top             =   4920
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
      TabIndex        =   44
      Top             =   3960
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
      TabIndex        =   43
      Top             =   3000
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
      TabIndex        =   42
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "Choose"
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
      TabIndex        =   41
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label13 
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
      Left            =   9360
      TabIndex        =   40
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label14 
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
      Height          =   375
      Left            =   7680
      TabIndex        =   39
      Top             =   3840
      Width           =   1815
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
      Left            =   7680
      TabIndex        =   38
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label19 
      Caption         =   "Agricultute"
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
      Left            =   7680
      TabIndex        =   37
      Top             =   1920
      Width           =   1575
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
      Left            =   2520
      TabIndex        =   6
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Result's of Seven"
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
      Left            =   2400
      TabIndex        =   5
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "Form44"
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
.Open "select * from ResultofSeven where ResultofSeven.StudentID='" & id & "'", cnn, adOpenKeyset, adLockOptimistic
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
Text13.Text = .Fields(1)
Text14.Text = .Fields(2)
Text15.Text = .Fields(3)
Text17.Text = .Fields(15)
End With

cnn2.Close
rs2.Close
End Sub

Private Sub Command1_Click()
Form44.Hide
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
 Text16.Text = (Text1.Text - -Text2.Text - -Text3.Text - -Text4.Text - -Text5.Text - -Text6.Text - -Text7.Text - -Text8.Text - -Text9.Text - -Text10.Text)
Text19.Text = (Text0.Text - -Text01.Text - -Text02.Text - -Text03.Text - -Text04.Text - -Text05.Text - -Text06.Text - -Text07.Text - -Text08.Text - -Text09.Text - -Text010.Text) / 9


End Sub

Private Sub Form_Load()
With cnn
.Provider = "Microsoft.Jet.OLEDB.4.0"
.ConnectionString = "C:\Project\TITASH GAS HIGH SCHOOL.mdb"
.Open
End With
With rs
.Open "Select ResultofSeven.StudentID from ResultofSeven where ResultofSeven.Class = 'Seven'", cnn, adOpenKeyset, adLockOptimistic
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
Text0.Text = 5
End If
If (Text1.Text) <= 79 Then
Text0.Text = 4
End If

If (Text1.Text) <= 69 Then
Text0.Text = 3.5
End If

If (Text1.Text) <= 59 Then
Text0.Text = 3
End If

If (Text1.Text) <= 49 Then
Text0.Text = 2
End If

If (Text1.Text) <= 39 Then
Text0.Text = 1
End If

If (Text1.Text) <= 32 Then
Text0.Text = 0
End If
If (Text2.Text) >= 80 Then
Text01.Text = 5
End If
If (Text2.Text) <= 79 Then
Text01.Text = 4
End If

If (Text2.Text) <= 69 Then
Text01.Text = 3.5
End If

If (Text2.Text) <= 59 Then
Text01.Text = 3
End If

If (Text2.Text) <= 49 Then
Text01.Text = 2
End If

If (Text2.Text) <= 39 Then
Text01.Text = 1
End If

If (Text2.Text) <= 32 Then
Text01.Text = 0
End If
If (Text3.Text) >= 80 Then
Text02.Text = 5
End If
If (Text3.Text) <= 79 Then
Text02.Text = 4
End If

If (Text3.Text) <= 69 Then
Text02.Text = 3.5
End If

If (Text3.Text) <= 59 Then
Text02.Text = 3
End If

If (Text3.Text) <= 49 Then
Text02.Text = 2
End If

If (Text3.Text) <= 39 Then
Text02.Text = 1
End If

If (Text3.Text) <= 32 Then
Text02.Text = 0
End If
If (Text4.Text) >= 80 Then
Text03.Text = 5
End If
If (Text4.Text) <= 79 Then
Text03.Text = 4
End If

If (Text4.Text) <= 69 Then
Text03.Text = 3.5
End If

If (Text4.Text) <= 59 Then
Text03.Text = 3
End If

If (Text4.Text) <= 49 Then
Text03.Text = 2
End If

If (Text4.Text) <= 39 Then
Text03.Text = 1
End If

If (Text4.Text) <= 32 Then
Text03.Text = 0
End If
If (Text5.Text) >= 80 Then
Text04.Text = 5
End If
If (Text5.Text) <= 79 Then
Text04.Text = 4
End If

If (Text5.Text) <= 69 Then
Text04.Text = 3.5
End If

If (Text5.Text) <= 59 Then
Text04.Text = 3
End If

If (Text5.Text) <= 49 Then
Text04.Text = 2
End If

If (Text5.Text) <= 39 Then
Text04.Text = 1
End If

If (Text5.Text) <= 32 Then
Text04.Text = 0
End If
If (Text6.Text) >= 80 Then
Text05.Text = 5
End If
If (Text6.Text) <= 79 Then
Text05.Text = 4
End If

If (Text6.Text) <= 69 Then
Text05.Text = 3.5
End If

If (Text6.Text) <= 59 Then
Text05.Text = 3
End If

If (Text6.Text) <= 49 Then
Text05.Text = 2
End If

If (Text6.Text) <= 39 Then
Text05.Text = 1
End If

If (Text6.Text) <= 32 Then
Text05.Text = 0
End If
If (Text7.Text) >= 80 Then
Text06.Text = 5
End If
If (Text7.Text) <= 79 Then
Text06.Text = 4
End If

If (Text7.Text) <= 69 Then
Text06.Text = 3.5
End If

If (Text7.Text) <= 59 Then
Text06.Text = 3
End If

If (Text7.Text) <= 49 Then
Text06.Text = 2
End If

If (Text7.Text) <= 39 Then
Text06.Text = 1
End If

If (Text8.Text) <= 32 Then
Text06.Text = 0
End If
If (Text8.Text) >= 80 Then
Text07.Text = 5
End If
If (Text8.Text) <= 79 Then
Text07.Text = 4
End If

If (Text8.Text) <= 69 Then
Text07.Text = 3.5
End If

If (Text8.Text) <= 59 Then
Text07.Text = 3
End If

If (Text8.Text) <= 49 Then
Text07.Text = 2
End If

If (Text8.Text) <= 39 Then
Text07.Text = 1
End If

If (Text8.Text) <= 32 Then
Text07.Text = 0
End If

If (Text9.Text) >= 80 Then
Text08.Text = 5
End If
If (Text9.Text) <= 79 Then
Text08.Text = 4
End If

If (Text9.Text) <= 69 Then
Text08.Text = 3.5
End If

If (Text9.Text) <= 59 Then
Text08.Text = 3
End If

If (Text9.Text) <= 49 Then
Text08.Text = 2
End If

If (Text9.Text) <= 39 Then
Text08.Text = 1
End If

If (Text9.Text) <= 32 Then
Text08.Text = 0
End If
If (Text10.Text) >= 80 Then
Text09.Text = 5
End If
If (Text10.Text) <= 79 Then
Text09.Text = 4
End If

If (Text10.Text) <= 69 Then
Text09.Text = 3.5
End If

If (Text10.Text) <= 59 Then
Text09.Text = 3
End If

If (Text10.Text) <= 49 Then
Text09.Text = 2
End If

If (Text10.Text) <= 39 Then
Text09.Text = 1
End If

If (Text10.Text) <= 32 Then
Text09.Text = 0
End If
If (Text11.Text) >= 80 Then
Text010.Text = 5
End If
If (Text11.Text) <= 79 Then
Text010.Text = 4
End If

If (Text11.Text) <= 69 Then
Text010.Text = 3.5
End If

If (Text11.Text) <= 59 Then
Text010.Text = 3
End If

If (Text11.Text) <= 49 Then
Text010.Text = 2
End If

If (Text11.Text) <= 39 Then
Text010.Text = 1
End If

If (Text11.Text) <= 32 Then
Text010.Text = 0
End If





End Sub







