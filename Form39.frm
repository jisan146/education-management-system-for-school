VERSION 5.00
Begin VB.Form Form39 
   Caption         =   "Result's of Two"
   ClientHeight    =   8280
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   10575
   LinkTopic       =   "Form39"
   ScaleHeight     =   8280
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   6840
      TabIndex        =   35
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3240
      TabIndex        =   34
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Total"
      Height          =   375
      Left            =   9120
      TabIndex        =   19
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Fomat to GPA"
      Height          =   375
      Left            =   9120
      TabIndex        =   18
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   4560
      TabIndex        =   15
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text18 
      Height          =   375
      Left            =   6360
      TabIndex        =   13
      Top             =   7320
      Width           =   735
   End
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   7320
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   7320
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3840
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
      Left            =   720
      TabIndex        =   3
      Text            =   "Select ID"
      Top             =   2880
      Width           =   1575
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
      Left            =   9120
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
      Left            =   9120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   9120
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
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
      Left            =   6960
      TabIndex        =   37
      Top             =   2400
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
      Left            =   5520
      TabIndex        =   36
      Top             =   2040
      Width           =   1455
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
      Left            =   4680
      TabIndex        =   33
      Top             =   2400
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
      Left            =   6360
      TabIndex        =   32
      Top             =   6840
      Width           =   1335
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
      Left            =   3480
      TabIndex        =   31
      Top             =   6840
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
      Left            =   1560
      TabIndex        =   30
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Bengali"
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
      Left            =   3240
      TabIndex        =   29
      Top             =   2400
      Width           =   1215
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
      Left            =   3240
      TabIndex        =   28
      Top             =   5280
      Width           =   2055
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
      Left            =   3240
      TabIndex        =   27
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "English"
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
      Left            =   3240
      TabIndex        =   26
      Top             =   3360
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
      Left            =   720
      TabIndex        =   25
      Top             =   5400
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
      Left            =   720
      TabIndex        =   24
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Result's of Two"
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
      TabIndex        =   23
      Top             =   240
      Width           =   4335
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
      Left            =   720
      TabIndex        =   22
      Top             =   3480
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
      Left            =   720
      TabIndex        =   21
      Top             =   2400
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
      Left            =   2400
      TabIndex        =   20
      Top             =   1080
      Width           =   4095
   End
End
Attribute VB_Name = "Form39"
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
.Open "select * from ResultofTwo where ResultofTwo.StudentID='" & id & "'", cnn, adOpenKeyset, adLockOptimistic
End With
With rs2
Text1.Text = .Fields(4)
Text2.Text = .Fields(5)
Text3.Text = .Fields(6)
Text4.Text = .Fields(7)
Text5.Text = .Fields(8)


Text13.Text = .Fields(1)
Text14.Text = .Fields(2)
Text15.Text = .Fields(3)

End With

cnn2.Close
rs2.Close
End Sub

Private Sub Command1_Click()
Form39.Hide
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
 Text16.Text = (Text1.Text - -Text2.Text - -Text3.Text - -Text4.Text - -Text5.Text)
Text18.Text = (Text6.Text - -Text7.Text - -Text8.Text - -Text9.Text - -Text10.Text) / 5


End Sub

Private Sub Form_Load()
With cnn
.Provider = "Microsoft.Jet.OLEDB.4.0"
.ConnectionString = "C:\Project\TITASH GAS HIGH SCHOOL.mdb"
.Open
End With
With rs
.Open "Select ResultofTwo.StudentID from ResultofTwo where ResultofTwo.Class = 'Two'", cnn, adOpenKeyset, adLockOptimistic
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
If (Text1.Text) >= 80 Then
Text6.Text = 5
End If
If (Text1.Text) <= 79 Then
Text6.Text = 4
End If

If (Text1.Text) <= 69 Then
Text6.Text = 3.5
End If

If (Text1.Text) <= 59 Then
Text6.Text = 3
End If

If (Text1.Text) <= 49 Then
Text6.Text = 2
End If

If (Text1.Text) <= 39 Then
Text6.Text = 1
End If

If (Text1.Text) <= 32 Then
Text6.Text = 0
End If
If (Text2.Text) >= 80 Then
Text7.Text = 5
End If
If (Text2.Text) <= 79 Then
Text7.Text = 4
End If

If (Text2.Text) <= 69 Then
Text7.Text = 3.5
End If

If (Text2.Text) <= 59 Then
Text7.Text = 3
End If

If (Text2.Text) <= 49 Then
Text7.Text = 2
End If

If (Text2.Text) <= 39 Then
Text7.Text = 1
End If

If (Text2.Text) <= 32 Then
Text7.Text = 0
End If
If (Text3.Text) >= 80 Then
Text8.Text = 5
End If
If (Text3.Text) <= 79 Then
Text8.Text = 4
End If

If (Text3.Text) <= 69 Then
Text8.Text = 3.5
End If

If (Text3.Text) <= 59 Then
Text8.Text = 3
End If

If (Text3.Text) <= 49 Then
Text8.Text = 2
End If

If (Text3.Text) <= 39 Then
Text8.Text = 1
End If

If (Text3.Text) <= 32 Then
Text8.Text = 0
End If
If (Text4.Text) >= 80 Then
Text9.Text = 5
End If
If (Text4.Text) <= 79 Then
Text9.Text = 4
End If

If (Text4.Text) <= 69 Then
Text9.Text = 3.5
End If

If (Text4.Text) <= 59 Then
Text9.Text = 3
End If

If (Text4.Text) <= 49 Then
Text9.Text = 2
End If

If (Text4.Text) <= 39 Then
Text9.Text = 1
End If

If (Text4.Text) <= 32 Then
Text9.Text = 0
End If
If (Text5.Text) >= 80 Then
Text10.Text = 5
End If
If (Text5.Text) <= 79 Then
Text10.Text = 4
End If

If (Text5.Text) <= 69 Then
Text10.Text = 3.5
End If

If (Text5.Text) <= 59 Then
Text10.Text = 3
End If

If (Text5.Text) <= 49 Then
Text10.Text = 2
End If

If (Text5.Text) <= 39 Then
Text10.Text = 1
End If

If (Text5.Text) <= 32 Then
Text10.Text = 0
End If
Command5.Visible = True
End Sub




