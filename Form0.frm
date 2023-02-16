VERSION 5.00
Begin VB.Form Form0 
   Caption         =   "ASAT Software"
   ClientHeight    =   6750
   ClientLeft      =   960
   ClientTop       =   495
   ClientWidth     =   8265
   Icon            =   "Form0.frx":0000
   LinkTopic       =   "Form37"
   Picture         =   "Form0.frx":0442
   ScaleHeight     =   6750
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
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
      Left            =   1560
      TabIndex        =   19
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6000
      TabIndex        =   18
      Text            =   "ghdfhhtydsgdfhdghsdg"
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2400
      TabIndex        =   17
      Top             =   6480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   7080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2400
      TabIndex        =   15
      Top             =   7680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   5280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   7320
      TabIndex        =   13
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      Top             =   6480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   7080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   7680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   7320
      TabIndex        =   9
      Top             =   8280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   4680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   9360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Text00 
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label18 
      Caption         =   "User"
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
      Left            =   2160
      TabIndex        =   35
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label17 
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
      Left            =   600
      TabIndex        =   34
      Top             =   4080
      Visible         =   0   'False
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
      Left            =   600
      TabIndex        =   33
      Top             =   5280
      Visible         =   0   'False
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
      Left            =   600
      TabIndex        =   32
      Top             =   5880
      Visible         =   0   'False
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
      Left            =   600
      TabIndex        =   31
      Top             =   6480
      Visible         =   0   'False
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
      Left            =   600
      TabIndex        =   30
      Top             =   7080
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   600
      TabIndex        =   29
      Top             =   7680
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   5760
      TabIndex        =   28
      Top             =   5280
      Visible         =   0   'False
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
      Left            =   5760
      TabIndex        =   27
      Top             =   6480
      Visible         =   0   'False
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
      Left            =   5760
      TabIndex        =   26
      Top             =   7080
      Visible         =   0   'False
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
      Left            =   5760
      TabIndex        =   25
      Top             =   7680
      Visible         =   0   'False
      Width           =   1455
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
      Left            =   5760
      TabIndex        =   24
      Top             =   8280
      Visible         =   0   'False
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
      Left            =   600
      TabIndex        =   23
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   5760
      TabIndex        =   22
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
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
      Left            =   5760
      TabIndex        =   21
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
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
      Left            =   600
      TabIndex        =   20
      Top             =   8280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
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
      Left            =   4200
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "ASAT Software"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Form0"
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
.Open "select * from Customize Where Customize.UserName='" & id & "'", cnn, adOpenKeyset, adLockOptimistic
End With
With rs2

Text1.Text = .Fields(1)








End With
cnn2.Close
rs2.Close
End Sub
Private Sub Command_Click()
Unload Me
Form2.Show

End Sub
Private Sub command2_click()
End
End Sub

Private Sub Command3_Click()
Form35.Hide
Form1.Show
End Sub

Private Sub Command00_Click()
Form99.Show
End Sub

Private Sub Form_Load()
With cnn
.Provider = "Microsoft.Jet.OLEDB.4.0"
.ConnectionString = "C:\Project\TITASH GAS HIGH SCHOOL.mdb"
.Open
End With
With rs
.Open "Select Customize.UserName from Customize", cnn, adOpenKeyset, adLockOptimistic
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




Private Sub command1_click()
If Text00.Text = Text1.Text Then
Form0.Hide
Form1.Show
End If
End Sub


