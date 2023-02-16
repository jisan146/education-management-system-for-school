VERSION 5.00
Begin VB.Form Form33 
   Caption         =   "Due Payment Infomation of Class Ten (Business)"
   ClientHeight    =   9060
   ClientLeft      =   1200
   ClientTop       =   495
   ClientWidth     =   11385
   LinkTopic       =   "Form33"
   Picture         =   "Form33.frx":0000
   ScaleHeight     =   9060
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text16 
      Height          =   495
      Left            =   9240
      TabIndex        =   44
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text17 
      Height          =   495
      Left            =   9240
      TabIndex        =   43
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   42
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   41
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   40
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   39
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   9840
      TabIndex        =   38
      Top             =   960
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
      Left            =   9840
      TabIndex        =   31
      Top             =   0
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
      Left            =   9840
      TabIndex        =   30
      Top             =   480
      Width           =   1215
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
      Left            =   0
      TabIndex        =   29
      Text            =   "Select ID"
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Monthly Due Payment"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2280
      TabIndex        =   16
      Top             =   4920
      Width           =   8895
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1560
         TabIndex        =   21
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3000
         TabIndex        =   20
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   4560
         TabIndex        =   19
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   6000
         TabIndex        =   18
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   7440
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "January"
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
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "February"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   27
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "March"
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
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "April"
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
         Left            =   4800
         TabIndex        =   25
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "May"
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
         Left            =   6240
         TabIndex        =   24
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "June"
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
         TabIndex        =   23
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Monthly Due Payment"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2280
      TabIndex        =   3
      Top             =   6840
      Width           =   8895
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   4560
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   6000
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   7440
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "July"
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
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "August"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "September"
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
         Left            =   3120
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "October"
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
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "November"
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
         Left            =   6000
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "December"
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
         Left            =   7440
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7560
      Width           =   1455
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Label Label19 
      Height          =   495
      Left            =   9240
      TabIndex        =   46
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label20 
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      TabIndex        =   45
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Please Select Student ID to see Due payment                                   Information"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   37
      Top             =   960
      Width           =   5055
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
      Left            =   0
      TabIndex        =   36
      Top             =   5040
      Width           =   1455
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
      Left            =   0
      TabIndex        =   35
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Due Payment Infomation of Class Ten (Business)"
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
      Left            =   1200
      TabIndex        =   34
      Top             =   120
      Width           =   8535
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
      Left            =   0
      TabIndex        =   33
      Top             =   7080
      Width           =   1455
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
      Left            =   0
      TabIndex        =   32
      Top             =   8040
      Width           =   1455
   End
End
Attribute VB_Name = "Form33"
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
.Open "select * from PaymentTenBusiness where PaymentTenBusiness.StudentID='" & id & "'", cnn, adOpenKeyset, adLockOptimistic
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
Text13.Text = .Fields(1)
Text14.Text = .Fields(2)
Text15.Text = .Fields(3)

End With
cnn2.Close
rs2.Close
End Sub

Private Sub Command1_Click()
Form33.Hide
Form4.Show
End Sub
Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form33.Hide
Form1.Show
End Sub

Private Sub Form_Load()
With cnn
.Provider = "Microsoft.Jet.OLEDB.4.0"
.ConnectionString = "C:\Project\TITASH GAS HIGH SCHOOL.mdb"
.Open
End With
With rs
.Open "Select PaymentTenBusiness.StudentID from PaymentTenBusiness where PaymentTenBusiness.Class = 'Ten(Business)'", cnn, adOpenKeyset, adLockOptimistic
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
Label19.Caption = (Text16.Text - -Text17.Text)
End Sub

Private Sub Command5_Click()
Label19.Caption = (Text16.Text - Text17.Text)
End Sub

Private Sub Command6_Click()
Label19.Caption = (Text16.Text * Text17.Text)
End Sub

Private Sub Command7_Click()
Label19.Caption = (Text16.Text / Text17.Text)
End Sub

Private Sub Command8_Click()
If Text19.Text = Text18.Text Then
Frame1.Visible = True
Frame2.Visible = True
Label21.Visible = False
Text18.Visible = False
Command8.Visible = False
End If
End Sub








