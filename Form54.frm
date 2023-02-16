VERSION 5.00
Begin VB.Form Form54 
   Caption         =   "Result's of Pre-Test (Human)"
   ClientHeight    =   8490
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11535
   LinkTopic       =   "Form54"
   ScaleHeight     =   8490
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text36 
      Height          =   375
      Left            =   10440
      TabIndex        =   69
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox Text34 
      Height          =   375
      Left            =   8880
      TabIndex        =   68
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Total"
      Height          =   375
      Left            =   10080
      TabIndex        =   39
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Fomat to GPA"
      Height          =   375
      Left            =   10080
      TabIndex        =   38
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text33 
      Height          =   375
      Left            =   8040
      TabIndex        =   37
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox Text32 
      Height          =   375
      Left            =   8040
      TabIndex        =   36
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox Text31 
      Height          =   375
      Left            =   8040
      TabIndex        =   35
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text30 
      Height          =   375
      Left            =   8040
      TabIndex        =   34
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text29 
      Height          =   375
      Left            =   5520
      TabIndex        =   33
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox Text28 
      Height          =   375
      Left            =   5520
      TabIndex        =   32
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox Text27 
      Height          =   375
      Left            =   5520
      TabIndex        =   31
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text26 
      Height          =   375
      Left            =   5520
      TabIndex        =   30
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text25 
      Height          =   375
      Left            =   3000
      TabIndex        =   29
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox Text24 
      Height          =   375
      Left            =   3000
      TabIndex        =   28
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox Text23 
      Height          =   375
      Left            =   3000
      TabIndex        =   27
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text22 
      Height          =   375
      Left            =   3000
      TabIndex        =   26
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text21 
      Height          =   375
      Left            =   4800
      TabIndex        =   25
      Top             =   7680
      Width           =   735
   End
   Begin VB.TextBox Text20 
      Height          =   375
      Left            =   1920
      TabIndex        =   24
      Top             =   7680
      Width           =   2655
   End
   Begin VB.TextBox Text19 
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   6720
      Width           =   1575
   End
   Begin VB.TextBox Text18 
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   21
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   7680
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   6360
      TabIndex        =   19
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   6360
      TabIndex        =   18
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   6360
      TabIndex        =   17
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3840
      TabIndex        =   16
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   3840
      TabIndex        =   15
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   6360
      TabIndex        =   14
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2880
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
      Left            =   0
      TabIndex        =   5
      Text            =   "Select ID"
      Top             =   1800
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
      Left            =   10080
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   10080
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   8880
      TabIndex        =   1
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox Text35 
      Height          =   375
      Left            =   10440
      TabIndex        =   0
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "Geography"
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
      Left            =   8880
      TabIndex        =   70
      Top             =   5280
      Width           =   1575
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
      Left            =   10440
      TabIndex        =   67
      Top             =   4320
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
      Left            =   8040
      TabIndex        =   66
      Top             =   2400
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
      Left            =   5520
      TabIndex        =   65
      Top             =   2400
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
      Left            =   3000
      TabIndex        =   64
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
      Left            =   4800
      TabIndex        =   63
      Top             =   7200
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
      Left            =   1920
      TabIndex        =   62
      Top             =   7200
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
      Left            =   0
      TabIndex        =   61
      Top             =   6240
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
      Left            =   0
      TabIndex        =   60
      Top             =   5280
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
      Left            =   0
      TabIndex        =   59
      Top             =   7200
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
      Left            =   6360
      TabIndex        =   58
      Top             =   3360
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
      Left            =   1680
      TabIndex        =   57
      Top             =   2400
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
      Left            =   1680
      TabIndex        =   56
      Top             =   3360
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
      Left            =   6360
      TabIndex        =   55
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label14 
      Caption         =   "Civis"
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
      TabIndex        =   54
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label13 
      Caption         =   "History"
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
      TabIndex        =   53
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "Economics"
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
      TabIndex        =   52
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label11 
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
      Left            =   3840
      TabIndex        =   51
      Top             =   5280
      Width           =   1815
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
      Left            =   6360
      TabIndex        =   50
      Top             =   2400
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
      Left            =   3840
      TabIndex        =   49
      Top             =   2400
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
      Left            =   1680
      TabIndex        =   48
      Top             =   5280
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
      Left            =   1680
      TabIndex        =   47
      Top             =   4320
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
      Left            =   0
      TabIndex        =   46
      Top             =   4320
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
      Left            =   0
      TabIndex        =   45
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Result's of Pre-Test (Human)"
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
      TabIndex        =   44
      Top             =   240
      Width           =   5055
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
      TabIndex        =   43
      Top             =   2400
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
      Left            =   0
      TabIndex        =   42
      Top             =   1320
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
      Left            =   3360
      TabIndex        =   41
      Top             =   1080
      Width           =   4095
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
      Left            =   8880
      TabIndex        =   40
      Top             =   4080
      Width           =   1455
   End
End
Attribute VB_Name = "Form54"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
Command5.Visible = True
End Sub
