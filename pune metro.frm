VERSION 5.00
Begin VB.Form A_Front_page 
   BackColor       =   &H00FFFF80&
   Caption         =   "Pune metro"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      Caption         =   "Admin Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6600
      Top             =   6720
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Feedback"
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
      Left            =   16800
      TabIndex        =   6
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Rules"
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
      Left            =   13800
      TabIndex        =   5
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SMS Service"
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
      Left            =   10800
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Fare Enquiry"
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
      Left            =   8040
      TabIndex        =   3
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Tickets"
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
      Left            =   5400
      TabIndex        =   2
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Metro Information"
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
      Left            =   600
      MaskColor       =   &H00FFFF80&
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Pune Metro Rail"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   10
      TabIndex        =   11
      Top             =   0
      Width           =   5535
   End
   Begin VB.Image Image4 
      Height          =   1905
      Left            =   15840
      Picture         =   "pune metro.frx":0000
      Top             =   7440
      Width           =   2805
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      Caption         =   "Create Your Account here:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15840
      TabIndex        =   10
      Top             =   9360
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FF00&
      Caption         =   "Interactive Route Map of Pune Metro Rail.  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      TabIndex        =   9
      Top             =   6600
      Width           =   3375
   End
   Begin VB.Image Image3 
      Height          =   4080
      Left            =   10440
      Picture         =   "pune metro.frx":255E
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   5400
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   $"pune metro.frx":E339
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   600
      TabIndex        =   8
      Top             =   7200
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "About Pune Metro Project"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   7
      Top             =   6600
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   5130
      Left            =   10560
      Picture         =   "pune metro.frx":E419
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   8490
   End
   Begin VB.Image Image1 
      Height          =   5115
      Left            =   600
      Picture         =   "pune metro.frx":14F02
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   10050
   End
End
Attribute VB_Name = "A_Front_page"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
metro_information.Visible = True
A_Front_page.Visible = False
End Sub

Private Sub Command2_Click()
Metro_timing.Visible = True
A_Front_page.Visible = False
End Sub

Private Sub Command3_Click()
Tickets.Visible = True
A_Front_page.Visible = False
End Sub

Private Sub Command4_Click()
SMS.Visible = True
A_Front_page.Visible = False
End Sub

Private Sub Command5_Click()
Fare_Enquiry.Show
Unload Me
End Sub

Private Sub Command6_Click()
Rules.Visible = True
A_Front_page.Visible = False
End Sub

Private Sub Command7_Click()
Feedback.Show
Unload Me
End Sub

Private Sub Command8_Click()
Admin.Show
Unload Me
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Label3_Click()
SMS.Visible = True
End Sub

Private Sub Label6_Click()
SIGN_UP.Visible = True
A_Front_page.Visible = False

End Sub

Private Sub Timer1_Timer()
If Label1.Left = 10 And Label1.Left <= 12000 Then
Label1.Left = Label1.Left + 100
Else
Label1.Left = 10
End If
End Sub
