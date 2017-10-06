VERSION 5.00
Begin VB.Form Report 
   BackColor       =   &H00FFFF00&
   Caption         =   "Reports"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "Logout"
      Height          =   495
      Left            =   6840
      TabIndex        =   8
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Routes"
      Height          =   615
      Left            =   9000
      TabIndex        =   5
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SMS"
      Height          =   615
      Left            =   6840
      TabIndex        =   4
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Account details"
      Height          =   615
      Left            =   4560
      TabIndex        =   3
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Feedabck"
      Height          =   615
      Left            =   9000
      TabIndex        =   2
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "User details"
      Height          =   615
      Left            =   6840
      TabIndex        =   1
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tickets"
      Height          =   615
      Left            =   4560
      TabIndex        =   0
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "    P                      M                   R                    T              S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10335
      Left            =   480
      TabIndex        =   7
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1905
      Left            =   2040
      Picture         =   "Report.frx":0000
      Top             =   840
      Width           =   2805
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Reports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   6
      Top             =   1800
      Width           =   3855
   End
End
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DataReport6.Show
End Sub

Private Sub Command2_Click()
DataReport3.Show
End Sub

Private Sub Command3_Click()
DataReport4.Show
End Sub

Private Sub Command4_Click()
DataReport5.Show
End Sub

Private Sub Command5_Click()
DataReport1.Show
End Sub

Private Sub Command6_Click()
DataReport2.Show
End Sub

Private Sub Command7_Click()
A_Front_page.Show
Unload Me
End Sub

