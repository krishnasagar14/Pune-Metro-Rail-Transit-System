VERSION 5.00
Begin VB.Form Rules 
   BackColor       =   &H00FFFF00&
   Caption         =   "Rules"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label6 
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
      Left            =   360
      TabIndex        =   5
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C000&
      Caption         =   "Go back on fron page click here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      TabIndex        =   4
      Top             =   8640
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   4185
      Left            =   8400
      Picture         =   "Rules.frx":0000
      Stretch         =   -1  'True
      Top             =   960
      Width           =   7500
   End
   Begin VB.Image Image1 
      Height          =   4155
      Left            =   2040
      Picture         =   "Rules.frx":A426
      Stretch         =   -1  'True
      Top             =   960
      Width           =   6450
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "3.If any one find in illegale activity during journey then immediately inform the near by police or call on this number 100."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   7680
      Width           =   13575
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   $"Rules.frx":FC62
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1920
      TabIndex        =   2
      Top             =   6360
      Width           =   13575
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "1.Once Ticket has been genrated is not liable to cancel or refundable."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   5880
      Width           =   13575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rules"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   7800
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "Rules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label5_Click()
A_Front_page.Visible = True
Rules.Visible = False
End Sub
