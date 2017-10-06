VERSION 5.00
Begin VB.Form metro_information 
   BackColor       =   &H0080FF80&
   Caption         =   "Metro information"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label12 
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
      TabIndex        =   11
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000FF00&
      Caption         =   "Go Back to front page click here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   10
      Top             =   7560
      Width           =   3735
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      Caption         =   "ABOUT US"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1905
      Left            =   1800
      Picture         =   "form2.frx":0000
      Top             =   840
      Width           =   2805
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF8080&
      Caption         =   "* Employees should discharge their responsibilities with pride, perfection and dignity"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   6480
      Width           =   13335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      Caption         =   "* Safety of Metro users is our paramount responsibility."
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   6240
      Width           =   13335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      Caption         =   "* All our structures should be aesthetically planned and well maintained."
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   6000
      Width           =   13335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      Caption         =   "* The Organization must be lean but effective."
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   5760
      Width           =   13335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      Caption         =   "* Personal integrity should never be in doubt, we should maintain full transparency in all our decisions and transactions."
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   5520
      Width           =   13335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "* We should be totally dedicated and committed to the Corporate Mission."
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   5280
      Width           =   13335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Our Corporate Culture"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   5040
      Width           =   13335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   $"form2.frx":255E
      Height          =   1215
      Left            =   1560
      TabIndex        =   1
      Top             =   3840
      Width           =   13335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   $"form2.frx":27EC
      Height          =   975
      Left            =   1560
      TabIndex        =   0
      Top             =   2880
      Width           =   13335
   End
End
Attribute VB_Name = "metro_information"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label11_Click()
A_Front_page.Visible = True
metro_information.Visible = False
End Sub
