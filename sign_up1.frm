VERSION 5.00
Begin VB.Form sign_up1 
   BackColor       =   &H00FFFF80&
   Caption         =   "sign_up(continued)"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      Caption         =   "ACCOUNTS_DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6975
      Left            =   8880
      TabIndex        =   1
      Top             =   3240
      Width           =   11775
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   3960
         Width           =   3975
      End
      Begin VB.CommandButton BACK 
         Caption         =   "BACK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   10
         Top             =   4800
         Width           =   1815
      End
      Begin VB.CommandButton CREATEACC 
         Caption         =   "CREATE A ACCOUNT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   4800
         Width           =   3855
      End
      Begin VB.TextBox bankname 
         Height          =   495
         Left            =   2640
         TabIndex        =   7
         Top             =   2760
         Width           =   3975
      End
      Begin VB.TextBox accno 
         Height          =   495
         Left            =   2640
         TabIndex        =   5
         Top             =   1680
         Width           =   3975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Card TYPE*     Eg Atm,Maestro, Credit card"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   840
         TabIndex        =   8
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label nameofbank 
         BackStyle       =   0  'Transparent
         Caption         =   "NAME OF BANK*"
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
         Left            =   840
         TabIndex        =   6
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label acctno 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNT NUMBER*"
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
         Left            =   840
         TabIndex        =   4
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label NAMEPLATE 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   3
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label NAME 
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Image Image2 
      Height          =   4155
      Left            =   1080
      Picture         =   "sign_up1.frx":0000
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   7770
   End
   Begin VB.Image Image1 
      Height          =   5475
      Left            =   1080
      Picture         =   "sign_up1.frx":583C
      Top             =   120
      Width           =   9540
   End
   Begin VB.Label Label2 
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
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "sign_up1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cs As New adodb.Connection
Dim rs As New adodb.Recordset
Dim v As SIGN_UP

Private Sub BACK_Click()
A_Front_page.Show
Unload Me
End Sub

Private Sub CREATEACC_Click()
rs.Open "Insert into account values('" & accno.Text & "',' " & bankname.Text & "' ,' " & Text1.Text & " ')", cs, adOpenDynamic, adLockOptimistic
MsgBox "Your accoutn is created succesfully"
accno.Text = ""
bankname.Text = ""
Text1.Text = ""

End Sub

Private Sub Form_Load()
cs.Open "Provider=MSDAORA.1;Password=vijay;User ID=vijay;Persist Security Info=True"
NAMEPLATE.Caption = SIGN_UP.NAME1.Text
End Sub

Private Sub NAMEPLATE_Click()
NAMEPLATE.Caption = v.NAME1.Text
End Sub
