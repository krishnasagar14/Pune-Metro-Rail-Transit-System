VERSION 5.00
Begin VB.Form Sign_up 
   BackColor       =   &H00FFFF80&
   Caption         =   "Sign up"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   FillColor       =   &H008080FF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox ADDRESS1 
      Height          =   1215
      Left            =   11880
      TabIndex        =   6
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Frame SIGN_UP 
      BackColor       =   &H00C0E0FF&
      Caption         =   "SIGN_UP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   9615
      Left            =   8760
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      Begin VB.TextBox gender1 
         Height          =   375
         Left            =   3120
         TabIndex        =   17
         Top             =   7560
         Width           =   4335
      End
      Begin VB.TextBox password 
         Height          =   495
         Left            =   3120
         TabIndex        =   16
         Top             =   6240
         Width           =   4455
      End
      Begin VB.CommandButton CANCEL 
         Caption         =   "CANCEL"
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
         Left            =   6120
         TabIndex        =   14
         Top             =   8640
         Width           =   2175
      End
      Begin VB.CommandButton NEXT 
         Caption         =   "NEXT"
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
         Left            =   3480
         TabIndex        =   13
         Top             =   8640
         Width           =   2175
      End
      Begin VB.TextBox PHNO 
         Height          =   495
         Left            =   3120
         TabIndex        =   10
         Top             =   5040
         Width           =   4455
      End
      Begin VB.TextBox DOB 
         Height          =   495
         Left            =   3120
         TabIndex        =   8
         Top             =   3840
         Width           =   4455
      End
      Begin VB.TextBox NAME1 
         Height          =   495
         Left            =   3120
         TabIndex        =   3
         Top             =   1080
         Width           =   4455
      End
      Begin VB.Label PASSWORD1 
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD*"
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
         Left            =   720
         TabIndex        =   15
         Top             =   6240
         Width           =   1935
      End
      Begin VB.Label GENDER 
         BackStyle       =   0  'Transparent
         Caption         =   "GENDER*"
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
         Left            =   720
         TabIndex        =   11
         Top             =   7680
         Width           =   1935
      End
      Begin VB.Label CONTACTNO 
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT NO*"
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
         Left            =   720
         TabIndex        =   9
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label DATEOFBIRTH 
         BackStyle       =   0  'Transparent
         Caption         =   "DATEOFBIRTH*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   7
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label ADDRESS 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS*"
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
         Left            =   720
         TabIndex        =   4
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label NAME 
         BackStyle       =   0  'Transparent
         Caption         =   "NAME*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   6975
      Left            =   1080
      Picture         =   "Sign_up.frx":0000
      ScaleHeight     =   6915
      ScaleWidth      =   12075
      TabIndex        =   5
      Top             =   120
      Width           =   12135
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   495
      Left            =   9600
      TabIndex        =   12
      Top             =   5280
      Width           =   1215
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
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Sign_up"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n, addr, dt, no As String
Dim cs As New adodb.Connection
Dim rs As New adodb.Recordset


Private Sub address1_KeyPress(KeyAscii As Integer)
If ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8 Or KeyAscii = 32 Or (KeyAscii >= 61 And KeyAscii <= 122) Or KeyAscii = 92 Or (KeyAscii >= 44 And KeyAscii <= 46) Or (KeyAscii >= 48 And KeyAscii <= 57)) Then
addr = ADDRESS1.Text
ADDRESS1.Tag = 1
Else
KeyAscii = 0
MsgBox "Enter address1 in only a-z,A-Z,0-9,',','.'.", vbOKOnly, "invalid address1"
End If
End Sub

Private Sub CANCEL_Click()
A_Front_page.Show
Unload Me
End Sub

Private Sub DOB_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 47 Or KeyAscii = 8 Or KeyAscii = 32 Then
dt = DOB.Text
DOB.Tag = 1
Else
KeyAscii = 0
MsgBox "Enter date of birth in only form(dd/mm/yyyy)or(dd-mm-yyyy).", vbOKOnly, "invalid date of birth"
End If
End Sub


Private Sub Form_Load()
cs.Open "Provider=MSDAORA.1;Password=vijay;User ID=vijay;Persist Security Info=True"
End Sub


Private Sub NAME1_KeyPress(KeyAscii As Integer)
If ((KeyAscii >= 65 And KeyAscii <= 122) Or KeyAscii = 8 Or KeyAscii = 32) Then
n = NAME1.Text
NAME1.Tag = 1
Else
KeyAscii = 0
MsgBox "Enter name in Uppercase only.", vbOKOnly, "invalid name"
End If
End Sub

Private Sub NEXT_Click()
rs.Open "Insert into signup values('" & NAME1.Text & "','" & ADDRESS1.Text & "','" & DOB.Text & "','" & PHNO.Text & "','" & password.Text & "','" & gender1.Text & "')", cs, adOpenDynamic, adLockOptimistic
sign_up1.Show
Unload Me
End Sub

Private Sub PHNO_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 32 And Len(PHNO.Text) <= 10 Then
no = PHNO.Text
PHNO.Tag = 1
Else
MsgBox "Enter 10 digit phone no .", vbOKOnly, "invalid phone no"
End If
End Sub
