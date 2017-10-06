VERSION 5.00
Begin VB.Form SMS 
   BackColor       =   &H00FFFF00&
   Caption         =   "SMS Services"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
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
      Left            =   10080
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   480
      Width           =   2775
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
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Back on to front Page click here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12360
      TabIndex        =   7
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Email_id:- *"
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
      Left            =   1680
      TabIndex        =   5
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Contact No:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Name:- *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   7545
      Left            =   1080
      Picture         =   "SMS.frx":0000
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   13620
   End
End
Attribute VB_Name = "SMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cs As New adodb.Connection
Dim rs As New adodb.Recordset
Private Sub Command1_Click()
rs.Open "Insert into sms values('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "')", cs, adOpenDynamic, adLockOptimistic
If Text1.Text = "" Then
MsgBox "Incomplete information.Please fill the form completely", vbOKOnly, "invalid form"
Text1.SetFocus
ElseIf Text3.Text = "" Then
MsgBox "Incomplete information.Please fill the form completely", vbOKOnly, "invalid form"
Text3.SetFocus
Else
MsgBox "Your SMS Service Activated Shortly Thank you for registration", vbYesNoCancel, "valid form"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End If
End Sub

Private Sub Command2_Click()
DataReport1.Show
End Sub

Private Sub Form_Load()
cs.Open "Provider=MSDAORA.1;Password=vijay;User ID=vijay;Persist Security Info=True"
End Sub

Private Sub Label4_Click()
A_Front_page.Visible = True
SMS.Visible = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If ((KeyAscii > 65 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 123) Or KeyAscii = 8 Or KeyAscii = 32) Then
Else
KeyAscii = 0
MsgBox "Enter Proper Name No numeric value allow.", vbOKOnly, "invalid name"
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If (Len(Text2.Text) = 10) Then
MsgBox "Enter proper contact number.", vbOKOnly, "Valid"
Text2.Text = ""
End If
If KeyAscii = 32 Then
MsgBox "Enter the 10 digit mobile number eg 99999-00000"
End If
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 64 And KeyAscii <= 90) Or (KeyAscii > 96 And KeyAscii <= 122) Then
Else
If KeyAscii = 46 Then
Else
If KeyAscii >= 48 And KeyAscii <= 57 Then
Else
If keysacii = 32 Then
Else
KeyAscii = 0
MsgBox "Enter email Id for eg vijayk512@gmail.com", vbOKOnly, "Invalid id"
End If
End If
End If
End If
End Sub
