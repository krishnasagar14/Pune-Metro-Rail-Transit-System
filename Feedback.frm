VERSION 5.00
Begin VB.Form Feedback 
   BackColor       =   &H00FFFF80&
   Caption         =   "Feedback"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
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
      Left            =   5760
      TabIndex        =   7
      Top             =   10080
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   1335
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   8040
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   5640
      TabIndex        =   4
      Top             =   7440
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   6720
      Width           =   3615
   End
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
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Feed back:-"
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
      Left            =   2880
      TabIndex        =   5
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact no:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   7440
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME :-"
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
      Left            =   2880
      TabIndex        =   1
      Top             =   6720
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FEEDBACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   8040
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin VB.Image Image2 
      Height          =   4275
      Left            =   9000
      Picture         =   "Feedback.frx":0000
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   6570
   End
   Begin VB.Image Image1 
      Height          =   4290
      Left            =   1320
      Picture         =   "Feedback.frx":583C
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   7770
   End
End
Attribute VB_Name = "Feedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New adodb.Recordset
Dim cs As New adodb.Connection

Private Sub Command1_Click()
rs.Open "Insert into feedback values('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "')", cs, adOpenDynamic, adLockOptimistic
If Text1.Text = "" Then
MsgBox "Incomplete information.Please fill the form completely", vbOKOnly, "invalid form"
Text1.SetFocus
ElseIf Text2.Text = "" Then
MsgBox "Incomplete information.Please fill the form completely", vbOKOnly, "invalid form"
Text2.SetFocus
ElseIf Text3.Text = "" Then
MsgBox "Incomplete information.Please fill the form completely", vbOKOnly, "invalid form"
Text3.SetFocus
Else
MsgBox "Form submitted succesfully.", vbYesNoCancel, "valid form"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End If
A_Front_page.Show
Unload Me
End Sub

Private Sub Form_Load()
cs.Open "Provider=MSDAORA.1;Password=vijay;User ID=vijay;Persist Security Info=True"
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
If ((KeyAscii > 33 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 123) Or KeyAscii = 8 Or KeyAscii = 32) Then
End If
End Sub
