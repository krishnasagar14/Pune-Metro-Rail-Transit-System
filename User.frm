VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form User 
   BackColor       =   &H00FFFF00&
   Caption         =   "User name and password"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "User.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   12240
      TabIndex        =   10
      Top             =   6600
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5280
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=OraOLEDB.Oracle.1;Password=vijay;Persist Security Info=True;User ID=vijay"
      OLEDBString     =   "Provider=OraOLEDB.Oracle.1;Password=vijay;Persist Security Info=True;User ID=vijay"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   8
      Top             =   8400
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go back to prevoius page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   7
      Top             =   8400
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   6
      Top             =   8400
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   7080
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      CausesValidation=   0   'False
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   6360
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   5640
      Width           =   2655
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
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Re Enter Password:-"
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
      Left            =   2520
      TabIndex        =   4
      Top             =   7080
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:-"
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
      Left            =   2520
      TabIndex        =   2
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :-"
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
      Left            =   2520
      TabIndex        =   0
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Image Image3 
      Height          =   3930
      Left            =   14640
      Picture         =   "User.frx":0342
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   5610
   End
   Begin VB.Image Image2 
      Height          =   3915
      Left            =   7560
      Picture         =   "User.frx":6E2B
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   7170
   End
   Begin VB.Image Image1 
      Height          =   3930
      Left            =   0
      Picture         =   "User.frx":C667
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   7650
   End
End
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cs As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
rs.Open "Insert into user_tickets values('" & Text1.Text & "','" & Text2.Text & "')", cs, adOpenDynamic, adLockOptimistic
On Error GoTo ER

If Text2.Text = Text3.Text Then
MsgBox "collect your ticket happy journey"
Else
User.Show
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End If
A_Front_page.Show
Unload Me
Exit Sub
ER:
MsgBox "Enter the correct password"

End Sub

Private Sub Command2_Click()
Tickets.Show
Unload Me
End Sub

Private Sub Command3_Click()
A_Front_page.Show
Unload Me
End Sub

Private Sub Command4_Click()
DataReport1.Show
End Sub

Private Sub Form_Load()
cs.Open "Provider=MSDAORA.1;Password=vijay;User ID=vijay;Persist Security Info=True"
End Sub

Private Sub Form_LostFocus()
rs.Close
cs.Close
End Sub
