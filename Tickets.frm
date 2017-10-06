VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Tickets 
   BackColor       =   &H00C0C000&
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Check For fare click here"
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13680
      TabIndex        =   10
      Top             =   9360
      Width           =   975
   End
   Begin VB.CommandButton REPLAN 
      Caption         =   "REPLAN"
      Height          =   375
      Left            =   8520
      TabIndex        =   9
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton CONFIRMTICKET 
      Caption         =   "GET THE TICKET"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   3120
      Width           =   5055
   End
   Begin VB.ComboBox DESTINATION1 
      Height          =   315
      ItemData        =   "Tickets.frx":0000
      Left            =   7920
      List            =   "Tickets.frx":002B
      TabIndex        =   4
      Top             =   840
      Width           =   2295
   End
   Begin VB.ComboBox SOURCE1 
      Height          =   315
      ItemData        =   "Tickets.frx":00AE
      Left            =   3600
      List            =   "Tickets.frx":00D9
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   13680
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=MSDAORA.1;Password=vijay;User ID=vijay;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=vijay;User ID=vijay;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "vijay"
      Password        =   "vijay"
      RecordSource    =   "FARE_ENQUIRY"
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
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.Label FARE 
      BackStyle       =   0  'Transparent
      Caption         =   "FARE"
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
      TabIndex        =   6
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label JOURNEYDISTANCE_KM 
      BackStyle       =   0  'Transparent
      Caption         =   "JOURNEY DISTANCE(KM)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label DESTINATION 
      BackStyle       =   0  'Transparent
      Caption         =   "DESTINATION"
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
      Left            =   6120
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label SOURCE 
      BackStyle       =   0  'Transparent
      Caption         =   "SOURCE"
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
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label PLANMYTRAVEL 
      BackStyle       =   0  'Transparent
      Caption         =   "PLAN MY TRAVEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   12375
   End
   Begin VB.Image Image1 
      Height          =   6135
      Left            =   1440
      Picture         =   "Tickets.frx":015C
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   12045
   End
End
Attribute VB_Name = "Tickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New adodb.Recordset
Dim cs As New adodb.Connection
Dim c_source, c_destination As String
Private Sub Command1_Click()
A_Front_page.Visible = True
Tickets.Visible = False
End Sub

Private Sub Command2_Click()
'While rs.EOF
 '   If StrComp(rs.Fields(0), Text1.Text) = 0 Then
        
    




rs.Open "select * from routes where source='" & Me.SOURCE1.Text & "' and destination='" & Me.DESTINATION1.Text & "'", cs

Text1.Text = rs.Fields(2)
Text2.Text = rs.Fields(3)
rs.Close
End Sub

Private Sub Form_Load()
cs.Open "Provider=MSDAORA.1;Password=vijay;User ID=vijay;Persist Security Info=True"

End Sub
Private Sub CONFIRMTICKET_Click()
rs.Open "Insert into ticket values('" & SOURCE1.Text & "','" & DESTINATION1.Text & "','" & Text2.Text & "')", cs, adOpenDynamic, adLockOptimistic

On Error GoTo ER
If SOURCE1.Text = "" Then
MsgBox "Please select source and destination"
SOURCE1.SetFocus
ElseIf DESTINATION1.Text = "" Then
MsgBox "Please select source and destination"
DESTINATION1.SetFocus
Else
If SOURCE1.Text = "" Or DESTINATION1.Text = "" Then
User.Visible = False
Tickets.Show
Else
MsgBox "Click on OK and Provide your user name and password"
SOURCE1.Text = ""
DESTINATION1.Text = ""
User.Show
Unload Me
End If
End If
Exit Sub
ER:
MsgBox "please provide all details"

End Sub

