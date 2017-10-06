VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Fare_Enquiry 
   BackColor       =   &H00FFFF80&
   Caption         =   "Fare Enquiry"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1560
      Top             =   4560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      RecordSource    =   "ROUTES"
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
   Begin VB.CommandButton Command1 
      Caption         =   "fare enquiry"
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
      Left            =   9240
      TabIndex        =   8
      Top             =   9000
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   9000
      Width           =   1935
   End
   Begin VB.ComboBox SOURCE1 
      DataField       =   "SOURCE"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "Fare_Enquiry.frx":0000
      Left            =   5520
      List            =   "Fare_Enquiry.frx":002B
      TabIndex        =   1
      Top             =   8280
      Width           =   2055
   End
   Begin VB.ComboBox DESTINATION1 
      DataField       =   "DESTINATION"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "Fare_Enquiry.frx":00AE
      Left            =   10800
      List            =   "Fare_Enquiry.frx":00D9
      TabIndex        =   0
      Top             =   8280
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C000&
      Caption         =   "Click here to go back on front page:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9120
      TabIndex        =   6
      Top             =   10080
      Width           =   2775
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
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Fare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   9000
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      Top             =   8280
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Source"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   9945
      Left            =   -480
      Picture         =   "Fare_Enquiry.frx":015C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18000
   End
End
Attribute VB_Name = "Fare_Enquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New adodb.Recordset
Dim cs As New adodb.Connection
Dim SOURCE As String

Private Sub Command1_Click()
SOURCE1.Text = ""
DESTINATION1.Text = ""
rs.Open "select * from routes where source='" & Me.SOURCE1.Text & "' and destination='" & Me.DESTINATION1.Text & "'", cs

Text1.Text = rs.Fields(3)
rs.Close
End Sub

Private Sub Form_Load()
cs.Open "Provider=MSDAORA.1;Password=vijay;User ID=vijay;Persist Security Info=True"
End Sub

Private Sub Label5_Click()
A_Front_page.Show
Unload Me
End Sub
