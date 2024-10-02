VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form login 
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Wicked Mouse"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   6615
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   1080
         Top             =   3720
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1085
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PRO\Desktop\Employees\Employee.mdb.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PRO\Desktop\Employees\Employee.mdb.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from tbtlogin"
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
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000D&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   6
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton cmdlogin 
         BackColor       =   &H8000000D&
         Caption         =   "&Login"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   5
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         DataField       =   "Password"
         BeginProperty Font 
            Name            =   "Swis721 Cn BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   3000
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         DataField       =   "Username"
         BeginProperty Font 
            Name            =   "Swis721 Cn BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Swis721 BlkCn BT"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Swis721 BlkCn BT"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Employee Management Information System"
      BeginProperty Font 
         Name            =   "Athena Rustic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1185
      TabIndex        =   7
      Top             =   720
      Width           =   5565
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdlogin_Click()
Adodc1.RecordSource = "SELECT * FROM tbtlogin WHERE Username='" & Text1.Text & "' AND Password='" & Text2.Text & "'"
Adodc1.Refresh

  
  If Adodc1.Recordset.EOF Then
  MsgBox "Login failed, try again!!", vbCritical, "Please enter correct Username and Password"
Else
MsgBox "Login Successful.", vbInformation, "Sucessful attempt"
main.show
End If
End Sub

Private Sub Command2_Click()
Unload Me

End Sub
