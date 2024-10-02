VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form EDIT 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Swis721 BlkCn BT"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1440
      Top             =   9000
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PRO\Desktop\Employees\Employee.mdb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PRO\Desktop\Employees\Employee.mdb.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tbtEmployeeDetails"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Swis721 BlkCn BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton CMDEDITOK 
      Caption         =   "O&K"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   19
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdprev 
      Caption         =   "Pr&evious"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Nex&t"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   17
      Top             =   8040
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      DataField       =   "Phone No"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Swis721 Cn BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   16
      Top             =   7200
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      DataField       =   "Country"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Swis721 Cn BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   14
      Top             =   6372
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      DataField       =   "City"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Swis721 Cn BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   13
      Top             =   5550
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      DataField       =   "Date of Joining"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Swis721 Cn BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   12
      Top             =   4728
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      DataField       =   "Designation"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Swis721 Cn BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   11
      Top             =   3906
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      DataField       =   "Sex"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Swis721 Cn BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   10
      Top             =   3084
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      DataField       =   "Employee Name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Swis721 Cn BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   9
      Top             =   2262
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      DataField       =   "Employee code"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Swis721 Cn BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   8
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000D&
      Caption         =   "Phone No."
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   600
      TabIndex        =   15
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      Caption         =   "Country"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   6372
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000D&
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   5550
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000D&
      Caption         =   "Date of joining"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   4728
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   3906
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   3084
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "Employee Name"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2262
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Employee Code"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Edit Employee Records"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   630
      Left            =   1785
      TabIndex        =   0
      Top             =   360
      Width           =   4965
   End
End
Attribute VB_Name = "EDIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDEDITOK_Click()
Unload Me
main.show

End Sub

Private Sub cmdnext_Click()
Adodc1.Recordset.MoveNext

End Sub

Private Sub cmdprev_Click()
Adodc1.Recordset.MovePrevious

End Sub

Private Sub Command1_Click()

End Sub
