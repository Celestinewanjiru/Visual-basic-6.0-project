VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ANE 
   BackColor       =   &H80000016&
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Swis721 BlkCn BT"
      Size            =   11.25
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "ANE.frx":0000
      Height          =   1455
      Left            =   6720
      TabIndex        =   19
      Top             =   1920
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2566
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Swis721 BlkCn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Swis721 BlkCn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "Employee code"
         Caption         =   "Employee code"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Employee Name"
         Caption         =   "Employee Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Sex"
         Caption         =   "Sex"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Designation"
         Caption         =   "Designation"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Date of Joining"
         Caption         =   "Date of Joining"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "City"
         Caption         =   "City"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Country"
         Caption         =   "Country"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Phone No"
         Caption         =   "Phone No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1094.74
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   975
      Left            =   1200
      Top             =   8280
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1720
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
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton addnew 
      Caption         =   "Add New"
      Height          =   375
      Left            =   840
      TabIndex        =   17
      Top             =   7680
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      DataField       =   "Phone No"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Swis721 Cn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2640
      TabIndex        =   16
      Top             =   6840
      Width           =   3135
   End
   Begin VB.TextBox Text7 
      DataField       =   "Country"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Swis721 Cn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2640
      TabIndex        =   15
      Top             =   6120
      Width           =   3135
   End
   Begin VB.TextBox Text6 
      DataField       =   "City"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Swis721 Cn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2640
      TabIndex        =   14
      Top             =   5400
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      DataField       =   "Date of Joining"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Swis721 Cn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2640
      TabIndex        =   13
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      DataField       =   "Designation"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Swis721 Cn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2640
      TabIndex        =   12
      Top             =   3960
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      DataField       =   "Sex"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Swis721 Cn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2640
      TabIndex        =   11
      Top             =   3240
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      DataField       =   "Employee Name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Swis721 Cn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2640
      TabIndex        =   10
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      DataField       =   "Employee code"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Swis721 Cn BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2640
      TabIndex        =   9
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Phone No."
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Country"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Date of Joining"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Employee Name"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Employee code"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "ADD NEW EMPLOYEE"
      BeginProperty Font 
         Name            =   "Swis721 Blk BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   3615
   End
End
Attribute VB_Name = "ANE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addnew_Click()
With Adodc1.Recordset
.Fields(0) = Text1.Text
.Fields(1) = Text2.Text
.Fields(2) = Text3.Text
.Fields(3) = Text4.Text
.Fields(4) = Text5.Text
.Fields(5) = Text6.Text
.Fields(6) = Text7.Text
.Fields(7) = Text8.Text
.addnew
MsgBox "New EMPLOYEE Successfully added!!", vbInformation + vbOKOnly
End With
End Sub

Private Sub cmdcancel_Click()
Unload Me

End Sub

Private Sub Command1_Click()

End Sub


Private Sub Form_Load()
With Adodc1
.CommandType = adCmdTable
.RecordSource = "tbtEmployeeDetails"
.Refresh
.Recordset.addnew
End With

End Sub
