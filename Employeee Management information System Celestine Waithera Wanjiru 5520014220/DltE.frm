VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form DltE 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Swis721 BlkCn BT"
      Size            =   9.75
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
      Bindings        =   "DltE.frx":0000
      Height          =   2535
      Left            =   6480
      TabIndex        =   21
      Top             =   2280
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   4471
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Swis721 BlkCn BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Swis721 BlkCn BT"
         Size            =   9.75
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
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "DELETE EMPLOYEE"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   360
         Top             =   7560
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox Text5 
         DataField       =   "Date of Joining"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Swis721 Cn BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2610
         TabIndex        =   20
         Top             =   3136
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         DataField       =   "Employee Name"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Swis721 Cn BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2610
         TabIndex        =   19
         Top             =   1234
         Width           =   2295
      End
      Begin VB.CommandButton Ccmdnxt 
         Caption         =   "Next"
         Height          =   375
         Left            =   2760
         TabIndex        =   18
         Top             =   5880
         Width           =   1935
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "Prev"
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Top             =   5880
         Width           =   1935
      End
      Begin VB.CommandButton Ccmdcancel 
         Caption         =   "Cancel"
         Height          =   615
         Left            =   3000
         TabIndex        =   16
         Top             =   6720
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         DataField       =   "Phone No"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Swis721 Cn BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2610
         TabIndex        =   15
         Top             =   5040
         Width           =   2295
      End
      Begin VB.TextBox Text7 
         DataField       =   "Country"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Swis721 Cn BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2610
         TabIndex        =   14
         Top             =   4404
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         DataField       =   "City"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Swis721 Cn BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2610
         TabIndex        =   13
         Top             =   3770
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         DataField       =   "Designation"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Swis721 Cn BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2610
         TabIndex        =   12
         Top             =   2502
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         DataField       =   "Sex"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Swis721 Cn BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2610
         TabIndex        =   11
         Top             =   1868
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         DataField       =   "Employee code"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Swis721 Cn BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2610
         TabIndex        =   10
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton cmddlt 
         Caption         =   "Delete Employee"
         Height          =   615
         Left            =   480
         TabIndex        =   9
         Top             =   6720
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Phone number"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Country"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   4404
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "City"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   3770
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Date  of Joining"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   3136
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Designation"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   2502
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Sex"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1868
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Employee Name"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1234
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Employee Code"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
   End
End
Attribute VB_Name = "DltE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Ccmdcancel_Click()
Unload DltE
main.show
End Sub

Private Sub Ccmdnxt_Click()
Adodc1.Recordset.MoveNext

End Sub

Private Sub cmddlt_Click()
    On Error Resume Next
    
       If Not Adodc1.Recordset.EOF Then
                
                If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion, "Confirm Deletion") = vbYes Then
        
            Adodc1.Recordset.Delete
            MsgBox "Record deleted successfully.", vbInformation, "Success"
        End If
    Else
        MsgBox "No record to delete.", vbExclamation, "Error"
    End If
End Sub

Private Sub cmdprev_Click()
Adodc1.Recordset.MovePrevious

End Sub


