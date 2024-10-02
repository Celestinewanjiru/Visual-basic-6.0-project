VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form inputsalary 
   Caption         =   "1."
   ClientHeight    =   6735
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
   ScaleHeight     =   6735
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   615
      Left            =   1080
      TabIndex        =   19
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdcalc 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   4440
      TabIndex        =   18
      Top             =   5880
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "showsalary.frx":0000
      Height          =   1695
      Left            =   6120
      TabIndex        =   17
      Top             =   1920
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   2990
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Employee Code"
         Caption         =   "Employee Code"
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
         DataField       =   "Salary"
         Caption         =   "Salary"
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
         DataField       =   "Rent Allowances"
         Caption         =   "Rent Allowances"
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
         DataField       =   "Vehicle Allowances"
         Caption         =   "Vehicle Allowances"
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
         DataField       =   "Hospital Allowances"
         Caption         =   "Hospital Allowances"
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
            ColumnWidth     =   2085.166
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
      EndProperty
   End
   Begin VB.TextBox Text6 
      DataField       =   "Hospital Allowances"
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
      Height          =   495
      Left            =   3120
      TabIndex        =   16
      Top             =   4920
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      DataField       =   "Rent Allowances"
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
      Height          =   495
      Left            =   3120
      TabIndex        =   15
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      DataField       =   "Vehicle Allowances"
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
      Height          =   495
      Left            =   3120
      TabIndex        =   14
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton cmdprev 
      Caption         =   "&Prev"
      Height          =   615
      Left            =   2880
      TabIndex        =   10
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdnxt 
      Caption         =   "N&ext"
      Height          =   615
      Left            =   1320
      TabIndex        =   9
      Top             =   5760
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   600
      Top             =   8280
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PRO\Desktop\Employees\Employee.mdb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PRO\Desktop\Employees\Employee.mdb.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tbtsalarydetails"
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000009&
      Caption         =   "C&ancel"
      Height          =   495
      Left            =   4680
      MaskColor       =   &H00FFFF00&
      TabIndex        =   8
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H80000009&
      Caption         =   "O&K"
      Height          =   495
      Left            =   2760
      MaskColor       =   &H00FFFF00&
      TabIndex        =   7
      Top             =   6840
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      DataField       =   "Salary"
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
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   3000
      Width           =   2175
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
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      DataField       =   "Employee Code"
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
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000D&
      Caption         =   "Hospital Allowance"
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000D&
      Caption         =   "Rent Allowance"
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Caption         =   "Vehicle Allowance"
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Caption         =   "Basic Salary"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "Employee Name"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Employee code"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Salary"
      BeginProperty Font 
         Name            =   "Swis721 BlkCn BT"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   2640
   End
End
Attribute VB_Name = "inputsalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcalc_Click()

    Dim basicSalary As Double
    basicSalary = CDbl(Text3.Text)
    
    Dim vehicleAllowance As Double
    vehicleAllowance = CDbl(Text4.Text)
    
    Dim rentAllowance As Double
    rentAllowance = CDbl(Text5.Text)
    
    Dim hospitalAllowance As Double
    hospitalAllowance = CDbl(Text6.Text)
    
    Dim totalSalary As Double
    totalSalary = basicSalary + vehicleAllowance + rentAllowance + hospitalAllowance
    
    MsgBox "Total Salary: " & totalSalary
End Sub


Private Sub cmdnxt_Click()
Adodc1.Recordset.MoveNext

End Sub

Private Sub cmdok_Click()
With Adodc1.Recordset
.Fields(0) = Text1.Text
.Fields(1) = Text2.Text
.Fields(2) = Text3.Text
.addnew
MsgBox "Salary successfully recorded!!", vbInformation + vbOKOnly
End With
End Sub

Private Sub cmdprev_Click()
Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command1_Click()
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

Private Sub Command2_Click()
Unload Me
main.show
End Sub

Private Sub Form_Load()
With Adodc1
.CommandType = adCmdTable
.RecordSource = "tbtsalarydetails"
.Refresh
.Recordset.addnew
End With
End Sub

Private Sub Timer1_Timer()
lbldate = Date
lbltime = Time
End Sub

Private Sub lbltime_Click(Index As Integer)

End Sub
