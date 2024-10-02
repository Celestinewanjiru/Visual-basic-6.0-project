VERSION 5.00
Begin VB.Form About 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   FillColor       =   &H00FFFFC0&
   BeginProperty Font 
      Name            =   "Swis721 Blk BT"
      Size            =   12
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
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   855
      Left            =   1560
      TabIndex        =   1
      Top             =   6360
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Employee Management Information System"
      ForeColor       =   &H8000000B&
      Height          =   5295
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   6375
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Swis721 Cn BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   6
         Text            =   "Programmer"
         Top             =   3960
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Swis721 Cn BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   5
         Text            =   "Celestine Wanjiru"
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "I HOPE THIS PROJECT WILLL FULFIL REQUIREMENTS OF THE COMPANY"
         BeginProperty Font 
            Name            =   "Swis721 Cn BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1680
         TabIndex        =   4
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "The project keeps records of employees working in the company it can add, edit, delete employee records"
         BeginProperty Font 
            Name            =   "Swis721 Cn BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   4815
      End
      Begin VB.Label Label1 
         Caption         =   "Employee  and  Payroll  System""  is  a  software  which  is Computerized  to  solve  the  problem related to employees ."
         BeginProperty Font 
            Name            =   "Swis721 Cn BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   5895
      End
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub File1_Click()

End Sub

Private Sub Command1_Click()
Unload About
main.show
End Sub
