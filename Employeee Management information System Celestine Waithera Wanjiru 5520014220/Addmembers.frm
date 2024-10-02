VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Swis721 BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   16
      Text            =   "Text8"
      Top             =   8005
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Swis721 BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   15
      Text            =   "Text7"
      Top             =   6960
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Swis721 BT"
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
      Text            =   "Text6"
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Swis721 BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Swis721 BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Swis721 BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   11
      Text            =   "Text3"
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Swis721 BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Swis721 BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Club Joined"
      BeginProperty Font 
         Name            =   "Swis721 BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Phone Number"
      BeginProperty Font 
         Name            =   "Swis721 BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   6990
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Swis721 BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Date Joined"
      BeginProperty Font 
         Name            =   "Swis721 BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Member Sex"
      BeginProperty Font 
         Name            =   "Swis721 BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Member DOB"
      BeginProperty Font 
         Name            =   "Swis721 BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Member Name"
      BeginProperty Font 
         Name            =   "Swis721 BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Members ID"
      BeginProperty Font 
         Name            =   "Swis721 BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Add members"
      BeginProperty Font 
         Name            =   "Wicked Mouse"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
