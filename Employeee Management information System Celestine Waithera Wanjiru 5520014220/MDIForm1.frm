VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu empdt 
      Caption         =   "&Employee Details"
      Begin VB.Menu addemp 
         Caption         =   "Add Employee"
      End
      Begin VB.Menu dltemp 
         Caption         =   "Delete Employee"
      End
      Begin VB.Menu editemp 
         Caption         =   "Edit Employee"
      End
      Begin VB.Menu emplview 
         Caption         =   "View Employee"
      End
      Begin VB.Menu datarpt 
         Caption         =   "Data Report"
      End
   End
   Begin VB.Menu salary 
      Caption         =   "&Salary"
      Begin VB.Menu show 
         Caption         =   "Show Salary Information"
      End
      Begin VB.Menu inpsalary 
         Caption         =   "Input Salary"
      End
      Begin VB.Menu dlt 
         Caption         =   "Delete Records"
      End
   End
   Begin VB.Menu abt 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub abt_Click()

End Sub
