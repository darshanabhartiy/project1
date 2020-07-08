VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H80000008&
   Caption         =   "MDIForm1"
   ClientHeight    =   8340
   ClientLeft      =   795
   ClientTop       =   1890
   ClientWidth     =   13800
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.Menu admin 
      Caption         =   "Admin"
      Begin VB.Menu login 
         Caption         =   "Login"
      End
      Begin VB.Menu logout 
         Caption         =   "Logout"
      End
      Begin VB.Menu d 
         Caption         =   "-"
      End
      Begin VB.Menu newuser 
         Caption         =   "Add New User"
      End
      Begin VB.Menu cpassword 
         Caption         =   "Change Password"
      End
   End
   Begin VB.Menu detail 
      Caption         =   "Bus Details"
   End
   Begin VB.Menu stops 
      Caption         =   "Bus Stops"
   End
   Begin VB.Menu ticket 
      Caption         =   "Trip Information"
   End
   Begin VB.Menu route 
      Caption         =   "Management of Route"
   End
   Begin VB.Menu trip 
      Caption         =   "Ticketing"
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cpassword_Click()
Load frmChangePassword
frmChangePassword.Show
End Sub

Private Sub detail_Click()
 Form2.Show

End Sub

Private Sub exit_Click()
Unload Me

End Sub

Private Sub login_Click()
Load Form6
Form6.Show
End Sub

Private Sub logout_Click()
    MDIForm1.route.Enabled = False
    MDIForm1.detail.Enabled = False
    MDIForm1.stops.Enabled = False
    MDIForm1.ticket.Enabled = False
    MDIForm1.trip.Enabled = False
    MDIForm1.logout.Enabled = False
    MDIForm1.newuser.Enabled = False
    MDIForm1.cpassword.Enabled = False
    MDIForm1.login.Enabled = True
End Sub

Private Sub MDIForm_Load()
route.Enabled = False
detail.Enabled = False
stops.Enabled = False
ticket.Enabled = False
trip.Enabled = False
logout.Enabled = False
newuser.Enabled = False
cpassword.Enabled = False
Load Form6
Form6.Show
End Sub

Private Sub route_Click()
Form1.Show
End Sub



Private Sub stops_Click()
Form3.Show
End Sub

Private Sub ticket_Click()
Form4.Show

End Sub

Private Sub trip_Click()
Form5.Show
End Sub
