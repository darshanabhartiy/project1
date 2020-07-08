VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADMIN"
   ClientHeight    =   5130
   ClientLeft      =   3240
   ClientTop       =   2610
   ClientWidth     =   8205
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8205
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   3015
      Left            =   1680
      TabIndex        =   0
      Top             =   1560
      Width           =   5775
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   2400
         Width           =   1455
      End
      Begin VB.ComboBox cmbUsername 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2640
         TabIndex        =   5
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Password"
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
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Username"
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
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   2175
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "M.S.R.T.C. ADMIN LOGIN"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   5055
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
MDIForm1.Show
Unload Me
End Sub

Private Sub cmdLogin_Click()
Set rs = con.Execute("select * from login where username='" + cmbUsername.Text + "' and password='" + txtPassword.Text + "'")
If (Not rs.EOF) Then
    MsgBox "Login Success"
    MDIForm1.route.Enabled = True
    MDIForm1.detail.Enabled = True
    MDIForm1.stops.Enabled = True
    MDIForm1.ticket.Enabled = True
    MDIForm1.trip.Enabled = True
    MDIForm1.logout.Enabled = True
    MDIForm1.newuser.Enabled = True
    MDIForm1.cpassword.Enabled = True
    MDIForm1.login.Enabled = False
    Unload Me
Else
    MsgBox "Login Failure! Try Again"
    cmbUsername.ListIndex = 0
    txtPassword.Text = ""
End If
End Sub

Private Sub Form_Load()
connectdb
Set rs = con.Execute("select * from login")
While (Not rs.EOF)
    cmbUsername.AddItem rs(0)
    rs.MoveNext
Wend
rs.Close

End Sub
