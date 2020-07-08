VERSION 5.00
Begin VB.Form frmChangePassword 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   6045
   ClientLeft      =   4170
   ClientTop       =   2970
   ClientWidth     =   7800
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   7800
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   6495
      Begin VB.CommandButton Cmdaddnew 
         Caption         =   "Clear"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox txtconpass 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   10
         Top             =   3360
         Width           =   2655
      End
      Begin VB.TextBox txtnpass 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   8
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox txtcpass 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   7
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
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
         Left            =   4200
         TabIndex        =   6
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change"
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
         Left            =   2160
         TabIndex        =   5
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox txtusername 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   2
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Confirm Password"
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
         Left            =   480
         TabIndex        =   9
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Current Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdaddnew_Click()
MsgBox ("Do you want to Clear")
txtconpass = " "
txtcpass = " "
txtnpass = " "
txtusername = " "


End Sub

Private Sub cmdCancel_Click()
MsgBox ("Do you want to Exit")
Unload Me
End Sub

Private Sub cmdChange_Click()
'On Error GoTo er
If (txtnpass.Text <> txtconpass.Text) Then
    MsgBox "Password mismatching"
Else
    Set rs = con.Execute("select * from login where username='" + txtusername.Text + "' and password='" + txtcpass.Text + "'")
    If (Not rs.EOF) Then
      '  con.Execute ("UPDATE login set password='" & txtnpass.Text & "' where username='" & txtusername.Text & "'")
        
        rs!Password = txtnpass.Text
        rs.Update
        MsgBox "Updated"
        'MsgBox rs.Fields(0)
    Else
       MsgBox "User not exist"
    End If
End If
'Exit Sub
'er:
'con.Execute ("UPDATE login set password='" + txtnpass.Text + "' where username='" + txtusername.Text + "'")
End Sub

Private Sub Form_Load()
connectdb
End Sub
