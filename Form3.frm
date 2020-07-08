VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7020
   ClientLeft      =   3135
   ClientTop       =   2310
   ClientWidth     =   10485
   Icon            =   "Form3.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   10485
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   1200
      TabIndex        =   0
      Top             =   1200
      Width           =   7575
      Begin VB.CommandButton Command3 
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
         Height          =   615
         Left            =   5400
         TabIndex        =   12
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add New"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   11
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         TabIndex        =   10
         Top             =   3720
         Width           =   1455
      End
      Begin VB.OptionButton Optno 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   9
         Top             =   2760
         Width           =   1215
      End
      Begin VB.OptionButton Optyes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   8
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Txtstopname 
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
         Left            =   3120
         TabIndex        =   7
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox Txtstopno 
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
         Left            =   3120
         TabIndex        =   6
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox Txtrouteno 
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
         Left            =   3120
         TabIndex        =   5
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fare Stage"
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
         Left            =   600
         TabIndex        =   4
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Stop Name"
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
         Left            =   600
         TabIndex        =   3
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Stop No:"
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
         Left            =   600
         TabIndex        =   2
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Route No:"
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
         Left            =   600
         TabIndex        =   1
         Top             =   960
         Width           =   2055
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "M.S.R.T.C. BUS STOPS"
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
      Height          =   495
      Left            =   2760
      TabIndex        =   13
      Top             =   480
      Width           =   5175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click(Index As Integer)
MsgBox ("Do You want to Exit")
Me.Hide
End Sub

Private Sub Cmdnew_Click()
Txtrnumber = " "
Txtsnumber = " "
txtsname = " "
End Sub

Private Sub Command1_Click()
If (Optyes.Value = True) Then
    x = "Yes"
Else
    x = "No"
End If



If Txtrouteno.Text = "" Then
MsgBox "Please Enter the routenumber.", vbInformation
Txtrouteno.SetFocus
Exit Sub
End If

If Txtstopno.Text = "" Then
MsgBox "Please Enter the Stop .", vbInformation
Txtstopno.SetFocus
Exit Sub
End If

If Txtstopname.Text = "" Then
MsgBox "Please Enter Stop Name .", vbInformation
Txtstopname.SetFocus
Exit Sub
End If
con.Execute ("insert into busstop values(" + Txtrouteno.Text + "," + Txtstopno.Text + ",'" + Txtstopname.Text + "', '" + x + "'  )")
MsgBox ("successfully saved")
End Sub



Private Sub Command2_Click()
MsgBox ("Do you want to Clear")
Txtrouteno = " "
Txtstopno = " "
Txtstopname = " "



End Sub

Private Sub Command3_Click()
Me.Hide

End Sub

Private Sub Form_Load()
connectdb
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub

Private Sub Optno_Click()
Optno.Enabled = True
Optyes.Visible = False



End Sub

Private Sub Optyes_Click()
Optyes.Enabled = True
Optno.Visible = False
End Sub
