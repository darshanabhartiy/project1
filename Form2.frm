VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   8985
   ClientLeft      =   3960
   ClientTop       =   1515
   ClientWidth     =   10110
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   10110
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   6855
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   8415
      Begin VB.TextBox Txtend 
         Height          =   495
         Left            =   2880
         TabIndex        =   23
         Top             =   5040
         Width           =   2415
      End
      Begin VB.TextBox Txtstart 
         Height          =   375
         Left            =   2880
         TabIndex        =   21
         Top             =   4440
         Width           =   2415
      End
      Begin VB.TextBox Txtadultfare 
         Height          =   375
         Left            =   2880
         TabIndex        =   19
         Top             =   3960
         Width           =   2415
      End
      Begin VB.TextBox Txtchildfare 
         Height          =   375
         Left            =   2880
         TabIndex        =   18
         Top             =   3360
         Width           =   2415
      End
      Begin VB.TextBox Txtfare 
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox Txtdepot 
         Height          =   375
         Left            =   2880
         TabIndex        =   16
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox Txtmincharge 
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Txtbno 
         Height          =   375
         Left            =   2880
         TabIndex        =   14
         Top             =   840
         Width           =   2415
      End
      Begin VB.CommandButton cmdnew 
         Caption         =   "New"
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
         Left            =   360
         TabIndex        =   5
         Top             =   6000
         Width           =   1815
      End
      Begin VB.CommandButton cmdsave 
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
         Left            =   2760
         TabIndex        =   4
         Top             =   6000
         Width           =   1575
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "Exit"
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
         Left            =   4800
         TabIndex        =   3
         Top             =   6000
         Width           =   1695
      End
      Begin VB.ComboBox Cmdbustype 
         Height          =   315
         Left            =   2880
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "End Stop"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   22
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Start Stop"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Bus Number"
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
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Minimum Charge"
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
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Depot"
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
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fare Increment"
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
         Left            =   240
         TabIndex        =   10
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Adult Fare"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Child Fare"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Width           =   15
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Bus Type"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   " M.S.R.T.C.  BUS DETAILS"
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
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   600
      Width           =   6375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
MsgBox ("Do you want to Exit")
Me.Hide
End Sub

Private Sub Cmdnew_Click()
MsgBox ("Do You want to Clear")
Txtbno = " "
Txtfare = " "
Txtdepot = " "
Txtadultfare = " "
Txtchildfare = " "
Txtmincharge = " "

End Sub
Private Sub cmdsave_Click()

If Cmdbustype = "" Then
MsgBox "Please select bustype.", vbInformation
Cmdbustype.SetFocus
Exit Sub
End If


If Txtbno.Text = "" Then
MsgBox "Please select bus Number.", vbInformation
Txtbno.SetFocus
Exit Sub
End If



If Txtfare.Text = "" Then
MsgBox "Please select bus Fare.", vbInformation
Txtfare.SetFocus
Exit Sub
End If



If Txtdepot.Text = "" Then
MsgBox "Please select bus Depot.", vbInformation
Txtdepot.SetFocus
Exit Sub
End If



If Txtadultfare.Text = "" Then
MsgBox "Please select Adult fare.", vbInformation
Txtadultfare.SetFocus
Exit Sub
End If


If Txtchildfare.Text = "" Then
MsgBox "Please select Child fare.", vbInformation
Txtchildfare.SetFocus
Exit Sub
End If



If Txtmincharge.Text = "" Then
MsgBox "Please select Mincharge.", vbInformation
Txtmincharge.SetFocus
Exit Sub
End If

con.Execute ("insert into busdetails values('" + Cmdbustype.Text + "'," + Txtbno.Text + "," + Txtmincharge.Text + ", '" + Txtdepot.Text + "'," + Txtfare.Text + "," + Txtchildfare.Text + "," + Txtadultfare.Text + ",'" + Txtstart + "','" + Txtend + "')")
MsgBox ("successfully saved")

End Sub

Private Sub Form_Load()
Call connectdb
Cmdbustype.AddItem "Ordinary"
Cmdbustype.AddItem "Express"
Cmdbustype.AddItem "Super Fast"
Cmdbustype.AddItem "Fast"
Cmdbustype.AddItem "A\C Volvo"
Cmdbustype.AddItem "SemiSleeper Volvo"
Cmdbustype.AddItem "A\C SemiSleeper Volvo"
''Set rs = con.Execute("select * from busdetails")
''While (Not rs.EOF)
   ''Cmdbustype.AddItem rs(0)
    ''rs.MoveNext
''Wend
''rs.Close
End Sub


Private Sub Form_Unload(Cancel As Integer)
con.Close

End Sub


