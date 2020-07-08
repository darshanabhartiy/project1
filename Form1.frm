VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8280
   ClientLeft      =   2955
   ClientTop       =   1080
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   10875
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   5775
      Left            =   1920
      TabIndex        =   1
      Top             =   1680
      Width           =   7695
      Begin VB.TextBox Txtestop 
         Height          =   375
         Left            =   4920
         TabIndex        =   15
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Txtbstop 
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Running Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Width           =   6615
         Begin VB.TextBox Txtetime 
            Height          =   375
            Left            =   4800
            TabIndex        =   19
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox Txtstime 
            Height          =   375
            Left            =   1440
            TabIndex        =   17
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0FFFF&
            Caption         =   "End Time"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            TabIndex        =   18
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Start Time"
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
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   1215
         End
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
         Left            =   4950
         TabIndex        =   10
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "save"
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
         Left            =   3030
         TabIndex        =   9
         Top             =   4800
         Width           =   1455
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
         Left            =   990
         Picture         =   "Form1.frx":0000
         TabIndex        =   8
         Top             =   4800
         Width           =   1695
      End
      Begin VB.TextBox Txtfare 
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox Txtstops 
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox Txtrnumber 
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ending Stop"
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
         Left            =   3480
         TabIndex        =   13
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Beginning Stop"
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
         TabIndex        =   12
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fare Stages"
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
         TabIndex        =   7
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Number of Stops"
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
         TabIndex        =   5
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Route Number"
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
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "M.S.R.T.C. ROUTE MANAGEMENT"
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
      Left            =   2400
      TabIndex        =   0
      Top             =   840
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String

Private Sub Command1_Click()

End Sub

Private Sub cmdexit_Click()
MsgBox ("Do you want to Exit")
Me.Hide

End Sub

Private Sub Cmdnew_Click()
MsgBox ("Do You want to clear")
Txtrnumber = " "
Txtstops = " "
Txtfare = " "
Txtrun = " "
Txtstime = ""
Txtetime = " "
Txtbstop = " "
Txtestop = " "



End Sub

Private Sub cmdsave_Click()
If Txtrnumber = "" Then
MsgBox "Please Enter the routenumber.", vbInformation
Txtrnumber.SetFocus
Exit Sub
End If

If Txtstops.Text = "" Then
MsgBox "Please Enter the Stop .", vbInformation
Txtstops.SetFocus
Exit Sub
End If

If Txtfare.Text = "" Then
MsgBox "Please Enter the fare .", vbInformation
Txtfare.SetFocus
Exit Sub
End If

If Txtbstop.Text = "" Then
MsgBox "Please Enter the biginning stop .", vbInformation
Txtbstop.SetFocus
Exit Sub
End If


If Txtestop.Text = "" Then
MsgBox "Please Enter the Ending stop .", vbInformation
Txtestop.SetFocus
Exit Sub
End If

If Txtstime.Text = "" Then
MsgBox "Please Enter the Starting time .", vbInformation
Txtstime.SetFocus
Exit Sub
End If


If Txtetime.Text = "" Then
MsgBox "Please Enter the ending time .", vbInformation
Txtetime.SetFocus
Exit Sub
End If

''connectdb
''str = "select * from routemanagement"
''rs.CursorLocation = adUseClient
''rs.Open str, con
''rs.AddNew
''rs.Fields(0) = Val(Txtrnumber.Text)
''rs.Fields(1) = Val(Txtstops.Text)
''rs.Fields(2) = Val(Txtfare.Text)
''rs.Fields(3) = Val(Txtbstop.Text)
''rs.Fields(4) = Val(Txtestop.Text)
''rs.Fields(5) = Val(Txtstime.Text)
''rs.Fields(6) = Val(Txtetime.Text)
''rs.Update
''rs.Close
''MsgBox ("Do you want to save")
con.Execute ("insert into route values(" + Txtrnumber.Text + "," + Txtstops.Text + "," + Txtfare.Text + ", '" + Txtbstop.Text + "','" + Txtestop.Text + "'," + Txtstime.Text + "," + Txtetime.Text + ")")

MsgBox ("successfully saved")
End Sub

Private Sub Form_Load()
Call connectdb
End Sub

Private Sub txtrnumber_Change()
If KeyAscii = 13 Then
txtnumber.Text = UCase(txtnumber.Text)
txtnumber.SetFocus
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
con.Close

End Sub
