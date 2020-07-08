VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form5"
   ClientHeight    =   9210
   ClientLeft      =   3435
   ClientTop       =   1665
   ClientWidth     =   11310
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   11310
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00C0FFFF&
      Height          =   7935
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   9015
      Begin VB.ComboBox cmbRtNum 
         Height          =   315
         Left            =   2040
         TabIndex        =   32
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox Txtchild 
         Height          =   375
         Left            =   2040
         TabIndex        =   31
         Top             =   3840
         Width           =   2295
      End
      Begin VB.ComboBox Cmdbusnumber 
         Height          =   315
         Left            =   2040
         TabIndex        =   30
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox Txttotal 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1920
         TabIndex        =   28
         Top             =   5760
         Width           =   6015
      End
      Begin VB.TextBox Txthalf 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1440
         TabIndex        =   26
         Top             =   4680
         Width           =   2055
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "No of Person[Child]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   24
         Top             =   3480
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "No of Person[Adult]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4680
         TabIndex        =   23
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox Txtenstop 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5760
         TabIndex        =   22
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox Txtsfrom 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2040
         TabIndex        =   21
         Top             =   3120
         Width           =   2295
      End
      Begin VB.ComboBox Cmbbustype 
         Height          =   315
         Left            =   2040
         TabIndex        =   18
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox Txtnoadult 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6480
         TabIndex        =   15
         Top             =   3840
         Width           =   2295
      End
      Begin VB.TextBox Txtfull 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6000
         TabIndex        =   13
         Top             =   4680
         Width           =   2415
      End
      Begin VB.CommandButton cmdrate 
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   11
         Top             =   6960
         Width           =   1575
      End
      Begin VB.CommandButton Cmdclear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   10
         Top             =   6960
         Width           =   1335
      End
      Begin VB.TextBox Txtbegstop 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CommandButton cmdprint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   5
         Top             =   6960
         Width           =   1455
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
         Height          =   495
         Left            =   6000
         TabIndex        =   4
         Top             =   6960
         Width           =   1455
      End
      Begin VB.TextBox Txtendto 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5760
         TabIndex        =   3
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Bus Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Half"
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
         Left            =   240
         TabIndex        =   27
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Full"
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
         Left            =   4920
         TabIndex        =   25
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "End To"
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
         Left            =   4560
         TabIndex        =   20
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Start From"
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
         TabIndex        =   19
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Lbldate 
         BackColor       =   &H00C0FFFF&
         Height          =   615
         Left            =   5280
         TabIndex        =   17
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Lbltime 
         BackColor       =   &H00C0FFFF&
         Height          =   615
         Left            =   1560
         TabIndex        =   16
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Total Fare"
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
         Left            =   360
         TabIndex        =   14
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Date"
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
         Left            =   5160
         TabIndex        =   12
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Begining Stop"
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
         TabIndex        =   8
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Bus Type"
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
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
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
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "End Stop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   2
         Top             =   2400
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   " M.S.R.T.C.  TICKETING"
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
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   6495
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c, a, t As Integer
Private Sub Command1_Click()
Text1 = " "
Text2 = " "
Text3 = " "
Text4 = " "
Text5 = " "
Text6 = " "
Text7 = " "

End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Cmdclear_Click()
Txtbegstop = " "
Txtendto = " "
Txtenstop = " "
Txtfull = " "
Txthalf = " "
Txtnoadult = " "
Txtnochild = " "
Txtsfrom = " "
Txttotal = " "
Lblmin = " "
Txtchild.Text = " "
End Sub

Private Sub cmdexit_Click()
MsgBox ("Do you want to Exit")
Me.Hide

End Sub

Private Sub cmdprint_Click()
MsgBox ("sorry no device connected")
End Sub

Private Sub cmdrate_Click()
Set rs = con.Execute("select Childfare,Adultfare from busdetails where Start='" + Txtsfrom.Text + "' and send='" + Txtendto.Text + "'  ")
If (Not rs.EOF) Then
    Txthalf.Text = rs(0)
    Txtfull.Text = rs(1)
    c = Val(Txtchild.Text) * Val(Txthalf.Text)
    a = Val(Txtnoadult.Text) * Val(Txtfull.Text)
   Txttotal.Text = c + a
    
Else
    MsgBox "Invalid Input", vbCritical, "E-Ticketing"
End If
rs.Close
End Sub

Private Sub Form_Load()
connectdb

Set rs = con.Execute("select Bustype from busdetails")
While (Not rs.EOF)
    Cmbbustype.AddItem rs(0)
     rs.MoveNext
     
Wend
rs.Close

Set rs = con.Execute("select Mincharge from busdetails where Bustype='" + Cmbbustype + "' ")
If (Not rs.EOF) Then
    Lblmin.Caption = rs(2)
End If
rs.Close

Set rs = con.Execute("select Busnumber from busdetails")
While (Not rs.EOF)
Cmdbusnumber.AddItem rs(0)
rs.MoveNext
Wend
rs.Close

Set rs = con.Execute("select RtNo from route")
While (Not rs.EOF)
    cmbRtNum.AddItem rs(0)
    rs.MoveNext
Wend
rs.Close
End Sub


