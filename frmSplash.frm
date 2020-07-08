VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5055
   ClientLeft      =   4800
   ClientTop       =   2445
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6840
      Top             =   3480
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   7035
      TabIndex        =   2
      Top             =   4560
      Width           =   7095
      Begin VB.Image Image1 
         Height          =   180
         Left            =   0
         Picture         =   "frmSplash.frx":000C
         Top             =   0
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Bus Route Management"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   720
         TabIndex        =   6
         Top             =   1440
         Width           =   5895
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Warning : Copyright Reserved 20016"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   3660
         Width           =   6855
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Loading Files......."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   4200
      Width           =   1935
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim x As Integer
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
   
    Load MDIForm1
    MDIForm1.Show
     Unload Me
End Sub

Private Sub Form_Load()
    'lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
   
    File1.FileName = App.Path
    x = File1.ListCount
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
If (Image1.Left <= 6600) Then
    Image1.Left = Image1.Left + 50
Else
    Image1.Left = 0
End If
If (i <= x) Then
    Label1.Caption = File1.List(i)
    i = i + 1
Else
Load MDIForm1
    MDIForm1.Show
    Unload Me
End If
End Sub
