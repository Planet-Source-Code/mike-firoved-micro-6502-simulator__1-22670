VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Save "
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3855
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Text            =   "*.asm"
         Top             =   4020
         Width           =   3375
      End
      Begin VB.DirListBox Dir1 
         Height          =   2565
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   3375
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   900
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Default         =   -1  'True
         Height          =   375
         Left            =   1500
         TabIndex        =   2
         Top             =   4500
         Width           =   915
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1500
         TabIndex        =   1
         Top             =   4920
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Please enter a path and a filename"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   3555
      End
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FtC As Boolean
Private Sub Command1_Click()
    On Error Resume Next
    Open Dir1.Path & "\" & Trim(Text1.Text) For Output As #1
        Print #1, Replace(FrmMain.txtprog.Text, vbCrLf, vbTab & vbCrLf)
    Close #1
    Me.Hide
End Sub

Private Sub Command2_Click()
    Me.Hide
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Text1.Text = "*.asm"
    FtC = True
End Sub

Private Sub Text1_Click()
    If FtC = True Then SendKeys ("{HOME}+{END}")
    FtC = False
End Sub
