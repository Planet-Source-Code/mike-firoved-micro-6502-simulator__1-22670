VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Print"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3075
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4575
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Default         =   -1  'True
         Height          =   375
         Left            =   3480
         TabIndex        =   8
         Top             =   2580
         Width           =   915
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   2580
         Width           =   915
      End
      Begin VB.CheckBox Check1 
         Caption         =   "All of the Above"
         Height          =   255
         Index           =   4
         Left            =   300
         TabIndex        =   5
         Top             =   2220
         Width           =   2475
      End
      Begin VB.CheckBox Check1 
         Caption         =   "CPU registers etc."
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   4
         Top             =   1680
         Width           =   2475
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Input / Output States"
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   3
         Top             =   1380
         Width           =   2475
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contents of the Memory"
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   2
         Top             =   1080
         Width           =   2475
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Program Code"
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   780
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.Label Label1 
         Caption         =   "What would you like to print?"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3795
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click(Index As Integer)
    If Index = 4 Then
        For aa = 0 To 3
            Check1(aa).Value = Check1(4).Value
        Next aa
    End If
End Sub

Private Sub Command1_Click()
    Dim aa As String
    
    If Check1(0).Value = 1 Then
        aa = "Program Code:" & vbCrLf & "--------------" & vbCrLf & FrmMain.txtprog.Text & vbCrLf & vbCrLf & vbCrLf
    End If
    
    If Check1(2).Value = 1 Then
        For cc = 0 To 7
            dd = "1"
            If FrmMain.LEDarray(cc).FillColor = vbBlack Then dd = "0"
            bb = bb & "Input #" & CStr(cc + 1) & "= " & dd & vbCrLf
        Next cc
        bb = bb & vbCrLf
        For cc = 0 To 7
            dd = CStr(FrmMain.InputArray(cc).Value)
            bb = bb & "Output #" & CStr(cc + 1) & "= " & dd & vbCrLf
        Next cc
        aa = aa & "Input / Outputs:" & vbCrLf & "--------------" & vbCrLf & bb & vbCrLf & vbCrLf & vbCrLf
    End If
    bb = ""
    If Check1(3).Value = 1 Then
        bb = bb & "A: " & FrmMain.txtA.Text & vbCrLf
        bb = bb & "X: " & FrmMain.txtX.Text & vbCrLf
        bb = bb & "Y: " & FrmMain.txtY.Text & vbCrLf
        bb = bb & "Zero: " & FrmMain.txtZ.Text & vbCrLf
        bb = bb & "Neg: " & FrmMain.txtN.Text & vbCrLf
        bb = bb & "Carry: " & FrmMain.txtC.Text & vbCrLf
        aa = aa & "CPU Registers, Etc:" & vbCrLf & "--------------" & vbCrLf & bb & vbCrLf & vbCrLf & vbCrLf
    End If
    bb = ""
    If Check1(1).Value = 1 Then
        For cc = 0 To 99
            If Right(FrmMain.lstMem.List(cc), 8) <> "00000000" Then bb = bb & FrmMain.lstMem.List(cc) & vbCrLf
        Next cc
        aa = aa & "Contents of Non-zero Memory:" & vbCrLf & "--------------" & vbCrLf & bb & vbCrLf & vbCrLf & vbCrLf
    End If

    Printer.Print ""
    Printer.Print aa
    Printer.NewPage
    Printer.EndDoc
    Me.Hide
End Sub

Private Sub Command2_Click()
    Me.Hide
End Sub
