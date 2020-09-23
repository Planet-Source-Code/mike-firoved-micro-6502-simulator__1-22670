VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "MF6502 Simulator"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5205
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtV 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   41
      Text            =   "0"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox txtYR 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   40
      Top             =   1500
      Width           =   495
   End
   Begin VB.TextBox txtXR 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   39
      Top             =   1140
      Width           =   495
   End
   Begin VB.TextBox txtAR 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   38
      Top             =   780
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop"
      Height          =   315
      Left            =   900
      TabIndex        =   37
      Top             =   4680
      Width           =   615
   End
   Begin VB.ListBox lstStat 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2700
      ItemData        =   "FrmMain.frx":0ECA
      Left            =   3840
      List            =   "FrmMain.frx":0ECC
      TabIndex        =   2
      Top             =   2280
      Width           =   1275
   End
   Begin VB.TextBox txtC 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   960
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "0"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox txtZ 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   600
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "0"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox txtN 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   240
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "0"
      Top             =   360
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   -60
      Top             =   5040
   End
   Begin VB.CheckBox InputArray 
      Caption         =   "bit 8"
      Height          =   195
      Index           =   7
      Left            =   3000
      TabIndex        =   16
      Top             =   1560
      Width           =   615
   End
   Begin VB.CheckBox InputArray 
      Caption         =   "bit 7"
      Height          =   195
      Index           =   6
      Left            =   3000
      TabIndex        =   15
      Top             =   1320
      Width           =   615
   End
   Begin VB.CheckBox InputArray 
      Caption         =   "bit 6"
      Height          =   195
      Index           =   5
      Left            =   3000
      TabIndex        =   14
      Top             =   1080
      Width           =   615
   End
   Begin VB.CheckBox InputArray 
      Caption         =   "bit 5"
      Height          =   195
      Index           =   4
      Left            =   3000
      TabIndex        =   13
      Top             =   840
      Width           =   615
   End
   Begin VB.CheckBox InputArray 
      Caption         =   "bit 4"
      Height          =   195
      Index           =   3
      Left            =   2340
      TabIndex        =   12
      Top             =   1560
      Width           =   615
   End
   Begin VB.CheckBox InputArray 
      Caption         =   "bit 3"
      Height          =   195
      Index           =   2
      Left            =   2340
      TabIndex        =   11
      Top             =   1320
      Width           =   615
   End
   Begin VB.CheckBox InputArray 
      Caption         =   "bit 2"
      Height          =   195
      Index           =   1
      Left            =   2340
      TabIndex        =   10
      Top             =   1080
      Width           =   615
   End
   Begin VB.CheckBox InputArray 
      Caption         =   "bit 1"
      Height          =   195
      Index           =   0
      Left            =   2340
      TabIndex        =   9
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   315
      Left            =   180
      TabIndex        =   22
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox txtprog 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2280
      Width           =   1395
   End
   Begin VB.ListBox lstMem 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2700
      ItemData        =   "FrmMain.frx":0ECE
      Left            =   1620
      List            =   "FrmMain.frx":0ED5
      TabIndex        =   1
      Top             =   2280
      Width           =   2115
   End
   Begin VB.TextBox txtY 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   240
      MaxLength       =   8
      TabIndex        =   8
      Top             =   1500
      Width           =   1155
   End
   Begin VB.TextBox txtX 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   240
      MaxLength       =   8
      TabIndex        =   7
      Top             =   1140
      Width           =   1155
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   240
      MaxLength       =   8
      TabIndex        =   6
      Top             =   780
      Width           =   1155
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   42
      Top             =   60
      Width           =   195
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "Status"
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
      Left            =   3840
      TabIndex        =   36
      Top             =   1980
      Width           =   1275
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   35
      Top             =   60
      Width           =   195
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   34
      Top             =   60
      Width           =   195
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "N"
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
      Left            =   240
      TabIndex        =   33
      Top             =   60
      Width           =   195
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "8Bit input into memory location 052"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   32
      Top             =   180
      Width           =   1155
   End
   Begin VB.Label Label15 
      Caption         =   "8"
      Height          =   195
      Left            =   4740
      TabIndex        =   31
      Top             =   1620
      Width           =   195
   End
   Begin VB.Label Label14 
      Caption         =   "7"
      Height          =   195
      Left            =   4740
      TabIndex        =   30
      Top             =   1380
      Width           =   195
   End
   Begin VB.Label Label13 
      Caption         =   "6"
      Height          =   195
      Left            =   4740
      TabIndex        =   29
      Top             =   1140
      Width           =   195
   End
   Begin VB.Label Label12 
      Caption         =   "4"
      Height          =   195
      Left            =   4200
      TabIndex        =   28
      Top             =   1620
      Width           =   195
   End
   Begin VB.Label Label11 
      Caption         =   "3"
      Height          =   195
      Left            =   4200
      TabIndex        =   27
      Top             =   1380
      Width           =   195
   End
   Begin VB.Label Label10 
      Caption         =   "2"
      Height          =   195
      Left            =   4200
      TabIndex        =   26
      Top             =   1140
      Width           =   195
   End
   Begin VB.Label Label9 
      Caption         =   "1"
      Height          =   195
      Left            =   4200
      TabIndex        =   25
      Top             =   900
      Width           =   195
   End
   Begin VB.Label Label8 
      Caption         =   "5"
      Height          =   195
      Left            =   4740
      TabIndex        =   24
      Top             =   900
      Width           =   195
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "8Bit output from memory location 051"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3900
      TabIndex        =   23
      Top             =   240
      Width           =   1095
   End
   Begin VB.Shape LEDarray 
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   7
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   135
   End
   Begin VB.Shape LEDarray 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   6
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   135
   End
   Begin VB.Shape LEDarray 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   135
   End
   Begin VB.Shape LEDarray 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape LEDarray 
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   4020
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   135
   End
   Begin VB.Shape LEDarray 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   4020
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   135
   End
   Begin VB.Shape LEDarray 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   4020
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   135
   End
   Begin VB.Shape LEDarray 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   4020
      Shape           =   3  'Circle
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Memory"
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
      Left            =   1680
      TabIndex        =   21
      Top             =   1980
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Program Code"
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
      Left            =   120
      TabIndex        =   20
      Top             =   1980
      Width           =   1395
   End
   Begin VB.Label Label3 
      Caption         =   "Y"
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
      Left            =   60
      TabIndex        =   19
      Top             =   1500
      Width           =   195
   End
   Begin VB.Label Label2 
      Caption         =   "X"
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
      Left            =   60
      TabIndex        =   18
      Top             =   1140
      Width           =   195
   End
   Begin VB.Label Label1 
      Caption         =   "A"
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
      Left            =   60
      TabIndex        =   17
      Top             =   780
      Width           =   195
   End
   Begin VB.Menu File1 
      Caption         =   "&File"
      Begin VB.Menu new1 
         Caption         =   "&New"
      End
      Begin VB.Menu save1 
         Caption         =   "&Save"
      End
      Begin VB.Menu ssar 
         Caption         =   "-"
      End
      Begin VB.Menu print1 
         Caption         =   "&Print"
      End
      Begin VB.Menu sdfgsdfg 
         Caption         =   "-"
      End
      Begin VB.Menu samples1 
         Caption         =   "Samples"
         Begin VB.Menu inputsample1 
            Caption         =   "Get Input Sample"
         End
         Begin VB.Menu loopsamp1 
            Caption         =   "Looping Sample"
         End
         Begin VB.Menu mult 
            Caption         =   "Multiply by Ten Sample"
         End
         Begin VB.Menu logicsamp1 
            Caption         =   "Logic Sample"
         End
         Begin VB.Menu branck1 
            Caption         =   "Branching Sample"
         End
      End
      Begin VB.Menu dfhgh 
         Caption         =   "-"
      End
      Begin VB.Menu exit1 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu edit1 
      Caption         =   "&Edit"
      Begin VB.Menu undo1 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu dsdf 
         Caption         =   "-"
      End
      Begin VB.Menu cut1 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu copyt1 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste2 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu SDF 
         Caption         =   "-"
      End
      Begin VB.Menu reset1 
         Caption         =   "&Reset Simulator"
      End
      Begin VB.Menu SDFGSDF 
         Caption         =   "-"
      End
      Begin VB.Menu cpu1 
         Caption         =   "Change CPU time"
      End
   End
   Begin VB.Menu help1 
      Caption         =   "&Help"
      Begin VB.Menu ghelp1 
         Caption         =   "General Help"
      End
      Begin VB.Menu opcode1 
         Caption         =   "Op Codes"
      End
      Begin VB.Menu about1 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public sleeptime As Double
Public BreakProg As Boolean
Public FmP As Boolean

Private Sub cmdHelp_Click()
    FrmHelp.Show
End Sub

Private Sub about1_Click()
    frmAbout.Show
End Sub

Private Sub branck1_Click()
    Timer1.Enabled = False
    Screen.MousePointer = 11
    lstMem.Visible = False
    Form_Load
    txtprog.Text = "CLA" & vbCrLf & "STA 51" & vbCrLf & "LDA 52" & vbCrLf & "BEQ 1" & vbCrLf & "BRK" & vbCrLf & "LDA #1" & vbCrLf & "STA 51"
    For aa = 0 To 7
        InputArray(aa).Value = 0
    Next aa
    Screen.MousePointer = 0
    lstMem.Visible = True
End Sub

Private Sub cmdRun_Click()
    
    Screen.MousePointer = vbArrowHourglass
    aa = executeCommands(txtprog.Text)
    Screen.MousePointer = 0

End Sub

Private Sub Command1_Click()
    BreakProg = True
End Sub

Private Sub copyt1_Click()
    SendKeys ("^C")
End Sub

Private Sub cpu1_Click()
    aaa = InputBox("The interval between each command can be changed in order for you to be able to see the actions slower or faster. For example, If you need to run a long program and just see the output then make the cpu time short and the program will run straight through. If you wanted to see every command executes and see what the registers say after each op then make the CPU time large. in any event the time is in units of milliseconds. One second is 1000 ms, 750 is a good rate of speed and is the default.", "CPU Speed", CStr(sleeptime))
    If Not aaa = "" Then sleeptime = Val(aaa)
End Sub

Private Sub cut1_Click()
    SendKeys ("^X")
End Sub

Private Sub exit1_Click()
    End
End Sub

Private Sub Form_Load()
        'INIT ALL DEFAULT VALUES

        'flash all leds
        For aa = 0 To 7
            LEDarray(aa).FillColor = vbBlack
        Next aa
        
        'make all memory values
        lstMem.Clear
        For aa = 1 To 512
            lstMem.AddItem Right("000" & CStr(aa), 3) & ": 00000000"
            DoEvents
        Next aa
        
        lstStat.Clear
        txtprog.Text = ""
        
        'clear all inputs
        For aa = 0 To 7
            InputArray(aa).Value = 0
        Next aa
        sleeptime = 750
        
        'clear a,y,x
        txtA.Text = "00000000"
        txtY.Text = "00000000"
        txtX.Text = "00000000"
        txtN.Text = "0"
        txtZ.Text = "0"
        txtC.Text = "0"
        'logicsamp1_Click
        Timer1.Enabled = True
End Sub



Private Sub ghelp1_Click()
    frmHelp2.Show
End Sub

Private Sub InputArray_Click(Index As Integer)
    For aa = 7 To 0 Step -1
        bb = bb & IIf(InputArray(aa).Value = 1, "1", "0")
    Next aa
    lstMem.List(51) = "052: " & bb
    lstMem.Selected(51) = True
End Sub



Private Sub inputsample1_Click()
    Timer1.Enabled = False
    Screen.MousePointer = 11
    lstMem.Visible = False
    Form_Load
    txtprog = "LDA $#FF" & vbCrLf & "STA 51" & vbCrLf & "CLA" & vbCrLf & "STA 51" & vbCrLf & "LDA 52" & vbCrLf & "STA 51"
    InputArray(1).Value = 1
    InputArray(3).Value = 1
    InputArray(5).Value = 1
    InputArray(7).Value = 1
    lstMem.Visible = True
    Screen.MousePointer = 0
End Sub

Private Sub logicsamp1_Click()
    Timer1.Enabled = False
    lstMem.Visible = False
    Screen.MousePointer = 11
    Form_Load
    txtprog = "CLC" & vbCrLf & "LDA #223" & vbCrLf & "AND $#FF" & vbCrLf & "STA 51" & vbCrLf & "STA 10" & vbCrLf & "NOP" & vbCrLf & "CLC" & vbCrLf & "LDA #0" & vbCrLf & "STA 51" & vbCrLf & "LDA #239" & vbCrLf & "ORA #15" & vbCrLf & "STA 51" & vbCrLf & "STA 11" & vbCrLf & "NOP" & vbCrLf & "CLC" & vbCrLf & "LDA #0" & vbCrLf & "STA 51" & vbCrLf & "LDA #248" & vbCrLf & "EOR #15" & vbCrLf & "STA 51" & vbCrLf & "STA 12" & vbCrLf & "LDA $#01" & vbCrLf & "STA $33"
    Screen.MousePointer = 0
    lstMem.Visible = True
End Sub

Private Sub loopsamp1_Click()
    Timer1.Enabled = False
    lstMem.Visible = False
    Screen.MousePointer = 11
    Form_Load
    txtprog = "LDA #16" & vbCrLf & "LDY #0" & vbCrLf & "LDX #0" & vbCrLf & "STY 1, X" & vbCrLf & "INX" & vbCrLf & "INY" & vbCrLf & "DEC" & vbCrLf & "BNE -5" & vbCrLf & "LDA #255" & vbCrLf & "STA 51"
    Screen.MousePointer = 0
    lstMem.Visible = True
End Sub

Private Sub lstMem_DblClick()
    aa = lstMem.Text
    If Left(aa, 3) = "052" Then Exit Sub
    bb = InputBox("What Value would you like to give the memory space " & Left(aa, 3) & "?", "MEMORY VALUE CHANGE", "00000000")
    If bb = "" Then Exit Sub
    cc = Val(Left(aa, 3)) - 1
    lstMem.List(cc) = Left(aa, 3) & ": " & Right("00000000" & bb, 8)
    
End Sub

Private Sub mult_Click()
    Timer1.Enabled = False
    Screen.MousePointer = 11
    lstMem.Visible = False
    Form_Load
    txtprog.Text = "LDA 52" & vbCrLf & "ASL" & vbCrLf & "STA 1" & vbCrLf & "ASL" & vbCrLf & "ASL" & vbCrLf & "CLC" & vbCrLf & "ADC 1" & vbCrLf & "STA 51"
    InputArray(0).Value = 1
    InputArray(3).Value = 1
    Screen.MousePointer = 0
    lstMem.Visible = True
End Sub

Private Sub new1_Click()
    aa = MsgBox("Are you sure you want to close the current Project?" & vbCrLf & "This will remove all memory, program and registers information." & vbCrLf & vbCrLf & "Do you want to Continue?", vbYesNoCancel + vbDefaultButton3 + vbExclamation, "New project?")
    If aa = vbYes Then
        Timer1.Enabled = False
        Form_Load
    End If
End Sub

Private Sub opcode1_Click()
    FrmHelp.Show
End Sub



Private Sub paste2_Click()
    SendKeys ("^V")
End Sub

Private Sub print1_Click()
    frmPrint.Show
End Sub

Private Sub reset1_Click()
    aa = MsgBox("Are you sure you want to reset?" & vbCrLf & "This will remove all memory, program and registers information." & vbCrLf & vbCrLf & "Do you want to reset?", vbYesNoCancel + vbDefaultButton3 + vbExclamation, "Reset All?")
    If aa = vbYes Then
        Timer1.Enabled = False
        lstMem.Visible = False
        Form_Load
        lstMem.Visible = True
    End If
End Sub

Private Sub save1_Click()
    frmSave.Show
End Sub

Private Sub Timer1_Timer()
    updateLEDs
    txtAR.Text = CStr(BinToDec(txtA.Text))
    txtXR.Text = CStr(BinToDec(txtX.Text))
    txtYR.Text = CStr(BinToDec(txtY.Text))
    If FmP = True Then
        txtZ = IIf(Val(txtA.Text) = 0, "1", "0")
        txtN = Left(txtA.Text, 1)
    End If
End Sub

Public Sub updateLEDs()
    On Error Resume Next
    For aa = 0 To 7 ' Step -1
        bb = Mid(lstMem.List(50), aa + 6, 1)
        LEDarray(7 - aa).FillColor = IIf(bb, vbRed, vbBlack)
    Next aa
End Sub


Public Function executeCommands(commands As String) As Boolean
    cmds = Split(commands, vbCrLf)
    bb = UBound(cmds)

    
    lstStat.Clear
    For aa = 0 To bb
        lstStat.AddItem cmds(aa)
    Next aa
    
    For aa = 0 To bb
        If BreakProg = True Then
            BreakProg = False
            lstStat.Selected(0) = True
            lstStat.Selected(0) = False
            Exit Function
        End If
        lstStat.Selected(aa) = True
        DoEvents
        'txtcur.Text = UCase(Left(cmds(aa), 3))
        DoEvents
        If InStr(1, cmds(aa), "$") > 0 Then
            If InStr(1, cmds(aa), "#") > 0 Then
                xc = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), "#"))
                xf = Left(cmds(aa), 3) & " #" & CStr(HexToDec(xc))
                cmds(aa) = xf
            Else
                xc = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), "$"))
                xf = Left(cmds(aa), 3) & " " & CStr(HexToDec(xc))
                cmds(aa) = xf
            End If
        End If
        Select Case Trim(UCase(Left(cmds(aa), 3)))
        
            Case "TAY"
                za = txtA.Text
                Sleep sleeptime
                txtY.Text = za
                FmP = True
                
            Case "TAX"
                za = txtA.Text
                Sleep sleeptime
                txtX.Text = za
                FmP = True
                
            Case "TXA"
                zx = txtX.Text
                Sleep sleeptime
                txtA.Text = zx
                FmP = True
                
            Case "TYA"
                zy = txtY.Text
                Sleep sleeptime
                txtA.Text = zy
                FmP = True
                
            Case "LDA"
                If InStr(1, Trim(Mid(cmds(aa), 4, 2)), "#") > 0 Then
                    xc = Val(Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), "#")))
                    txtA.Text = DecToBin(xc)
                Else
                    xc = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                    txtA.Text = Right(lstMem.List(Val(xc) - 1), 8)
                End If
                Sleep sleeptime
                FmP = True
                
            Case "LDX"
                If InStr(1, Trim(Mid(cmds(aa), 4, 2)), "#") > 0 Then
                    xc = Val(Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), "#")))
                    txtX.Text = DecToBin(xc)
                Else
                    xc = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                    txtX.Text = Right(lstMem.List(Val(xc) - 1), 8)
                End If
                Sleep sleeptime
                FmP = True
                
            Case "LDY"
                If InStr(1, Trim(Mid(cmds(aa), 4, 2)), "#") > 0 Then
                    xc = Val(Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), "#")))
                    txtY.Text = DecToBin(xc)
                Else
                    xc = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                    txtY.Text = Right(lstMem.List(Val(xc) - 1), 8)
                End If
                Sleep sleeptime
                FmP = True
                
            Case "ADC"
                If InStr(1, Trim(Mid(cmds(aa), 4, 2)), "#") > 0 Then
                    xc = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), "#"))
                    bb = txtA.Text
                    cc = BinToDec(bb)
                    cc = cc + Val(xc)

                Else
                    xc = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                    bb = Right(lstMem.List(Val(xc) - 1), 8)
                    cc = BinToDec(bb)
                    cc = cc + BinToDec(txtA.Text)
                End If
                If cc > 255 Then
                    txtC.Text = "1"
                    cc = cc - 255
                End If
                txtA.Text = DecToBin(cc)
                txtV.Text = CStr(Val(Trim(txtN.Text)) Xor Val(Trim(txtC.Text)))
                Sleep sleeptime
                FmP = True
                
                
            Case "SBC"
                If InStr(1, Trim(Mid(cmds(aa), 4, 2)), "#") > 0 Then
                    xc = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), "#"))
                    bb = txtA.Text
                    cc = BinToDec(bb)
                    cc = cc - Val(xc)
                Else
                    xc = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                    bb = Right(lstMem.List(Val(xc) - 1), 8)
                    cc = BinToDec(bb)
                    cc = cc - BinToDec(txtA.Text)
                End If
                If Val(xc) > cc Then
                    txtC.Text = "1"
                    cc = cc + 255
                End If
                txtA.Text = DecToBin(cc)
                txtV.Text = CStr(Val(Trim(txtN.Text)) Xor Val(Trim(txtC.Text)))
                Sleep sleeptime
                FmP = True
                
            Case "AND"
                If InStr(1, Trim(Mid(cmds(aa), 4, 2)), "#") > 0 Then
                    bb = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), "#"))
                    xc = Val(bb)
                Else
                    bb = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                    xc = BinToDec(Right(lstMem.List(Val(xc) - 1), 8))
                End If
                
                xf = BinToDec(txtA.Text)
                cc = xc And xf
                Sleep sleeptime
                txtA = DecToBin(cc)
                FmP = True
                
            Case "ORA"
                If InStr(1, Trim(Mid(cmds(aa), 4, 2)), "#") > 0 Then
                    bb = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), "#"))
                    xc = Val(bb)
                Else
                    bb = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                    xc = BinToDec(Right(lstMem.List(Val(xc) - 1), 8))
                End If
                
                xf = BinToDec(txtA.Text)
                cc = xc Or xf
                Sleep sleeptime
                txtA = DecToBin(cc)
                FmP = True
                
            Case "EOR"
                If InStr(1, Trim(Mid(cmds(aa), 4, 2)), "#") > 0 Then
                    bb = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), "#"))
                    xc = Val(bb)
                Else
                    bb = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                    xc = BinToDec(Right(lstMem.List(Val(xc) - 1), 8))
                End If
                
                xf = BinToDec(txtA.Text)
                cc = xc Xor xf
                Sleep sleeptime
                txtA = DecToBin(cc)
                FmP = True
                
'            Case "NOT"
'                bb = txtA.Text
'                bb = Replace(bb, "1", "2")
'                bb = Replace(bb, "0", "1")
'                bb = Replace(bb, "2", "0")
'                txtA.Text = bb
'                Sleep sleeptime
'                FmP = True
'
            Case "STA"
                If InStr(1, cmds(aa), ",") > 0 Then
                    cc = InStr(1, cmds(aa), ",")
                    bb = Right(cmds(aa), Len(cmds(aa)) - cc)
                    xp = Trim(txtY.Text)
                    If LCase(Trim(bb)) = "x" Then xp = Trim(txtX.Text)
                    xc = Mid(cmds(aa), InStr(1, cmds(aa), " "), cc - InStr(1, cmds(aa), " "))
                    xj = Right("00000000" & xp, 8)
                    xk = BinToDec(xj)
                    cc = Val(xc) + xk
                    xn = CStr(cc)
                Else
                    xn = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                End If
                lstMem.List(Val(xn) - 1) = Left(lstMem.List(Val(xn) - 1), 5) & txtA.Text
                lstMem.Selected(Val(xn) - 1) = True
                Sleep sleeptime
                FmP = True
                
            Case "STX"
                If InStr(1, cmds(aa), ",") > 0 Then
                    cc = InStr(1, cmds(aa), ",")
                    bb = Right(cmds(aa), Len(cmds(aa)) - cc)
                    xp = Trim(txtY.Text)
                    If LCase(Trim(bb)) = "x" Then xp = Trim(txtX.Text)
                    xc = Mid(cmds(aa), InStr(1, cmds(aa), " "), cc - InStr(1, cmds(aa), " "))
                    xj = Right("00000000" & xp, 8)
                    xk = BinToDec(xj)
                    cc = Val(xc) + xk
                    xn = CStr(cc)
                Else
                    xn = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                End If
                lstMem.List(Val(xn) - 1) = Left(lstMem.List(Val(xn) - 1), 5) & txtX.Text
                lstMem.Selected(Val(xn) - 1) = True
                Sleep sleeptime
                FmP = True
                
                
            Case "STY"
                If InStr(1, cmds(aa), ",") > 0 Then
                    cc = InStr(1, cmds(aa), ",")
                    bb = Right(cmds(aa), Len(cmds(aa)) - cc)
                    xp = Trim(txtY.Text)
                    If LCase(Trim(bb)) = "x" Then xp = Trim(txtX.Text)
                    xc = Mid(cmds(aa), InStr(1, cmds(aa), " "), cc - InStr(1, cmds(aa), " "))
                    xj = Right("00000000" & xp, 8)
                    xk = BinToDec(xj)
                    cc = Val(xc) + xk
                    xn = CStr(cc)
                Else
                    xn = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                End If
                lstMem.List(Val(xn) - 1) = Left(lstMem.List(Val(xn) - 1), 5) & txtY.Text
                lstMem.Selected(Val(xn) - 1) = True
                Sleep sleeptime
                FmP = True
                
            Case "INC"
                bb = txtA.Text
                cc = BinToDec(bb)
                If cc = 255 Then
                    cc = 0
                Else
                    cc = cc + 1
                End If
                Sleep sleeptime
                txtA.Text = DecToBin(cc)
                FmP = True
                
            Case "DEC"
                bb = txtA.Text
                cc = BinToDec(bb)
                If cc = 0 Then
                    cc = 255
                Else
                    cc = cc - 1
                End If
                Sleep sleeptime
                txtA.Text = DecToBin(cc)
                FmP = True
                
            Case "INX"
                bb = txtX.Text
                cc = BinToDec(bb)
                If cc = 255 Then
                    cc = 0
                Else
                    cc = cc + 1
                End If
                Sleep sleeptime
                txtX.Text = DecToBin(cc)
                FmP = True
                
            Case "DEX"
                bb = txtX.Text
                cc = BinToDec(bb)
                If cc = 0 Then
                    cc = 255
                Else
                    cc = cc - 1
                End If
                Sleep sleeptime
                txtX.Text = DecToBin(cc)
                FmP = True
                
            Case "INY"
                bb = txtY.Text
                cc = BinToDec(bb)
                If cc = 255 Then
                    cc = 0
                Else
                    cc = cc + 1
                End If
                Sleep sleeptime
                txtY.Text = DecToBin(cc)
                FmP = True
                
            Case "DEY"
                bb = txtY.Text
                cc = BinToDec(bb)
                If cc = 0 Then
                    cc = 255
                Else
                    cc = cc - 1
                End If
                Sleep sleeptime
                txtY.Text = DecToBin(cc)
                FmP = True
                
            Case "CLA"
                Sleep sleeptime
                txtA.Text = "00000000"
                FmP = True
                
            Case "CLX"
                Sleep sleeptime
                txtX.Text = "00000000"
                FmP = True
                
            Case "CLY"
                Sleep sleeptime
                txtY.Text = "00000000"
                FmP = True
                
            Case "NOP"
                Sleep sleeptime
                FmP = True
                
            Case "BPL"
                If Left(txtA.Text, 1) = "0" Then
                    bb = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                    cc = Val(bb)
                    aa = aa + cc
                End If
                Sleep sleeptime
                FmP = True
                
            Case "BMI"
                If Left(txtA.Text, 1) = "1" Then
                    bb = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                    cc = Val(bb)
                    aa = aa + cc
                End If
                Sleep sleeptime
                FmP = True
                
            Case "BNE"
                If txtA.Text <> "00000000" Then
                    bb = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                    cc = Val(bb)
                    aa = aa + cc
                End If
                Sleep sleeptime
                FmP = True
                
            Case "BEQ"
                If txtA.Text = "00000000" Then
                    bb = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                    cc = Val(bb)
                    aa = aa + cc
                End If
                Sleep sleeptime
                FmP = True
                
            Case "BCC"
                If txtC.Text = "0" Then
                    bb = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                    cc = Val(bb)
                    aa = aa + cc
                End If
                Sleep sleeptime
                FmP = True
                
            Case "BCS"
                If txtC.Text = "1" Then
                    bb = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                    cc = Val(bb)
                    aa = aa + cc
                End If
                Sleep sleeptime
                FmP = True
                
            Case "ASL"
                txtC.Text = Left(txtA.Text, 1)
                txtA.Text = Right(txtA.Text, 7) & "0"
                Sleep sleeptime
                FmP = True
                
            Case "LSR"
                txtC.Text = Right(txtA.Text, 1)
                txtA.Text = "0" & Left(txtA.Text, 7)
                Sleep sleeptime
                FmP = True
                
            Case "ROR"
                cr = Right(txtA.Text, 1)
                txtA.Text = txtC.Text & Left(txtA.Text, 7)
                txtC.Text = cr
                Sleep sleeptime
                FmP = True
                
            Case "ROL"
                cr = Left(txtA.Text, 1)
                txtA.Text = Right(txtA.Text, 7) & txtC.Text
                txtC.Text = cr
                Sleep sleeptime
                FmP = True
                
            Case "CLC"
                txtC.Text = "0"
                Sleep sleeptime
                FmP = True
                
            Case "SEC"
                txtC.Text = "1"
                Sleep sleeptime
                FmP = True
                
            Case "BRK"
                Sleep sleeptime
                Exit Function
                FmP = True
                
            Case "CMP"
                'GET MEMORY LOCATION
                    xn = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                
                'GET VALUE OF MEMORY
                    cc = Val(xn) - 1
                    bb = Right(lstMem.List(cc), 8)
                    cc = BinToDec(bb)
                    
                'GET VALUE OF COMPAREE
                    va = BinToDec(txtA.Text)
                    FmP = False
                
                'IS COMPAREE >, <  OR  = TO  MEM
                    If va < cc Then
                        txtN.Text = "1"
                        txtZ.Text = "0"
                        txtC.Text = "0"
                    End If
                    If va = cc Then
                        txtN.Text = "0"
                        txtZ.Text = "1"
                        txtC.Text = "1"
                    End If
                    If va > cc Then
                        txtN.Text = "0"
                        txtZ.Text = "0"
                        txtC.Text = "1"
                    End If
                Sleep sleeptime
                
            Case "CPX"
                'GET MEMORY LOCATION
                    xn = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                
                'GET VALUE OF MEMORY
                    cc = Val(xn) - 1
                    bb = Right(lstMem.List(cc), 8)
                    cc = BinToDec(bb)
                    
                'GET VALUE OF COMPAREE
                    va = BinToDec(txtX.Text)
                    FmP = False
                    
                'IS COMPAREE >, <  OR  = TO  MEM
                    If va < cc Then
                        txtN.Text = "1"
                        txtZ.Text = "0"
                        txtC.Text = "0"
                    End If
                    If va = cc Then
                        txtN.Text = "0"
                        txtZ.Text = "1"
                        txtC.Text = "2"
                    End If
                    If va > cc Then
                        txtN.Text = "0"
                        txtZ.Text = "0"
                        txtC.Text = "1"
                    End If
                Sleep sleeptime
                
                
            Case "CPY"
                'GET MEMORY LOCATION
                    xn = Right(cmds(aa), Len(cmds(aa)) - InStr(1, cmds(aa), " "))
                
                'GET VALUE OF MEMORY
                    cc = Val(xn) - 1
                    bb = Right(lstMem.List(cc), 8)
                    cc = BinToDec(bb)
                    
                'GET VALUE OF COMPAREE
                    va = BinToDec(txtY.Text)
                    FmP = False
                    
                'IS COMPAREE >, <  OR  = TO  MEM
                    If va < cc Then
                        txtN.Text = "1"
                        txtZ.Text = "0"
                        txtC.Text = "0"
                    End If
                    If va = cc Then
                        txtN.Text = "0"
                        txtZ.Text = "1"
                        txtC.Text = "2"
                    End If
                    If va > cc Then
                        txtN.Text = "0"
                        txtZ.Text = "0"
                        txtC.Text = "1"
                    End If
                Sleep sleeptime
                
            Case "CLV"
                'clear overflow
                txtV.Text = "0"
                Sleep sleeptime
                FmP = True
            
            Case ""
                'do nothing
            
            Case Else
                MsgBox "ERROR ON LINE" & CStr(aa + 1), vbExclamation + vbDefaultButton1 + vbOKCancel, "ERROR"
                Exit Function
        End Select
        DoEvents
        DoEvents
    Next aa
    lstMem.Selected(0) = True
End Function


Public Function BinToDec(ByVal BinStr As String) As Double
    Dim mult As Double
    Dim DecNum As Double
    mult = 1
    DecNum = 0
    
    Dim i As Integer
    For i = Len(BinStr) To 1 Step -1
        If Mid(BinStr, i, 1) = "1" Then
            DecNum = DecNum + mult
        End If
        mult = mult * 2
    Next i
    BinToDec = DecNum
End Function

Public Function DecToBin(ByVal DecNum As Double) As String
    Dim BinStr As String
    BinStr = ""
    Do While DecNum <> 0
        If (DecNum Mod 2) = 1 Then   'This method Blows!!!!!! Causes Overflow
            BinStr = "1" & BinStr
        Else
            BinStr = "0" & BinStr
        End If
        DecNum = DecNum \ 2
    Loop
    If BinStr = "" Then BinStr = "00000000"
    DecToBin = Right("00000000" & BinStr, 8)
End Function

Public Function HexToBin(ByVal HexStr As String) As String
    Dim BinStr As String
    BinStr = ""
    Dim i As Integer
    For i = 1 To Len(HexStr)
        BinStr = BinStr & DecToBin(HexToDec(Mid(HexStr, i, 1)))
    Next i
    HexToBin = BinStr
End Function

Public Function BinToHex(ByVal BinStr As String) As String
    Dim HexStr As String
    HexStr = ""
    Dim i As Integer
    For i = 1 To Len(BinStr) Step 4
        HexStr = HexStr & DecToHex(BinToDec(Mid(BinStr, i, 4)))
    Next i
    BinToHex = HexStr
End Function

Public Function HexToDec(ByVal HexStr As String) As Double
    Dim mult As Double
    Dim DecNum As Double
    Dim ch As String
    mult = 1
    DecNum = 0

    Dim i As Integer
    For i = Len(HexStr) To 1 Step -1
        ch = Mid(HexStr, i, 1)
        If (ch >= "0") And (ch <= "9") Then
            DecNum = DecNum + (Val(ch) * mult)
        Else
            If (ch >= "A") And (ch <= "F") Then
                DecNum = DecNum + ((Asc(ch) - Asc("A") + 10) * mult)
            Else
                If (ch >= "a") And (ch <= "f") Then
                    DecNum = DecNum + ((Asc(ch) - Asc("a") + 10) * mult)
                Else
                    HexToDec = 0
                    Exit Function
                End If
            End If
        End If
        mult = mult * 16
    Next i
    HexToDec = DecNum
End Function

Public Function DecToHex(ByVal DecNum As Double) As String
    Dim remainder As Integer
    Dim HexStr As String
    HexStr = ""
    Do While DecNum <> 0
        remainder = DecNum Mod 16   'This method Blows!!!!!! Causes Overflow
        If remainder <= 9 Then
            HexStr = Chr(Asc(remainder)) & HexStr
        Else
            HexStr = Chr(Asc("A") + remainder - 10) & HexStr
        End If
        DecNum = DecNum \ 16
    Loop
    If HexStr = "" Then HexStr = "0"
    DecToHex = HexStr
End Function





Private Sub undo1_Click()
    SendKeys ("^Z")
End Sub
