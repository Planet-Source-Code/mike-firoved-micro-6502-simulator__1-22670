VERSION 5.00
Begin VB.Form frmHelp2 
   Caption         =   "General Help"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHelp 
      Height          =   2595
      Index           =   8
      Left            =   1620
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   120
      Width           =   5175
   End
   Begin VB.TextBox txtHelp 
      Height          =   2595
      Index           =   7
      Left            =   1620
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   120
      Width           =   5175
   End
   Begin VB.TextBox txtHelp 
      Height          =   2595
      Index           =   6
      Left            =   1620
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   120
      Width           =   5175
   End
   Begin VB.TextBox txtHelp 
      Height          =   2595
      Index           =   5
      Left            =   1620
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   120
      Width           =   5175
   End
   Begin VB.TextBox txtHelp 
      Height          =   2595
      Index           =   4
      Left            =   1620
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   120
      Width           =   5175
   End
   Begin VB.TextBox txtHelp 
      Height          =   2595
      Index           =   3
      Left            =   1620
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   120
      Width           =   5175
   End
   Begin VB.TextBox txtHelp 
      Height          =   2595
      Index           =   2
      Left            =   1620
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   120
      Width           =   5175
   End
   Begin VB.TextBox txtHelp 
      Height          =   2595
      Index           =   1
      Left            =   1620
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   120
      Width           =   5175
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   180
      Top             =   2940
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2880
      Width           =   1395
   End
   Begin VB.TextBox txtHelp 
      Height          =   2595
      Index           =   0
      Left            =   1620
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.ListBox lstHelp 
      Height          =   2595
      ItemData        =   "FrmHelp2.frx":0000
      Left            =   120
      List            =   "FrmHelp2.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmHelp2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    With lstHelp
        .Clear
        .AddItem "CPU Registers"
        .AddItem "Memory"
        .AddItem "Inputs"
        .AddItem "Outputs"
        .AddItem "Troubleshooting"
        .AddItem ""
    End With
    txtHelp(0).ZOrder
    txtHelp(0).Text = "CPU Registers" & vbCrLf & vbCrLf & "There are three registers in this CPU sample, A, X, Y. They can hold 8 bit numbers to preform calculations. There also is a Processor status, each flag holds a 1 bit status marker. The 'N' (negative) flag is when the number in the A register is negative number (when bit 7 is a one). The next flag is a 'Z' (Zero) flag, this flag will be 1 when the  A register is zero (bit 0-7 is o). The last flag is the C (carry) flag, it's function is to change depending on the outcome of an math or bitwise function."
    txtHelp(4).Text = "Troubleshooting" & vbCrLf & vbCrLf & "Problem: I tryed to use these commands but they return errors: TXS, TSX, PHA, PLA, PHP, PLP, CLI, SEI, CLD, SED. These are valid 6502 commands, what could be the problem?" & vbCrLf & vbCrLf & _
                      "Answer: In this version of my simulator, I have not included these commands. So these opcodes may be valid for the 6502, but this version of the simulator does not have them." & vbCrLf & vbCrLf & vbCrLf
    txtHelp(2).Text = "Inputs" & vbCrLf & vbCrLf & "There are 8 (1bit) inputs that are tied to the memory location 052. When the input changes, the memory changes. So in essence, memory location 052 is the input buffer. All eight On/Off logic inputs are combined together in 1 8-bit number. So if input one (bit 0) is 1 and the input eight (bit 7) is one and all other inputs are zero, then memory location 052 will be [10000001] or {129} decimal. this means if you dont care if any other inputs are one or zero but you are looking for one input, you would need to do some logical manuvering with AND, EOR or ORA. The inputs can be controlled by checking the checkboxes that say Inputs."
    txtHelp(1).Text = "Memory" & vbCrLf & vbCrLf & "This Simulator has only 512 bytes of ram, most of the old 6502 machines had 16 to 64 K bytes but for this simulator I will keep it to half of a K for simplicity. Also remember that this memory will not be holding your program, it will not have any subroutines nor any other usage other than the two bytes for input and output. If you would like to edit a certain cell, just double-click on it. There are two memory locations that are noteworthy, 051 and 052. Memory location 051 outputs directly to a 8 bit LED array (virtual of course). Any Value that is in location 051 is also shown on the LED array. Memory location 052 is a direct input for eight (virtual) toggle switches, and the value of those eight switches is tied into the location 052." & _
                      "This makes location 052 a dangerous place to store data. On the other hand, location 51 will never change unless you change it."
    txtHelp(3).Text = "Outputs" & vbCrLf & vbCrLf & "There is only one output, the LED (Light Emitting Diode) array. Each bit on the memory location 51 is tied to one LED of the [simulated] LED array. This allows you to view the output of the cpu in easy to read form."
End Sub




Private Sub Timer1_Timer()
    If (Not lstHelp.Text = lstHelp.Tag) Then
        lstHelp.Tag = lstHelp.Text
        Select Case lstHelp.Text
            Case "CPU Registers"
                aa = 0
            Case "Memory"
                aa = 1
            Case "Inputs"
                aa = 2
            Case "Outputs"
                aa = 3
            Case "Troubleshooting"
                aa = 4
            Case Else
                aa = 8
        End Select
'(Not Trim(lstHelp.Text) = "") And
        txtHelp(aa).ZOrder
    End If
End Sub
