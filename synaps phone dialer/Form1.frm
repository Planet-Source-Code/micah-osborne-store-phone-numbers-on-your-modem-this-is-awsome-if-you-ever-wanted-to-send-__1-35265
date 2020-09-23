VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command17 
      Height          =   375
      Left            =   3720
      TabIndex        =   24
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command16 
      Height          =   375
      Left            =   3720
      TabIndex        =   23
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command14 
      Height          =   375
      Left            =   3720
      TabIndex        =   22
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      Height          =   375
      Left            =   2160
      TabIndex        =   21
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Caption         =   "&Store Number"
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Get Data"
      Height          =   375
      Left            =   2160
      TabIndex        =   19
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Read S Register"
      Height          =   375
      Left            =   2160
      TabIndex        =   18
      Top             =   1560
      Width           =   1455
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   510
      Left            =   4440
      TabIndex        =   17
      Top             =   1320
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   900
      _Version        =   327682
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   7560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Pick Up"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Hang Up"
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Dial"
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   12000
      TabIndex        =   11
      Top             =   0
      Width           =   12000
      Begin VB.CommandButton Command10 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   6240
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Speed Dial"
         Height          =   375
         Left            =   5040
         TabIndex        =   14
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Input Status :"
         Height          =   255
         Left            =   6480
         TabIndex        =   15
         Top             =   120
         Width           =   975
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   9840
         X2              =   12000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number :"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         X1              =   6000
         X2              =   9840
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FFFF&
         X1              =   0
         X2              =   6000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Synaps Phone Dialer"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.Timer Timer1 
      Left            =   5280
      Top             =   4440
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5640
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command5 
      Caption         =   "X"
      Height          =   255
      Left            =   11400
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Trudy &Van"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Will Jones"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Micah Osborne"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UP Stairs"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed Dial :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dial "4235592649"
End Sub

Private Sub Command10_Click()
End
End Sub

Private Sub Command11_Click()
If MSComm1.PortOpen = False Then
MSComm1.PortOpen = True
End If
MSComm1.Output = "AT " & "&Z2=" & Text2.Text & vbCrLf
End Sub

Private Sub Command12_Click()
If MSComm1.PortOpen = False Then
MSComm1.PortOpen = True
End If
MSComm1.Output = "AT XN" & vbCrLf
End Sub

Private Sub Command13_Click()
If MSComm1.PortOpen = False Then
MSComm1.PortOpen = True
End If
MSComm1.Output = "AT " & "&V" & vbCrLf
End Sub

Private Sub Command14_Click()

If MSComm1.PortOpen = False Then
MSComm1.PortOpen = True
End If
MSComm1.Output = "AT " & "IN I0" & vbCrLf
End Sub

Private Sub Command15_Click()
If MSComm1.PortOpen = False Then
MSComm1.PortOpen = True
End If
MSComm1.Output = "AT " & "Sn?" & vbCrLf
End Sub

Private Sub Command3_Click()

Dial "14234724195"
End Sub

Private Sub Command4_Click()
Dial "4521314"
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
Dial Text2.Text
End Sub

Private Sub Command7_Click()
On Error Resume Next
If MSComm1.PortOpen = True Then
MSComm1.PortOpen = False
AddStatus "Port Closed"
End If
End Sub

Private Sub Dial(Number As String)
If MSComm1.PortOpen = True Then
MSComm1.PortOpen = False
End If
MSComm1.PortOpen = True

MSComm1.Output = "ATDT " & Number & vbCrLf

End Sub

Private Sub AddStatus(Text As String)
Text1.Text = Text1.Text & Text & vbCrLf

End Sub

Private Sub Command8_Click()
Dim micah As String
MSComm1.PortOpen = True
MSComm1.Output = "atdt" & vbCrLf
MSComm1.InputMode = comInputModeText
MSComm1.NullDiscard = False
MSComm1.ParityReplace = "Micah Osborne"
MSComm1.Tag = "micah"

micah = MSComm1.Input
Text1.Text = micah
Timer1_Timer
End Sub

Private Sub Command9_Click()
If Form1.Height = 1095 Then
Form1.Height = 3480
Exit Sub
End If
If Form1.Height = 3480 Then
Form1.Height = 1095
Exit Sub
End If

End Sub

Private Sub Form_Load()

Timer1_Timer
End Sub

Private Sub Slider1_Click()
If MSComm1.PortOpen = False Then
MSComm1.PortOpen = True
End If
MSComm1.Output = "AT " & "Ln" & vbCrLf
Timer1_Timer
End Sub

Private Sub Timer1_Timer()
Dim micah As String
Timer1.Interval = Timer1.Interval + 1
If MSComm1.PortOpen = True Then
micah = MSComm1.Input

If micah = "BUSY" Then
 MsgBox "The Phone line is busy."
 MSComm1.PortOpen = False
End If
If micah = "NO DIALTONE" Then
MsgBox "NO dial tone"
End If
Text1.Text = Text1.Text & micah
End If
End Sub

Private Sub Timer22_Timer()
Dim micah As String
Timer1.Interval = Timer1.Interval + 1

micah = MSComm1.Input


 MSComm1.PortOpen = False


Text1.Text = Text1.Text & micah
End If
End Sub
