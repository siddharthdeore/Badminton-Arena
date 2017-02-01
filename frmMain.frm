VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   Caption         =   "SRM Team Robocon"
   ClientHeight    =   13335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   26820
   LinkTopic       =   "Form1"
   ScaleHeight     =   14000
   ScaleMode       =   0  'User
   ScaleWidth      =   26102.19
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdResetEncoder 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "Reset Encoder"
      Height          =   735
      Left            =   12720
      TabIndex        =   12
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Comunication"
      ForeColor       =   &H8000000B&
      Height          =   3135
      Left            =   12960
      TabIndex        =   8
      Top             =   9960
      Width           =   2295
      Begin VB.ComboBox cmbBaud 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "frmMain.frx":0000
         Left            =   120
         List            =   "frmMain.frx":0013
         TabIndex        =   11
         Text            =   "115200"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox cmbComPort 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "frmMain.frx":003A
         Left            =   120
         List            =   "frmMain.frx":003C
         TabIndex        =   10
         Text            =   "Select Com port"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton cmdOpenCom 
         Caption         =   "OPEN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   9
         Top             =   2400
         Width           =   1935
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   240
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         InputMode       =   1
      End
   End
   Begin VB.TextBox txtBotENC 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   3
      Left            =   16320
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtBotENC 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   15120
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtBotENC 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   13920
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtBotENC 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   12720
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtTerminal 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1635
      Left            =   12720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Timer comTimer 
      Interval        =   10
      Left            =   16800
      Top             =   0
   End
   Begin VB.TextBox txtXY 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   13560
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtXY 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   12720
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   12855
      Left            =   240
      ScaleHeight     =   13939.73
      ScaleMode       =   0  'User
      ScaleWidth      =   12656.16
      TabIndex        =   0
      Top             =   240
      Width           =   12375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000C000&
      Height          =   1095
      Left            =   12840
      TabIndex        =   15
      Top             =   8880
      Width           =   6495
   End
   Begin VB.Label lblGSdata 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   84
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   12015
      Left            =   19200
      TabIndex        =   14
      Top             =   480
      Width           =   5775
   End
   Begin VB.Label lblError 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   12960
      TabIndex        =   13
      Top             =   9000
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

' program logic control
Private bLocalEcho              As Boolean
Private bMessageMode            As Boolean

' constants for setting the LED images,
' used as an index for imgLED()
Private Const RedOff            As Long = 0
Private Const RedOn             As Long = 1
Private Const GreenOff          As Long = 2
Private Const GreenOn           As Long = 3


' the sendmessage API is used to write
' to the textbox to reduce flicker, this
' not required for serial communications.

' Win32 API constants
Private Const EM_GETSEL         As Long = &HB0
Private Const EM_SETSEL         As Long = &HB1
Private Const EM_GETLINECOUNT   As Long = &HBA
Private Const EM_LINEINDEX      As Long = &HBB
Private Const EM_LINELENGTH     As Long = &HC1
Private Const EM_LINEFROMCHAR   As Long = &HC9
Private Const EM_SCROLLCARET    As Long = &HB7
Private Const WM_SETREDRAW      As Long = &HB
Private Const WM_GETTEXTLENGTH  As Long = &HE

' Win32 API declarations
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long




Public pmmX, pmmY, mmpX, mmPY, encX, encY, LmX, LmY, rCount, RPM As Double

Private Sub cmdOK_Click()
Form1.Caption = "Hello"
End Sub

Private Sub cmdOpenCom_Click()
PortNum = Val(Mid$(cmbComPort.Text, 4))
If cmdOpenCom.Caption = "OPEN" Then
    If PortNum > 0 Then
        MSComm1.CommPort = PortNum
        MSComm1.Settings = cmbBaud.Text + ",n,8,1"        ' baud, parity, data bits, stop bits
        MSComm1.PortOpen = True
        cmdOpenCom.Caption = "CLOSE"
    End If
Else
'    If cmdBlinkTest.Caption = "Blink turn ON" Then
        MSComm1.PortOpen = False
        cmdOpenCom.Caption = "OPEN"
   ' End If
End If

End Sub

Private Sub cmdResetEncoder_Click()
On Error Resume Next
MSComm1.Output = "^"
End Sub

Private Sub comTimer_Timer()
pic.Cls
plot
pic.Circle (((((encX * mmpX * 10)) - 6100 * mmpX) * -1), ((((encY * mmPY * 10)) - 6700 * mmPY) * -1)), 150 * mmpX, ColorConstants.vbGreen

lblGSdata.Caption = Trim(RPM) + vbCrLf + vbCrLf + Str(rCount)
'Label1.Caption = Str(Rnd * (1000)) + vbCrLf + Str(Rnd * (1000)) + Str(Rnd * (1000)) + Str(Rnd * (1000)) + vbCrLf + Str(Rnd * (1000)) + Str(Rnd * (1000)) + Str(Rnd * (1000))
'On Error Resume Next
'If Len(MSComm1.Input) > 0 Then
'str1 = MSComm1.Input
'If InStr(str1, "<") > 0 Then
'txtData.Text = str1
'End If
'End If
' MSComm1.Input
End Sub

Private Sub Form_Activate()
plot
    
    ' setup the default comm port settings
    MSComm1.RThreshold = 1                  ' use 'on comm' event processing
    MSComm1.Settings = "115200,n,8,1"         ' baud, parity, data bits, stop bits
    MSComm1.SThreshold = 1                  ' allows us to track Tx LED
    MSComm1.InputMode = comInputModeBinary  ' binary mode, you can also use
                                            ' comInputModeText for text only use

' find available com port
 For i = 1 To 50
 On Error Resume Next
 MSComm1.CommPort = i
 On Error Resume Next
 MSComm1.PortOpen = True
 On Error Resume Next
 MSComm1.PortOpen = False
 ' add to list if no error
 If Err.Number = 0 Then
    cmbComPort.AddItem ("COM" + Str(i))
 End If
 Next i
 'MSComm1.PortOpen = True

PortNum = Val(Mid$(cmbComPort.List(0), 4))
If cmdOpenCom.Caption = "OPEN" Then
    If PortNum > 0 Then
        MSComm1.CommPort = PortNum
        MSComm1.PortOpen = True
        cmdOpenCom.Caption = "CLOSE"
        cmbComPort.Text = "COM" + Str(PortNum)
    End If
Else
'    If cmdBlinkTest.Caption = "Blink turn ON" Then
        MSComm1.PortOpen = False
        cmdOpenCom.Caption = "OPEN"
   ' End If
End If


End Sub

Private Sub Form_Load()
pmmX = 6100 / pic.Width
pmmY = 6700 / pic.Height
mmpX = 1 / pmmX
mmPY = 1 / pmmY
End Sub


Private Sub MSComm1_OnComm()
   
'******************************************************************************
' Synopsis:     Handle incoming characters, 'On Comm' Event
'
' Description:  By setting MSComm1.RThreshold = 1, this event will fire for
'               each character that arrives in the comm controls input buffer.
'               Set MSComm1.RThreshold = 0 if you want to poll the control
'               yourself, either via a TImer or within program execution loop.
'
'               In most cases, OnComm Event processing shown here is the prefered
'               method of processing incoming characters.
'
'******************************************************************************

    
    Static sBuff    As String           ' buffer for holding incoming characters
    Const MTC       As String = vbCrLf  ' message terminator characters (ususally vbCrLf)
    Const LenMTC    As Long = 2         ' number of terminator characters, must match MTC
    Dim iPtr        As Long             ' pointer to terminatior character

    ' OnComm fires for multiple Events
    ' so get the Event ID & process
    Select Case MSComm1.CommEvent
        
        ' Received RThreshold # of chars, in our case 1.
        Case comEvReceive
        
            ' read all of the characters from the input buffer
            ' StrConv() is required when using MSComm in binary mode,
            ' if you set MSComm1.InputMode = comInputModeText, it's not required
            
            sBuff = sBuff & StrConv(MSComm1.Input, vbUnicode)
                 txtTerminal.Text = sBuff
                 ParseStr sBuff
                'PostTerminal sBuff
                sBuff = vbNullString
            
            
            ' flash the Rx LED
    End Select
End Sub


Private Sub pic_KeyDown(KeyCode As Integer, Shift As Integer)
Me.Caption = KeyCode + Shift
End Sub

Private Sub pic_KeyPress(KeyAscii As Integer)
Me.Caption = KeyAscii
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
On Error GoTo er
'MSComm1.Output = "$" + txtXY(0).Text + "," + txtXY(1).Text + "\r\n"
er:
lblError.Caption = Error
End If
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtXY(0).Text = Format$(((X * pmmX / 10) - 610) * -1, "#")
txtXY(1).Text = Format$(((Y * pmmY / 10) - 670) * -1, "#")
LmX = X
LmY = Y
'pic.Cls
'plot
'pic.Circle (x, y), 550 * mmpX, vbRed
'pic.Circle (encX * mmpX, encY * mmPY), 550 * mmpX, vbRed

End Sub

Public Sub plot()



pic.Line (0 * mmpX, 0 * mmPY)-Step(40 * mmpX, 6700 * mmPY), vbYellow, B 'Left V Bar
    
'pic.Line (40 * mmpX, 0)-(40 * mmpX, pic.Height), vbWhite
'pic.Line (0 * mmpX, 0)-(0 * mmpX, pic.Height), vbWhite

'pic.Line (pic.Width - 40 * mmpX, 0)-(pic.Width - 40 * mmpX, pic.Height), vbWhite
'pic.Line (pic.Width, 0)-(pic.Width, pic.Height), vbWhite


'pic.Line (460 * mmpX, 0)-(460 * mmpX, pic.Height), vbWhite
'pic.Line (500 * mmpX, 0)-(500 * mmpX, pic.Height), vbWhite

pic.Line (460 * mmpX, 0 * mmPY)-Step(40 * mmpX, 6700 * mmPY), vbWhite, B  'Left2 V Bar

'center line
pic.Line (3030 * mmpX, 2020 * mmPY)-Step(40 * mmpX, 4680 * mmPY), vbWhite, B

pic.Line (5600 * mmpX, 0 * mmPY)-Step(40 * mmpX, 6700 * mmPY), vbWhite, B 'rIGHT LINE 2

'pic.Line (6060 * mmpX, 0)-(6060 * mmpX, pic.Height), vbWhite
'pic.Line (6100 * mmpX, 0)-(6100 * mmpX, pic.Height), vbWhite
pic.Line (6060 * mmpX, 0 * mmPY)-Step(40 * mmpX, 6700 * mmPY), vbWhite, B 'rIGHT LIN

pic.Line (0, 0)-(6100 * mmpX, 0), vbWhite
pic.Line (0, 20 * mmPY)-(6100 * mmpX, 20 * mmPY), vbWhite

pic.Line (0, 1980 * mmPY)-(6100 * mmpX, 1980 * mmPY), vbWhite
pic.Line (0, 2020 * mmPY)-(6100 * mmpX, 2020 * mmPY), vbWhite

pic.Line (0, 5900 * mmPY)-(6100 * mmpX, 5900 * mmPY), vbWhite
pic.Line (0, 5940 * mmPY)-(6100 * mmpX, 5940 * mmPY), vbWhite

pic.Line (0, 6660 * mmPY)-(6100 * mmpX, 6660 * mmPY), vbWhite
pic.Line (0, 6700 * mmPY)-(6100 * mmpX, 6700 * mmPY), vbWhite

pic.Line (3600 * mmpX, 3925 * mmPY)-Step(2000 * mmpX, 1250 * mmPY), vbYellow, B 'Service Rectangle

'Y 5175
'Y2 3925


pic.Circle (LmX, LmY), 450 * mmpX, vbGreen ' plot mouse position
End Sub



Public Sub PostTerminal(ByVal sNewData As String)

    ' display incoming characters in the
    ' textbox 'terminal' window. API is
    ' used only to reduce flicker.
    
    Dim lPtr    As Long
    ' this is faster and has less flicker but requires use of the Win API
    With txtTerminal
        lPtr = SendMessage(.hwnd, EM_GETLINECOUNT, 0, ByVal 0&)
        If lPtr > 550 Then
            'LockWindowUpdate .hWnd
            Call SendMessage(.hwnd, WM_SETREDRAW, False, ByVal 0&)
            lPtr = SendMessage(.hwnd, EM_LINEINDEX, 100, ByVal 0&)
            .SelStart = 0
            .SelLength = IIf(lPtr > 0, lPtr, 1000)
            .SelText = vbNullString
            Call SendMessage(.hwnd, WM_SETREDRAW, True, ByVal 0&)
            ' LockWindowUpdate 0
        End If
        .SelStart = SendMessage(.hwnd, WM_GETTEXTLENGTH, True, ByVal 0&)
        .SelText = sNewData
        .SelStart = SendMessage(.hwnd, WM_GETTEXTLENGTH, True, ByVal 0&)
    End With

End Sub
'<52,67,255,87

Public Sub ParseStr(ByVal sData As String)
cnt = 0
  For i = 1 To Len(sData)
        If Mid$(sData, i, 1) = "," Then cnt = cnt + 1
  Next
If InStr(1, sData, "<") And cnt > 0 Then
sData = Mid$(sData, 2, Len(sData))
X = Split(sData, ",")
rCount = Val(X(0))
RPM = Val(X(1))
'txtBotENC(2) = X(0)
'txtBotENC(3) = X(1)
'encX = Val(X(0))
'encY = Val(X(1))
End If
End Sub

