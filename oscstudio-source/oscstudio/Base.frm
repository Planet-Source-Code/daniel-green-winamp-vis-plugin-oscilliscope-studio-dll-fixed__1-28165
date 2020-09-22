VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Base 
   BackColor       =   &H80000004&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "O s c i l l i s c o p e  * S t u d i o *"
   ClientHeight    =   7335
   ClientLeft      =   1905
   ClientTop       =   2955
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmOsc 
      Caption         =   "Oscilliscope Options"
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   4455
      Begin VB.Frame frmPerFrame 
         Caption         =   "Per Frame"
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4215
         Begin VB.TextBox txtBufferClear 
            Height          =   285
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   17
            Text            =   "0"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtFrameY 
            Height          =   285
            Left            =   2400
            TabIndex        =   6
            Text            =   "68.5"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtFrameX 
            Height          =   285
            Left            =   360
            TabIndex        =   4
            Text            =   "0"
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Frames until the buffer is clear="
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label lblY 
            Caption         =   "Y="
            Height          =   255
            Left            =   2160
            TabIndex        =   7
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblX 
            Caption         =   "X="
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame frmPerPixel 
         Caption         =   "Per Pixel"
         Height          =   1335
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   4215
         Begin VB.TextBox txtBlue 
            Height          =   285
            Left            =   2760
            MaxLength       =   3
            TabIndex        =   15
            Text            =   "255"
            Top             =   840
            Width           =   1275
         End
         Begin VB.TextBox txtGreen 
            Height          =   285
            Left            =   1440
            MaxLength       =   3
            TabIndex        =   14
            Text            =   "0"
            Top             =   840
            Width           =   1275
         End
         Begin VB.TextBox txtRed 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   13
            Text            =   "255"
            Top             =   840
            Width           =   1275
         End
         Begin VB.TextBox txtPixelX 
            Height          =   285
            Left            =   360
            TabIndex        =   10
            Text            =   "0"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtPixelY 
            Height          =   285
            Left            =   2400
            TabIndex        =   9
            Text            =   "0"
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblBlue 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Blue="
            Height          =   255
            Left            =   2760
            TabIndex        =   21
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblGreen 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Green="
            Height          =   255
            Left            =   1440
            TabIndex        =   20
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblRed 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Red="
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblPixelX 
            Caption         =   "X="
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblPixelY 
            Caption         =   "Y="
            Height          =   255
            Left            =   2160
            TabIndex        =   11
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame frmDesc 
         Caption         =   "Variables && Description of Functions"
         Height          =   2055
         Left            =   120
         TabIndex        =   22
         Top             =   2520
         Width           =   4215
         Begin VB.CommandButton cmdSupportedFunctions 
            Caption         =   "Function Help"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CommandButton cmdRemoveConst 
            Caption         =   "Remove Constant"
            Height          =   375
            Left            =   2280
            TabIndex        =   30
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CommandButton cmdAddConst 
            Caption         =   "Add Constant"
            Height          =   375
            Left            =   2280
            TabIndex        =   29
            Top             =   840
            Width           =   1815
         End
         Begin VB.ListBox lstConstants 
            BeginProperty Font 
               Name            =   "Lucida Sans Unicode"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   120
            TabIndex        =   28
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox txtConstantValue 
            Height          =   285
            Left            =   2280
            TabIndex        =   25
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtConstantName 
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label lblSpacer 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "<-->"
            Height          =   255
            Left            =   1920
            TabIndex        =   27
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Constant Value="
            Height          =   255
            Left            =   2280
            TabIndex        =   26
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblConstantName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Constant Name="
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1815
         End
      End
   End
   Begin VB.Timer tmrScope 
      Interval        =   500
      Left            =   2280
      Top             =   1800
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   7080
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4101
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4101
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   2280
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.osc"
      Filter          =   "*.osc"
   End
   Begin VB.PictureBox ScopeBuf 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   293
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.PictureBox ScopePic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2115
      ScaleWidth      =   4395
      TabIndex        =   18
      Top             =   120
      Width           =   4455
   End
   Begin VB.Menu mnuPresets 
      Caption         =   "&Presets"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuPresetsLoad 
         Caption         =   "&Load"
      End
      Begin VB.Menu mnuPresetsSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuPresetsNew 
         Caption         =   "&New"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Base"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'you must must must give me credit, or i'll kill you.
'-dan green
'
Option Explicit
' Keep track of user-defined constants
Dim ConstNames As New Collection
Dim thebeat As Double
Private arrColor() As Byte
Dim Parser As New clsParse
Dim fps, fade, numfade, frames, flamehead, pixelx As String, pixely As String, framex As String, framey As String, r, g, b, totalfps, peaks, rr, gg, bb, BeatAverage, a
Private MasterWindow As Long
Private WaveData() As Byte
Private Const WM_USER = &H400
Private Const WM_LBUTTONUP = &H202
Const BLACKNESS = &H42          ' (DWORD) dest = BLACK
Const DSTINVERT = &H550009      ' (DWORD) dest = (NOT dest)
Const MERGECOPY = &HC000CA      ' (DWORD) dest = (source AND pattern)
Const MERGEPAINT = &HBB0226     ' (DWORD) dest = (NOT source) OR dest
Const NOTSRCCOPY = &H330008     ' (DWORD) dest = (NOT source)
Const NOTSRCERASE = &H1100A6    ' (DWORD) dest = (NOT src) AND (NOT dest)
Const PATCOPY = &HF00021        ' (DWORD) dest = pattern
Const PATINVERT = &H5A0049      ' (DWORD) dest = pattern XOR dest
Const PATPAINT = &HFB0A09       ' (DWORD) dest = DPSnoo
Const SRCAND = &H8800C6         ' (DWORD) dest = source AND dest
Const SRCCOPY = &HCC0020        ' (DWORD) dest = source
Const SRCERASE = &H440328       ' (DWORD) dest = source AND (NOT dest )
Const SRCINVERT = &H660046      ' (DWORD) dest = source XOR dest
Const SRCPAINT = &HEE0086       ' (DWORD) dest = source OR dest
Const WHITENESS = &HFF0062      ' (DWORD) dest = WHITE
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SendMessageM Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Sub cmdLaunch_Click()

End Sub

Private Sub cmdAddConst_Click()
On Error GoTo cmdAddConst_ErrHandler
    
    ' Validity checks
    If Trim(txtConstantName) = "" Then
        MsgBox "A valid constant name must begin with a letter. " & _
            "Additional letters may include the alphanumeric letters or underscores. " & _
            "For Example: MyConst_2", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(txtConstantValue) Then
        MsgBox "Please enter a valid number!", vbExclamation
        Exit Sub
    End If
    
    Parser.AddConstant txtConstantName, CDbl(txtConstantValue)
    
    lstConstants.AddItem txtConstantName & " - " & txtConstantValue
    ConstNames.Add txtConstantName.text
    txtConstantName = ""
    txtConstantValue = ""
    Exit Sub

cmdAddConst_ErrHandler:
    ' If the AddConstant call raises an error, it's
    ' description will be shown to the user
    MsgBox "Error:" & vbCrLf & Err.Description, _
            vbCritical
End Sub

Private Sub cmdRemoveConst_Click()
On Error GoTo cmdRemoveConst_ErrHandler

    If lstConstants.ListIndex = -1 Then
        MsgBox "Select a constant to remove!", vbExclamation
        Exit Sub
    End If

    Parser.RemoveConstant ConstNames(lstConstants.ListIndex + 1)
    
    ConstNames.Remove lstConstants.ListIndex + 1
    lstConstants.RemoveItem lstConstants.ListIndex
    Exit Sub
    
cmdRemoveConst_ErrHandler:
    MsgBox "Error:" & vbCrLf & Err.Description, _
            vbCritical
End Sub

Private Sub cmdSupportedFunctions_Click()
frmFunctions.Show
End Sub

Private Sub Form_Load()
On Error Resume Next
    If Val(Command$) = 0 Then
        MsgBox "Oscilliscope *Studio* by Dan Green (dan.green@morphedmedia.com)" + vbCrLf + _
            "http://www.morphedmedia.com/winamp" + vbCrLf + _
            "Written in Visual Basic using VisRelay by Murphy McCauley.", vbInformation, "About..."
        Unload Me
        Exit Sub
    End If
    
    'Check to see if I'm already running, and bomb out if I am...
    If App.PrevInstance Then
        MsgBox "Sorry, you can only run one instance of" + vbCrLf + _
        "this plugin at once.  Please close all of its" + vbCrLf + _
        "other windows.", vbExclamation, "Sorry!"
        Unload Me
        Exit Sub
    End If
    
    'VisRelay's control window handle is sent on the commandline...
    MasterWindow = Val(Command$)
    
    'Put my window permanently on top...
    'SetOnTop Me
    
    'Set up the waveform data input...
    If CreateArray(WaveformArray, WaveData) = False Then
        MsgBox "Wave data couldn't read!", vbCritical, "Ack!"
        Unload Me
        Exit Sub
    End If
    
    'Set up what info we want from WinAMP and stuff...
    Dim TheModule As WinAMPVisModule
    GetModule TheModule
    TheModule.WaveformNch = 1
    TheModule.SpectrumNch = 0
    TheModule.DelayMs = 10
    TheModule.LatencyMs = 2
    SetModule TheModule
    fps = 0
    fade = 0
    numfade = 0
    frames = 0
    flamehead = 0
    pixely = txtPixelY
    pixelx = txtPixelX
    framex = txtFrameX
    framey = txtFrameY
    totalfps = 0
    thebeat = 0
    r = txtRed
    txtConstantName = "beat"
    txtConstantValue = thebeat
    cmdAddConst_Click
    txtConstantName = "frameX"
    txtConstantValue = framex
    cmdAddConst_Click
    txtConstantName = "frameY"
    txtConstantValue = framey
    cmdAddConst_Click
    txtConstantName = "pixelX"
    txtConstantValue = pixelx
    cmdAddConst_Click
    txtConstantName = "pixelY"
    txtConstantValue = pixely
    cmdAddConst_Click
    lblRed.ForeColor = RGB(txtRed, 0, 0)
    g = txtGreen
    lblGreen.ForeColor = RGB(0, txtGreen, 0)
    b = txtBlue
    lblBlue.ForeColor = RGB(0, 0, txtBlue)
    Open "c:\oscstudio_debug.txt" For Output As #1
    Print #1, "Oscilliscope *Studio* started: " & Now
    Close #1
    arrColor() = Array(vbWhite, vbBlack, vbBlue, vbRed)
    'Tell VisRelay to simulate the left mouse button up event on
    'the ScopeBuf window (which will never get mouse events otherwise
    'since it's invisible)...  This is a dumb (but easy) way to do
    'this.  The proper way would be to subclass the window (either
    'by hand or using an ActiveX control like Deeth MessagePeek or
    'Antenna or that Desaware tool...
    RegisterNotifyWindowEx ScopeBuf.hwnd, WM_LBUTTONUP

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    'Turn off the notifications...
    RegisterNotifyWindow 0
    
    'Free the array...
    FreeArray WaveData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
    If MasterWindow <> 0 Then
        'Tell the plugin to quit.
        SendMessageM MasterWindow, WM_USER, vrTerminate, 0
    End If
    Open "c:\oscstudio_debug.txt" For Append As #1
    Print #1, "Oscilliscope *Studio* ended: " & Now
    Print #1, "With a Total FPS of: " & totalfps / flamehead
    Close #1
    Unload frmAbout
    Unload frmFunctions
    Unload Me
    
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show
'MsgBox "Oscilliscope *Studio* by Dan Green (dan.green@morphedmedia.com)" + vbCrLf + _
'            "http://www.morphedmedia.com/winamp" + vbCrLf + _
'            "Written in, oh my hell, Visual Basic." + vbCrLf + _
'            "a 750mhz amd duron got avg. fps of 42.7 :)" + vbCrLf + _
'            "not that fast, but gets the job done." + vbCrLf + _
'            "check c:\oscstudio_debug.txt for debug info." + vbCrLf + _
'            "Greetz to Murph McCauley for VisRelay." + vbCrLf + _
'            "Coming soon: evaluating expressions from text-boxes.", vbInformation, "About..."
End Sub

Private Sub mnuPresetsLoad_Click()
Dialog.DialogTitle = "Select Preset to Load..."
Dialog.InitDir = App.path & "\OscStudio"
Dialog.ShowOpen
Dim i
i = ReadText(Dialog.Filename)
txtFrameX = Word(i, 1)
txtFrameY = Word(i, 2)
txtBufferClear = Word(i, 3)
txtPixelX = Word(i, 4)
txtPixelY = Word(i, 5)
txtRed = Word(i, 6)
txtGreen = Word(i, 7)
txtBlue = Word(i, 8)
End Sub

Private Sub mnuPresetsSave_Click()
On Error Resume Next
Dim i
i = InputBox("Enter the title of this preset: ", "Save Preset...")
Open App.path & "\OscStudio\" & i & ".osc" For Append As #1
Print #1, txtFrameX & " " & txtFrameY & " " & txtBufferClear & " " & txtPixelX & " " & txtPixelY & " " & txtRed & " " & txtGreen & " " & txtBlue
Close #1
End Sub

Private Sub ScopeBuf_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    'This event gets fired when VisRelay wants to communicate.
    'With the PictureBox's scale mode set to pixels, X is the low word
    'of LParam and Y is the high word.  Neat.
    
    Select Case x
        Case vrDataUpdate
            Call Render
        Case vrTerminate
            Unload Me
    End Select
End Sub

Private Sub Render()
On Error Resume Next
    'Draw the waveform...
    
    Dim x As Single, pos
    peaks = 0
    framex = Parser.ParseExpression(framex)
    framey = Parser.ParseExpression(framey)
    pixelx = Parser.ParseExpression(pixelx)
    pixely = Parser.ParseExpression(pixely)
    Parser.mConstants.Remove "frameX"
    Parser.AddConstant "frameX", CDbl(framex)
    Parser.mConstants.Remove "framey"
    Parser.AddConstant "frameY", CDbl(framey)
    Parser.mConstants.Remove "pixelX"
    Parser.AddConstant "pixelX", CDbl(pixelx)
    Parser.mConstants.Remove "pixelY"
    Parser.AddConstant "pixelY", CDbl(pixely)
    Status.Panels(2).text = Parser.mConstants.Item("frameX") & " " & framey & " " & pixelx & " " & pixely
    'Using ScopeBuf the way I am is a simple way to do double-buffering
    'in VB.  It's not *amazingly* fast, but it's so easy it hurts.
    If fade >= txtBufferClear Then
    ScopeBuf.Cls
    fade = 0
    numfade = numfade + 1
    End If
    'ScopeBuf.Cls
    '(WaveData(0) Xor 128) \ 2
    ScopeBuf.CurrentX = framex
    ScopeBuf.CurrentY = framey
    BeatAverage = 0
    For x = 0 To 288
        ScopeBuf.Line Step(pixelx, pixely)-(x, (WaveData(x) Xor 128) / 2), RGB(Abs(r - (WaveData(x))), Abs(g - (WaveData(x))), Abs(b - (WaveData(x))))
        'peaks = peaks + WaveData(X)
    Next
    For a = 1 To 20
        BeatAverage = BeatAverage + WaveData(a)
    Next
    BeatAverage = BeatAverage / 20
    If BeatAverage <= 100 Then
    thebeat = 1
    'RunBLT
    'Launch
    End If
    If BeatAverage >= 101 Then
    thebeat = 0
    End If
    Parser.mConstants.Remove "beat"
    Parser.AddConstant "beat", thebeat
    Parser.EditConstant thebeat, "beat"
    EditConst thebeat, "beat"
    Open "c:\oscstudio_debug.txt" For Append As #1
    Write #1, "ba " & BeatAverage
    Write #1, "const " & Parser.mConstants.Item("beat")
    Close #1
    ScopePic.Picture = ScopeBuf.Image
    fps = fps + 1
    fade = fade + 1
    frames = frames + 1
End Sub

Public Function GetRandomInteger(LowerBound, UpperBound) As Long
On Error Resume Next
GetRandomInteger = Int((UpperBound - LowerBound + 1) * (Rnd + LowerBound))
End Function

Private Sub tmrScope_Timer()
On Error Resume Next
'flamehead = flamehead + 1
'Dim i
Status.Panels(1).text = "FPS: " & (fps * 2) & " BEAT: " & thebeat
'Status.Panels(3).text = "avg. peak: " & Left((peaks / 288), 5)
'totalfps = totalfps + (fps * 2)
'i = InStr(totalfps / flamehead, ".")
'Status.Panels(2).text = "flamehead: " & Left(totalfps / flamehead, i + 1)
    fps = 0
End Sub

Function GetColor(cur, curr)
On Error Resume Next
GetColor = Abs(curr - cur)
End Function

Private Sub txtBlue_Change()
On Error Resume Next
lblBlue.ForeColor = RGB(0, 0, txtBlue)
b = txtBlue
End Sub

Private Sub txtFrameX_Change()
On Error Resume Next
If txtFrameX = "" Then txtFrameX = "0"
framex = txtFrameX
End Sub

Private Sub txtFrameY_Change()
On Error Resume Next
If txtFrameY = "" Then txtFrameY = "0"
framey = txtFrameY
End Sub

Private Sub txtGreen_Change()
On Error Resume Next
lblGreen.ForeColor = RGB(0, txtGreen, 0)
g = txtGreen
End Sub

Private Sub txtPixelX_Change()
On Error Resume Next
If txtPixelX = "" Then txtPixelX = "0"
pixelx = txtPixelX
End Sub

Private Sub txtPixelY_Change()
On Error Resume Next
If txtPixelY = "" Then txtPixelY = "0"
pixely = txtPixelY
End Sub

Private Sub txtRed_Change()
On Error Resume Next
lblRed.ForeColor = RGB(txtRed, 0, 0)
r = txtRed
End Sub

Private Sub RunBLT()
Dim ret
'If ScopePic.BackColor = vbBlack Then
'ScopePic.BackColor = vbWhite
'ScopeBuf.BackColor = vbWhite
'Exit Sub
'End If
'ScopePic.BackColor = vbBlack
'ScopeBuf.BackColor = vbBlack
ret = BitBlt(ScopeBuf.hDC, 0, 0, ScopeBuf.ScaleWidth, ScopeBuf.ScaleHeight, ScopeBuf.hDC, 0, 0, vbDstInvert)
End Sub

Sub EditConst(value, name)
Parser.mConstants.Item(name) = value
End Sub
