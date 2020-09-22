Attribute VB_Name = "WinAMPPlugin"
' Murphy McCauley (MurphyMc@Concentric.NET)
' http://www.fullspectrum.com/deeth/

' This code module contains the declarations for working with
' the Deeth VisRelay plugin, which is part of the Deeth
' WinAMP Visualization Toolkit for VisualBasic.
' Release 1 (Sept 11, 1999)

Option Explicit

'The vis-info Get routines.  They're easy, but not as fast as
'the arrays...
Declare Sub GetWaveformData Lib "vis_oscstudio.dll" ( _
    WaveformData As VisSingleData)
Declare Sub GetSpectrumData Lib "vis_oscstudio.dll" ( _
    SpectrumData As VisSingleData)
Declare Sub GetBothData Lib "vis_oscstudio.dll" ( _
    BothData As VisDoubleData)
    
'These are for working with the mod structure.
Declare Sub GetModule Lib "vis_oscstudio.dll" ( _
    TheModule As WinAMPVisModule)
Declare Sub SetModule Lib "vis_oscstudio.dll" ( _
    TheModule As WinAMPVisModule)

'These are for making custom arrays for the visualization info...
'they're the fastest way to do it.
Declare Function CreateArray Lib "vis_oscstudio.dll" ( _
    ByVal WhichKind As Long, TheArray() As Byte) As Boolean
Declare Function FreeArray Lib "vis_oscstudio.dll" ( _
    TheArray() As Byte) As Boolean

'These are for setting up a window for VisRelay to send Window Messages
'to (like the all-important Data Update message).
Declare Sub RegisterNotifyWindow Lib "vis_oscstudio.dll" ( _
    ByVal TheHWnd As Long)
Declare Sub RegisterNotifyWindowEx Lib "vis_oscstudio.dll" ( _
    ByVal TheHWnd As Long, ByVal WindowMessageToSend As Long)


'I use these custom data types just to make life a little easier.
Type VisDoubleData
    Spectrum(0 To 575, 0 To 1) As Byte
    Waveform(0 To 575, 0 To 1) As Byte
End Type
Type VisSingleData
    TheData(0 To 575, 0 To 1) As Byte
End Type

'These are constants that get used with Window Messages...
Public Const vrTerminate = 1
Public Const vrDataUpdate = 2

'These are constants for CreateArray()...
Public Const WaveformArray = 1
Public Const SpectrumArray = 2


Type WinAMPVisModule 'Most comments here are verbatim from VIS mini-SDK.
    DescriptionP As Long    'Useless.  Do not alter.
    hWndParent As Long      'WinAMP Window
    hDLLInstance As Long    'hInstance of VisRelay DLL
    sRate As Long           'Sample rate of currently playing item
    nCh As Long             'Number of channels of currently playing item
    LatencyMs As Long       'Latency between getting a Data Update and
                            'when the data is actually shown to the user
    DelayMs As Long         'Delay between data updates

    
    SpectrumNch As Long     'Put the number of Spectrum Analysis channels
                            'you want here.
    WaveformNch As Long     'Put the number of Waveform channels you want
                            'here.
                            
    SpectrumData(0 To 575, 0 To 1) As Byte
        'Spectrum analysis data -- it's slow to fetch this way, though!
        
    WaveformData(0 To 575, 0 To 1) As Byte
        'Waveform data -- it's slow to fetch this way, though!
        
    ConfigFP As Long        'Useless.  Do not alter.
    InitFP As Long          'Useless.  Do not alter.
    RenderFP As Long        'Useless.  Do not alter.
    QuitFP As Long          'Useless.  Do not alter.

    UserDataP As Long       'User data.  Optional.
End Type

