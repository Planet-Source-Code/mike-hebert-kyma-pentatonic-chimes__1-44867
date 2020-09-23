Attribute VB_Name = "Module1"
Option Explicit

'Declarations for DirectSound

Public dx As New DirectX7
Public ds As DirectSound
Public dsBuffer(5) As DirectSoundBuffer
Public dsCursor(5) As DSCURSORS
Public dsCaps As DSBCAPS
Public waveFormat As WAVEFORMATEX

'Program variables

Public scaleName As String
Public FileName As String
Public chimeStyle As String
Public presetName As String
Public presetDesc As String
Public i As Integer
Public stopflag As Boolean
Public note As Integer
Public length As Long
Public pan As Long
Public delay As Long
Public ratio1, ratio2, ratio3 As Single
Public ratio4, ratio5, ratio6 As Single
Public rNum1, rNum2, rNum3, rNum4, rNum5, rNum6 As String
Public rDiv1, rDiv2, rDiv3, rDiv4, rDiv5, rDiv6 As String
Public speed As String
Public gust As Integer
Public pitch As String
Public volume As String
Public Cancel As Integer

'Variables for cents to ratio conversion

Public cents As Double
Public dec As Double
Public fNum As Integer
Public fDiv As Integer
Public df As Double

'Function for pause between notes

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub Load_Buffers()

    'Set the DirectSound buffer flags
    'and create the buffers
    
    Dim bufferDesc As DSBUFFERDESC
    bufferDesc.lFlags = DSBCAPS_STATIC Or _
        DSBCAPS_CTRLFREQUENCY Or _
        DSBCAPS_CTRLPAN Or _
        DSBCAPS_CTRLVOLUME Or _
        DSBCAPS_GLOBALFOCUS
    For i = 0 To 5
        Set dsBuffer(i) = ds.CreateSoundBufferFromFile(App.Path & "\Chimes\" & chimeStyle, bufferDesc, waveFormat)
    Next i
    
    'Get the length of the buffer
    
    dsBuffer(0).GetCaps dsCaps
    length = dsCaps.lBufferBytes
    
    'Set the default playback rates(frequencies)
    
    SetFreq
    
End Sub

Public Sub SetFreq()

    dsBuffer(0).SetFrequency CLng(44100 * (ratio1 / pitch))
    dsBuffer(1).SetFrequency CLng(44100 * (ratio2 / pitch))
    dsBuffer(2).SetFrequency CLng(44100 * (ratio3 / pitch))
    dsBuffer(3).SetFrequency CLng(44100 * (ratio4 / pitch))
    dsBuffer(4).SetFrequency CLng(44100 * (ratio5 / pitch))
    dsBuffer(5).SetFrequency CLng(44100 * (ratio6 / pitch))

End Sub

Public Sub LoadScale()
    
    'Load the scale file
    
    Open scaleName For Input As #1
        
        Input #1, scaleName
        Input #1, rNum1
        Input #1, rDiv1
        Input #1, rNum2
        Input #1, rDiv2
        Input #1, rNum3
        Input #1, rDiv3
        Input #1, rNum4
        Input #1, rDiv4
        Input #1, rNum5
        Input #1, rDiv5
        Input #1, rNum6
        Input #1, rDiv6
    
    Close #1
    
    'Set the ratios
    
    ratio1 = Val(rNum1) / Val(rDiv1)
    ratio2 = Val(rNum2) / Val(rDiv2)
    ratio3 = Val(rNum3) / Val(rDiv3)
    ratio4 = Val(rNum4) / Val(rDiv4)
    ratio5 = Val(rNum5) / Val(rDiv5)
    ratio6 = Val(rNum6) / Val(rDiv6)
    
End Sub
