VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChimes 
   Caption         =   "Kyma Pentatonic Chimes"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   Icon            =   "frmChimes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd6th 
      Caption         =   "&6th"
      Height          =   255
      Left            =   4080
      TabIndex        =   51
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmd5th 
      Caption         =   "&5th"
      Height          =   255
      Left            =   4080
      TabIndex        =   50
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmd4th 
      Caption         =   "&4th"
      Height          =   255
      Left            =   4080
      TabIndex        =   49
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmd3rd 
      Caption         =   "&3rd"
      Height          =   255
      Left            =   4080
      TabIndex        =   48
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmd2nd 
      Caption         =   "&2nd "
      Height          =   255
      Left            =   4080
      TabIndex        =   47
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmd1st 
      Caption         =   "&1st"
      Height          =   255
      Left            =   4080
      TabIndex        =   46
      Top             =   2880
      Width           =   615
   End
   Begin RichTextLib.RichTextBox txtDescription 
      Height          =   975
      Left            =   240
      TabIndex        =   45
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1720
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmChimes.frx":0CCA
   End
   Begin VB.TextBox txtScale 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   44
      Text            =   "frmChimes.frx":0D93
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtChime 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   43
      Text            =   "frmChimes.frx":0DAB
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtPreset 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmChimes.frx":0DC3
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit Program"
      Height          =   375
      Left            =   4080
      TabIndex        =   42
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   5040
      TabIndex        =   41
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   4080
      TabIndex        =   40
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdDeletePreset 
      Caption         =   "Delete Preset"
      Height          =   375
      Left            =   240
      TabIndex        =   38
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdDeleteScale 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5040
      TabIndex        =   37
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdSaveScale 
      Caption         =   "Save"
      Height          =   375
      Left            =   4080
      TabIndex        =   25
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtRatio 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   35
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtCents 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   34
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox txtDiv6 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   32
      Text            =   "frmChimes.frx":0DDC
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtNum6 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   30
      Text            =   "frmChimes.frx":0DDE
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtDiv5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtNum5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtDiv4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtNum4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtDiv3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtNum3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtDiv2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtNum2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtDiv1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "frmChimes.frx":0DE0
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtNum1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "frmChimes.frx":0DE4
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton cmdSavePreset 
      Caption         =   "Save As Preset"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3960
      Width           =   855
   End
   Begin VB.HScrollBar scrVolume 
      Height          =   255
      LargeChange     =   100
      Left            =   240
      Max             =   0
      Min             =   3000
      TabIndex        =   9
      Top             =   3600
      Value           =   1500
      Width           =   1815
   End
   Begin VB.HScrollBar scrPitch 
      Height          =   255
      LargeChange     =   5
      Left            =   240
      Max             =   10
      Min             =   60
      TabIndex        =   7
      Top             =   3000
      Value           =   15
      Width           =   1815
   End
   Begin VB.HScrollBar scrSpeed 
      Height          =   255
      LargeChange     =   10
      Left            =   240
      Max             =   1
      Min             =   100
      TabIndex        =   5
      Top             =   2400
      Value           =   6
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Notes"
      Height          =   255
      Left            =   4080
      TabIndex        =   52
      Top             =   840
      Width           =   615
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   6000
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   6000
      X2              =   6000
      Y1              =   120
      Y2              =   5400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   6000
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   5400
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "Description"
      Height          =   255
      Left            =   240
      TabIndex        =   39
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label21 
      Caption         =   "="
      Height          =   255
      Left            =   4800
      TabIndex        =   36
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "Convert Cents to Ratio"
      Height          =   255
      Left            =   4080
      TabIndex        =   33
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "/"
      Height          =   255
      Left            =   5270
      TabIndex        =   31
      Top             =   2910
      Width           =   135
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "/"
      Height          =   255
      Left            =   5270
      TabIndex        =   29
      Top             =   2550
      Width           =   135
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "/"
      Height          =   255
      Left            =   5270
      TabIndex        =   28
      Top             =   2190
      Width           =   135
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "/"
      Height          =   255
      Left            =   5270
      TabIndex        =   27
      Top             =   1830
      Width           =   135
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "/"
      Height          =   255
      Left            =   5270
      TabIndex        =   26
      Top             =   1470
      Width           =   135
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "/"
      Height          =   255
      Left            =   5270
      TabIndex        =   16
      Top             =   1110
      Width           =   135
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Scale Ratios"
      Height          =   255
      Left            =   4800
      TabIndex        =   13
      Top             =   840
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   4455
      Left            =   2160
      Picture         =   "frmChimes.frx":0DE8
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Volume Differential"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3315
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Relative Pitch"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2715
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Wind Speed"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2115
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Chime"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Scale"
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Preset"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmChimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd1st_Click()
    
    If Not (dsBuffer(0) Is Nothing) Then
        dsBuffer(0).Stop
        dsBuffer(0).SetCurrentPosition 0
        dsBuffer(0).SetVolume 0
        dsBuffer(0).SetPan -7500
        dsBuffer(0).SetFrequency CLng(44100 * (Val(txtNum1.Text) / Val(txtDiv1.Text)) / pitch)
        dsBuffer(0).Play DSBPLAY_DEFAULT
    End If
    
End Sub

Private Sub cmd2nd_Click()

    If txtNum2.Text = "" Or txtDiv2.Text = "" Then
        Exit Sub
    End If

    If Not (dsBuffer(1) Is Nothing) Then
        dsBuffer(1).Stop
        dsBuffer(1).SetCurrentPosition 0
        dsBuffer(1).SetVolume 0
        dsBuffer(1).SetPan -5000
        dsBuffer(1).SetFrequency CLng(44100 * (Val(txtNum2.Text) / Val(txtDiv2.Text)) / pitch)
        dsBuffer(1).Play DSBPLAY_DEFAULT
    End If
    
End Sub

Private Sub cmd3rd_Click()
    
    If txtNum3.Text = "" Or txtDiv3.Text = "" Then
        Exit Sub
    End If
    
    If Not (dsBuffer(2) Is Nothing) Then
        dsBuffer(2).Stop
        dsBuffer(2).SetCurrentPosition 0
        dsBuffer(2).SetVolume 0
        dsBuffer(2).SetPan -2500
        dsBuffer(2).SetFrequency CLng(44100 * (Val(txtNum3.Text) / Val(txtDiv3.Text)) / pitch)
        dsBuffer(2).Play DSBPLAY_DEFAULT
    End If
    
End Sub

Private Sub cmd4th_Click()
    
    If txtNum4.Text = "" Or txtDiv4.Text = "" Then
        Exit Sub
    End If
    
    If Not (dsBuffer(3) Is Nothing) Then
        dsBuffer(3).Stop
        dsBuffer(3).SetCurrentPosition 0
        dsBuffer(3).SetVolume 0
        dsBuffer(3).SetPan 2500
        dsBuffer(3).SetFrequency CLng(44100 * (Val(txtNum4.Text) / Val(txtDiv4.Text)) / pitch)
        dsBuffer(3).Play DSBPLAY_DEFAULT
    End If
    
End Sub

Private Sub cmd5th_Click()
    
    If txtNum5.Text = "" Or txtDiv5.Text = "" Then
        Exit Sub
    End If
    
    If Not (dsBuffer(4) Is Nothing) Then
        dsBuffer(4).Stop
        dsBuffer(4).SetCurrentPosition 0
        dsBuffer(4).SetVolume 0
        dsBuffer(4).SetPan 5000
        dsBuffer(4).SetFrequency CLng(44100 * (Val(txtNum5.Text) / Val(txtDiv5.Text)) / pitch)
        dsBuffer(4).Play DSBPLAY_DEFAULT
    End If
    
End Sub

Private Sub cmd6th_Click()
    
    If Not (dsBuffer(5) Is Nothing) Then
        dsBuffer(5).Stop
        dsBuffer(5).SetCurrentPosition 0
        dsBuffer(5).SetVolume 0
        dsBuffer(5).SetPan 7500
        dsBuffer(5).SetFrequency CLng(44100 * (Val(txtNum6.Text) / Val(txtDiv6.Text)) / pitch)
        dsBuffer(5).Play DSBPLAY_DEFAULT
    End If
    
End Sub

Private Sub cmdAbout_Click()

    dlgAbout.Show vbModal
    
End Sub

Private Sub cmdDeletePreset_Click()
    
    If txtPreset.Text = "Click to Select a Preset" Then
        Exit Sub
    End If
    
    If MsgBox("Confirm delete?", vbYesNo) = vbYes Then
        Kill App.Path & "\Presets\" & txtPreset & ".txt"
    End If
    
End Sub

Private Sub cmdDeleteScale_Click()
    
    If txtScale.Text = "Click to Select a Scale" Then
        Exit Sub
    End If
    
    If MsgBox("Confirm delete?", vbYesNo) = vbYes Then
        Kill App.Path & "\Scales\" & txtScale & ".txt"
    End If
    
End Sub

Private Sub cmdExit_Click()
        
        'Call the Form_Unload event
        
        Unload Me
            
End Sub

Private Sub cmdHelp_Click()

    frmHelp.Show
    
End Sub

Private Sub cmdPlay_Click()

    'Exit if no scale or chime selected
    
    If scaleName = "" Then
        Exit Sub
    End If
    
    If chimeStyle = "" Then
        Exit Sub
    End If
    
    'Re-initialize the stopflag
    
    stopflag = False
    
    'Make certain the ratios are set
    
    ratio1 = rNum1 / rDiv1
    ratio2 = rNum2 / rDiv2
    ratio3 = rNum3 / rDiv3
    ratio4 = rNum4 / rDiv4
    ratio5 = rNum5 / rDiv5
    ratio6 = rNum6 / rDiv6
    
    'Initialize the pitch
    
    pitch = scrPitch.Value / 10
    
    'Set the buffer playback rates
    
    dsBuffer(0).SetFrequency CLng(44100 * (ratio1 / pitch))
    dsBuffer(1).SetFrequency CLng(44100 * (ratio2 / pitch))
    dsBuffer(2).SetFrequency CLng(44100 * (ratio3 / pitch))
    dsBuffer(3).SetFrequency CLng(44100 * (ratio4 / pitch))
    dsBuffer(4).SetFrequency CLng(44100 * (ratio5 / pitch))
    dsBuffer(5).SetFrequency CLng(44100 * (ratio6 / pitch))

    Do Until stopflag = True
    
        'Seed the random number generator
    
        Randomize

        'Make sure to capture stop events
        
        DoEvents
        
        'Pause between successive notes
        
        Pause
        
        'Sound a randomly selected note at
        'a random position and volume
        
        note = Int((Rnd * 5) + 0.5)
        pan = Int((Rnd * 20000) - 10000)
        volume = Int(Rnd * scrVolume.Value)
        
        'Make sure the selected buffer is not playing
        'If it is ready then play the note
        
            dsBuffer(note).GetCurrentPosition dsCursor(note)
            If dsCursor(note).lPlay = 0 Or dsCursor(note).lPlay = length Then
                dsBuffer(note).SetPan pan
                dsBuffer(note).SetVolume volume * -1
                dsBuffer(note).Play DSBPLAY_DEFAULT
            End If
        
        'Continue until a stop event is processed
        
    Loop
    
End Sub

Private Sub cmdSavePreset_Click()

    If txtChime.Text = "Click to Select a Chime" Or _
        txtScale.Text = "Click to Select a Scale" Then
        MsgBox ("You must first select a Chime and Scale")
        Exit Sub
    End If
        
    dlgPresetSave.Show vbModal
    
End Sub

Private Sub cmdSaveScale_Click()

    dlgScaleName.Show vbModal
    
End Sub

Private Sub cmdStop_Click()

    'Set the stopflag
    
    stopflag = True
    
    'Stop playing without clearing buffers
    
    For i = 0 To 5
        dsBuffer(i).Stop
    Next i
    
    frmChimes.SetFocus
    
End Sub

Private Sub ChimeFile(chimeStyle)
    
    If FileName = "" Then
        Exit Sub
    End If
    
    chimeStyle = App.Path & "\Chimes\" & chimeStyle & ".wav"
    
    txtChime.Text = Left(FileName, (Len(FileName) - 4))
    
    Load_Buffers
    
    cmdPlay_Click
    
End Sub

Private Sub PresetFile()

    txtDescription.Text = ""
    
    Open App.Path & "\Presets\" & FileName _
        For Input As #1
        
        Input #1, presetName
            txtPreset.Text = presetName
        Input #1, scaleName
            scaleName = scaleName & ".txt"
        Input #1, chimeStyle
            chimeStyle = chimeStyle
            txtChime.Text = Left(chimeStyle, (Len(chimeStyle) - 4))
        Input #1, presetDesc
            txtDescription.Text = presetDesc
        Input #1, speed
            speed = Val(speed)
        Input #1, pitch
            pitch = Val(pitch)
        Input #1, volume
            volume = Val(volume)
            
    Close #1
    
    Load_Buffers
    
    LoadScale
            
    scrSpeed.Value = Val(speed)
    scrPitch.Value = Val(pitch)
    scrVolume.Value = Val(volume)
    
    cmdPlay_Click
    
End Sub

Private Sub ScaleFile(scaleName)

    If FileName = "" Then
        Exit Sub
    End If
    
    scaleName = App.Path & "\Scales\" & scaleName
    
    txtScale = Left(FileName, (Len(FileName) - 4))
    
    LoadScale
    
    cmdPlay_Click
        
End Sub

Private Sub Form_Load()

    'Create the DirectSound object
    
    Set ds = dx.DirectSoundCreate("")
    If Err.Number <> 0 Then
        MsgBox "Unable to start DirectSound"
        End
    End If
    
    'Set the priority
    
    ds.SetCooperativeLevel Me.hWnd, DSSCL_EXCLUSIVE
       
    'Preset the pitch
    
    pitch = 1.5
    
    'Display the form
    
    'Me.Show
        
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Query quit if close button clicked
    
    If MsgBox("Quit Kyma Chimes? Are you sure?", vbYesNo + vbQuestion + vbApplicationModal, "Kyma Pentatonic Chimes") = vbYes Then
        
        'Exit program if "Yes"
                
        For i = 0 To 5
            If Not (dsBuffer(i) Is Nothing) Then
                dsBuffer(i).Stop
                Set dsBuffer(i) = Nothing
            End If
        Next i
        
        'Unload DirectSound
        
        Set ds = Nothing
        Set dx = Nothing
    
        'Dump the form
        
        Unload Me
        
        'Terminate the program
        
        End
        
    Else
    
        'Resume program if "No"
        
        Cancel = 1
        
    End If
    
End Sub

Private Sub scrPitch_Change()

    pitch = Val(scrPitch.Value) / 10
    
    dsBuffer(0).SetFrequency CLng(44100 * (ratio1 / pitch))
    dsBuffer(1).SetFrequency CLng(44100 * (ratio2 / pitch))
    dsBuffer(2).SetFrequency CLng(44100 * (ratio3 / pitch))
    dsBuffer(3).SetFrequency CLng(44100 * (ratio4 / pitch))
    dsBuffer(4).SetFrequency CLng(44100 * (ratio5 / pitch))
    dsBuffer(5).SetFrequency CLng(44100 * (ratio6 / pitch))
    
    cmdPlay.SetFocus
    
End Sub

Private Sub scrPitch_Scroll()

    pitch = Val(scrPitch.Value) / 10
    
    dsBuffer(0).SetFrequency CLng(44100 * (ratio1 / pitch))
    dsBuffer(1).SetFrequency CLng(44100 * (ratio2 / pitch))
    dsBuffer(2).SetFrequency CLng(44100 * (ratio3 / pitch))
    dsBuffer(3).SetFrequency CLng(44100 * (ratio4 / pitch))
    dsBuffer(4).SetFrequency CLng(44100 * (ratio5 / pitch))
    dsBuffer(5).SetFrequency CLng(44100 * (ratio6 / pitch))
    
    cmdPlay.SetFocus
    
End Sub

Private Sub scrSpeed_Change()

    speed = Val(scrSpeed.Value)
    
    cmdPlay.SetFocus
    
End Sub

Private Sub scrSpeed_Scroll()

    speed = Val(scrSpeed.Value)
    
    cmdPlay.SetFocus
    
End Sub

Private Sub scrVolume_Change()

    volume = Val(scrVolume.Value)
    
    cmdPlay.SetFocus
    
End Sub

Private Sub scrVolume_Scroll()

    volume = Val(scrVolume.Value)
    
    cmdPlay.SetFocus
    
End Sub

Private Sub LoadScale()
    
    If scaleName = "" Then
        Exit Sub
    End If
    
    Open App.Path & "\Scales\" & scaleName For Input As #1
        
        Input #1, scaleName
            txtScale.Text = scaleName
        Input #1, rNum1
        Input #1, rDiv1
            ratio1 = rNum1 / rDiv1
        Input #1, rNum2
            txtNum2.Text = rNum2
        Input #1, rDiv2
            txtDiv2.Text = rDiv2
            ratio2 = rNum2 / rDiv2
        Input #1, rNum3
            txtNum3.Text = rNum3
        Input #1, rDiv3
            txtDiv3.Text = rDiv3
            ratio3 = rNum3 / rDiv3
        Input #1, rNum4
            txtNum4.Text = rNum4
        Input #1, rDiv4
            txtDiv4.Text = rDiv4
            ratio4 = rNum4 / rDiv4
        Input #1, rNum5
            txtNum5.Text = rNum5
        Input #1, rDiv5
            txtDiv5.Text = rDiv5
            ratio5 = rNum5 / rDiv5
        Input #1, rNum6
        Input #1, rDiv6
            ratio6 = rNum6 / rDiv6
    
    Close #1
        
End Sub

Private Sub Pause()
    
    'Set a random delay between notes
    
    Randomize
    
    gust = 1 + Int(Rnd(1) * 10)
    
    delay = (Rnd(1) * scrSpeed.Value) * 5
        
    If gust < 6 Then
        delay = delay / gust
    End If
    If gust = 6 Then
        delay = delay
    End If
    If gust > 6 Then
        delay = delay * gust
    End If
    
    Sleep delay
    
End Sub

Private Sub txtCents_Change()

    txtRatio.Text = ""
    
    If txtCents.Text <> "" And Val(txtCents.Text) < 1200 Then
        If Not IsNumeric(txtCents.Text) Then
            txtCents.Text = ""
            Exit Sub
        End If
        cents = CDbl(txtCents.Text)
        dec = Round(2 ^ (cents / 1200), 4)
        Dec2Frac (dec)
    End If
    
End Sub

Private Sub txtChime_Click()

    dlgFileSelect.File1.Path = App.Path & "\Chimes\"
    dlgFileSelect.Show vbModal
    chimeStyle = FileName
    ChimeFile (chimeStyle)
    
End Sub

Private Sub txtDiv2_Change()

    rDiv2 = Val(txtDiv2.Text)
    If Not IsNumeric(rDiv2) Then
        txtDiv2.Text = ""
        Exit Sub
    End If
    
    If Not txtDiv2.Text = "" Then
        ratio2 = rNum2 / rDiv2
    End If
    
End Sub

Private Sub txtDiv2_GotFocus()

    txtDiv2.Text = ""
    
End Sub

Private Sub txtDiv3_Change()

    rDiv3 = Val(txtDiv3.Text)
    If Not IsNumeric(rDiv3) Then
        txtDiv3.Text = ""
        Exit Sub
    End If
    
    If Not txtDiv3.Text = "" Then
        ratio3 = rNum3 / rDiv3
    End If
    
End Sub

Private Sub txtDiv3_GotFocus()

    txtDiv3.Text = ""
    
End Sub

Private Sub txtDiv4_Change()

    rDiv4 = Val(txtDiv4.Text)
    If Not IsNumeric(rDiv4) Then
        txtDiv4.Text = ""
        Exit Sub
    End If
    
    If Not txtDiv4.Text = "" Then
        ratio4 = rNum4 / rDiv4
    End If
    
End Sub

Private Sub txtDiv4_GotFocus()

    txtDiv4.Text = ""
    
End Sub

Private Sub txtDiv5_Change()

    rDiv5 = Val(txtDiv5.Text)
    If Not IsNumeric(rDiv5) Then
        txtDiv5.Text = ""
        Exit Sub
    End If
    
    If Not txtDiv5.Text = "" Then
        ratio5 = rNum5 / rDiv5
    End If
    
End Sub

Private Sub txtDiv5_GotFocus()

    txtDiv5.Text = ""
    
End Sub

Private Sub txtNum2_Change()

    rNum2 = Val(txtNum2.Text)
    If Not IsNumeric(rNum2) Then
        txtNum2.Text = ""
        Exit Sub
    End If
    
    If Not txtNum2.Text = "" Then
        ratio2 = rNum2 / 1
    End If
    
End Sub

Private Sub txtNum2_GotFocus()

    txtNum2.Text = ""
    
End Sub

Private Sub txtNum3_Change()

    rNum3 = Val(txtNum3.Text)
    If Not IsNumeric(rNum3) Then
        txtNum3.Text = ""
        Exit Sub
    End If
    
    If Not txtNum3.Text = "" Then
        ratio3 = rNum3 / 1
    End If
    
End Sub

Private Sub txtNum3_GotFocus()

    txtNum3.Text = ""
    
End Sub

Private Sub txtNum4_Change()

    rNum4 = Val(txtNum4.Text)
    If Not IsNumeric(rNum4) Then
        txtNum4.Text = ""
        Exit Sub
    End If
    
    If Not txtNum4.Text = "" Then
        ratio4 = rNum4 / 1
    End If
    
End Sub

Private Sub txtNum4_GotFocus()

    txtNum4.Text = ""
    
End Sub

Private Sub txtNum5_Change()

    rNum5 = Val(txtNum5.Text)
    If Not IsNumeric(rNum5) Then
        txtNum5.Text = ""
        Exit Sub
    End If
    
    If Not txtNum5.Text = "" Then
        ratio5 = rNum5 / 1
    End If
    
End Sub

Private Sub txtNum5_GotFocus()

    txtNum5.Text = ""
    
End Sub

Private Sub txtPreset_Click()

    dlgFileSelect.File1.Path = App.Path & "\Presets\"
    dlgFileSelect.Show vbModal
    
    If FileName <> "" Then
        For i = 0 To 5
            If Not (dsBuffer(i) Is Nothing) Then
                dsBuffer(i).Stop
                Set dsBuffer(i) = Nothing
            End If
        Next i
        PresetFile
    End If
    
End Sub

Private Sub txtScale_Click()

    dlgFileSelect.File1.Path = App.Path & "\Scales\"
    dlgFileSelect.Show vbModal
    scaleName = FileName
    ScaleFile (scaleName)

End Sub

Private Function Dec2Frac(ByVal dec As Double) As String
        
    fNum = 1
    fDiv = 1
    
    df = fNum / fDiv
    
    While (df <> dec)
        If (df < dec) Then
            fNum = fNum + 1
        Else
            fDiv = fDiv + 1
            fNum = dec * fDiv
        End If
        df = fNum / fDiv
    Wend
    
    txtRatio.Text = CStr(fNum) & "/" & CStr(fDiv)

End Function

