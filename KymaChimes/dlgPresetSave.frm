VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form dlgPresetSave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kyma Preset Titler"
   ClientHeight    =   2910
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2280
   Icon            =   "dlgPresetSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   2280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox txtDesc 
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1931
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"dlgPresetSave.frx":0CCA
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   855
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   2160
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   2160
      X2              =   2160
      Y1              =   120
      Y2              =   2760
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   2160
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   2760
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Optional description"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Title of your new preset"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "dlgPresetSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()

    Unload Me
    
End Sub

Private Sub OKButton_Click()

    If txtTitle.Text = "" Then
        Exit Sub
    End If
    
    Open App.Path & "\Presets\" & presetName _
        For Output As #1
        
        Write #1, txtTitle.Text
        Write #1, frmChimes.txtScale.Text
        Write #1, frmChimes.txtChime.Text & ".wav"
        If presetDesc = "" Then
            presetDesc = "No description available."
        End If
        Write #1, presetDesc
        Write #1, Str(frmChimes.scrSpeed.Value)
        Write #1, Str(frmChimes.scrPitch.Value)
        Write #1, Str(frmChimes.scrVolume.Value)
        
    Close #1
    
    frmChimes.txtPreset.Text = txtTitle.Text
    frmChimes.txtDescription.Text = presetDesc
    
    Unload Me
    
End Sub

Private Sub txtDesc_Change()

    presetDesc = txtDesc.Text
    
End Sub

Private Sub txtTitle_Change()
    
    presetName = txtTitle.Text & ".txt"
    
End Sub
