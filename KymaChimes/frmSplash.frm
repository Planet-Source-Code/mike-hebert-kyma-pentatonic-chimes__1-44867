VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5070
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4875
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5385
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   1920
         Top             =   3120
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Pentatonic"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   1815
         TabIndex        =   6
         Top             =   1800
         Width           =   3285
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   5
         Top             =   3240
         Width           =   1395
      End
      Begin VB.Label lblWarning 
         Alignment       =   2  'Center
         Caption         =   "This program requires DirectX 7.0 or later"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   4560
         Width           =   5055
      End
      Begin VB.Label lblCompany 
         Caption         =   "By Michael Hebert - Kyma Software"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   3960
         Width           =   2775
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright 2003"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   3720
         Width           =   2775
      End
      Begin VB.Image imgLogo 
         Height          =   4185
         Left            =   120
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Chimes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2250
         TabIndex        =   1
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   1500
         Left            =   1830
         Picture         =   "frmSplash.frx":19FA
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3345
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    frmChimes.Show
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = "Pentatonic"
End Sub

Private Sub Frame1_Click()
    frmChimes.Show
    Unload Me
End Sub

Private Sub Timer1_Timer()
    While Timer1.Interval > 0
        DoEvents
        Timer1.Interval = Timer1.Interval - 1
    Wend
    Frame1_Click
End Sub
