VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   Caption         =   "Help for Kyma Pentatonic Chimes"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About Kyma Chimes"
      Height          =   375
      Left            =   3465
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit Help"
      Height          =   375
      Left            =   1665
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox rtfHelp 
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6165
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmHelp.frx":0CCA
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   6600
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   6600
      X2              =   6600
      Y1              =   120
      Y2              =   4320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   6600
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   4320
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbout_Click()

    dlgAbout.Show vbModal
    
End Sub

Private Sub cmdExit_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    rtfHelp.FileName = App.Path & "/KymaChimes.rtf"
    
End Sub
