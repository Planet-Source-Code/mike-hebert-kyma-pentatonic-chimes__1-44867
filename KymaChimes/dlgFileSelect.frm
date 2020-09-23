VERSION 5.00
Begin VB.Form dlgFileSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kyma File Selector"
   ClientHeight    =   2415
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2355
   Icon            =   "dlgFileSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   2355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   855
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   2160
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   2160
      X2              =   2160
      Y1              =   120
      Y2              =   2280
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
      Y2              =   2280
   End
End
Attribute VB_Name = "dlgFileSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    
    FileName = ""
    Unload Me
    
End Sub

Private Sub File1_Click()

    OKButton.SetFocus
    
End Sub
Private Sub File1_DblClick()

    OKButton_Click
    
End Sub

Private Sub OKButton_Click()

    FileName = File1.FileName
    
    Unload Me
    
End Sub
