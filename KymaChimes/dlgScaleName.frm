VERSION 5.00
Begin VB.Form dlgScaleName 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scale Titler"
   ClientHeight    =   1935
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2190
   Icon            =   "dlgScaleName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   2190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtScaleName 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Accept"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   2040
      X2              =   2040
      Y1              =   120
      Y2              =   1800
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   120
      X2              =   2040
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   2040
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a Scale Name ..."
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "dlgScaleName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()

    Unload Me
    
End Sub

Private Sub OKButton_Click()

    If txtScaleName.Text = "" Then
        Exit Sub
    End If
    
    Open App.Path & "\Scales\" & txtScaleName.Text & ".txt" _
        For Output As #1
        
        Write #1, txtScaleName.Text
        Write #1, frmChimes.txtNum1.Text
        Write #1, frmChimes.txtDiv1.Text
        Write #1, frmChimes.txtNum2.Text
        Write #1, frmChimes.txtDiv2.Text
        Write #1, frmChimes.txtNum3.Text
        Write #1, frmChimes.txtDiv3.Text
        Write #1, frmChimes.txtNum4.Text
        Write #1, frmChimes.txtDiv4.Text
        Write #1, frmChimes.txtNum5.Text
        Write #1, frmChimes.txtDiv5.Text
        Write #1, frmChimes.txtNum6.Text
        Write #1, frmChimes.txtDiv6.Text
        
    Close #1
    
    frmChimes.txtScale.Text = txtScaleName.Text
    
    Unload Me
    
End Sub
