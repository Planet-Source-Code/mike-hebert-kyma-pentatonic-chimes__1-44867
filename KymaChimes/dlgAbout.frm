VERSION 5.00
Begin VB.Form dlgAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Kyma Pentatonic Chimes"
   ClientHeight    =   3615
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3510
   Icon            =   "dlgAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   1575
      Left            =   240
      Picture         =   "dlgAbout.frx":0CCA
      ScaleHeight     =   1515
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton cmdThankYou 
      Caption         =   "Thank You!"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Kyma Pentonic Chimes is copyright 2003 by Michael Hebert - Kyma Software"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label lblHTTP 
      Alignment       =   2  'Center
      Caption         =   "http://kymasoft.netfirms.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1880
      Width           =   3015
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   3360
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   3360
      X2              =   3360
      Y1              =   120
      Y2              =   3480
   End
   Begin VB.Label lblMail 
      Alignment       =   2  'Center
      Caption         =   "email: kymasoft@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   3360
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "dlgAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdThankYou_Click()
    
    Unload Me
    
End Sub

Private Sub lblHTTP_Click()

    Shell "start " & "http://kymasoft.netfirms.com/"
End Sub

Private Sub lblMail_Click()
    
    Shell "start " & "mailto:kymasoft@hotmail.com"
    
End Sub
