VERSION 5.00
Begin VB.Form frmSendGetData 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "frmSendGetData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer fancy 
      Interval        =   100
      Left            =   5520
      Top             =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Text here"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1440
      TabIndex        =   0
      Top             =   2880
      Width           =   3180
   End
End
Attribute VB_Name = "frmSendGetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Created with DragonSource Diamond. Copyright � 2012 DragonSource


Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyEscape) Then
        Call DestroyDirectX
        Call StopMidi
        InGame = False
        frmMain.Socket.Close
        frmMainMenu.Visible = True
        Connucted = False
        Unload Me
    End If
End Sub

Private Sub Form_Load()
'empty sub
End Sub

Private Sub Label1_Click()
    Call DestroyDirectX
    Call StopMidi
    InGame = False
    frmMain.Socket.Close
    frmMainMenu.Visible = True
    Connucted = False
    Unload Me
End Sub
