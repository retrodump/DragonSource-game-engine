VERSION 5.00
Begin VB.Form frmDeleteAccount 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Account"
   ClientHeight    =   6000
   ClientLeft      =   195
   ClientTop       =   405
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmDeleteAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDeleteAccount.frx":1CCA
   ScaleHeight     =   6000
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   450
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2400
      Width           =   2460
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   450
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1755
      Width           =   2460
   End
   Begin VB.Label picCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Back To Main Menu"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2220
      TabIndex        =   5
      Top             =   5385
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackColor       =   &H00789298&
      BackStyle       =   0  'Transparent
      Caption         =   "Account Name"
      Height          =   225
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Password"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   2205
      Width           =   1545
   End
   Begin VB.Label picConnect 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Your Account"
      Height          =   210
      Left            =   480
      TabIndex        =   2
      Top             =   2910
      Width           =   2025
   End
End
Attribute VB_Name = "frmDeleteAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##################################'
'##### MADE WITH DRAGONSOURCE #####'
'# http://www.source.draignet.com #'
'##################################'

Option Explicit

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmDeleteAccount.Visible = False
End Sub

Private Sub picConnect_Click()
    If Trim(txtName.Text) <> "" And Trim(txtPassword.Text) <> "" Then
        Call MenuState(MENU_STATE_DELACCOUNT)
    End If
End Sub

