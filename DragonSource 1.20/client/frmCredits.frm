VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Engine Credits"
   ClientHeight    =   6000
   ClientLeft      =   165
   ClientTop       =   390
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCredits.frx":1CCA
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      Caption         =   "Back To Main Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2430
      TabIndex        =   5
      Top             =   5400
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Website: http://www.source.draignet.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   375
      TabIndex        =   4
      Top             =   1560
      Width           =   3030
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DragonSource 1.2 is a re-release of the Konfuze Engine originally created by Sean (GSD), Kieran Gill and Liam Stewart."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   5160
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Primary Programmer: Liam Stewart"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   375
      TabIndex        =   2
      Top             =   2910
      Width           =   2280
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmer: Kieran Gill"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   375
      TabIndex        =   1
      Top             =   3180
      Width           =   2280
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmer: Godsentdeath"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   375
      TabIndex        =   0
      Top             =   3435
      Width           =   2055
   End
End
Attribute VB_Name = "frmCredits"
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
    frmCredits.Visible = False
End Sub

