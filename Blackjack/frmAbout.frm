VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Blackjack"
   ClientHeight    =   3570
   ClientLeft      =   2475
   ClientTop       =   2340
   ClientWidth     =   5655
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2464.077
   ScaleMode       =   0  'User
   ScaleWidth      =   5310.337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   3120
      Width           =   1454
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Coded by Daniel Brady."
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2760
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "This is freeware, you can distribute it freely as long as it stays free."
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   1470
      Left            =   1800
      Picture         =   "frmAbout.frx":0000
      Top             =   120
      Width           =   1965
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Blackjack 1.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   5535
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
Unload Me
End Sub
