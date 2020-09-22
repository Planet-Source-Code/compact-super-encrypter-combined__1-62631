VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Coder"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Text            =   "this is your password"
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   5040
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decrypt"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "crypt-sample.frx":0000
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub Command1_Click()
Text2.Text = encrypt(Text1.Text, Text4.Text)
DoEvents
End Sub

Private Sub Command2_Click()
Text3.Text = decrypt(Text2.Text, Text4.Text)
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub
