VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Api Error Tester"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   FillColor       =   &H00800000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Execute"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Comments and suggestions r welcomed."
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Although this example looks harmless, you never know with windows so use at your own risk."
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   5415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "John Galanopoulos, GreekThought@yahoo.gr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   5535
   End
   Begin VB.Label Label3 
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "(This doesn't raise an error, but it returns the error description from the system.)"
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   2880
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "System Error"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
On Error Resume Next
   Text2.Text = LastDLLErrorDescription(CLng(Text1.Text))
End Sub

Private Sub Command2_Click()
  End
End Sub
