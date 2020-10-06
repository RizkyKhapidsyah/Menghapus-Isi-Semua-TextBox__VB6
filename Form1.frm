VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menghapus Isi Semua TextBox"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hapus"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   2280
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
Dim Contrl As Control
  For Each Contrl In Form1.Controls
    If (TypeOf Contrl Is TextBox) Then Contrl.Text = ""
  Next Contrl
End Sub

