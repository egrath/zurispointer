VERSION 5.00
Begin VB.Form frmDataBuilder 
   Caption         =   "Data Builder"
   ClientHeight    =   960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3570
   Icon            =   "Data Builder.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   960
   ScaleWidth      =   3570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lbl 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmDataBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGo_Click()
    Me.Caption = "Working"
    WriteFile
End Sub

Private Sub Form_Load()
    OpenData
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    CleanUp
End Sub
