VERSION 5.00
Begin VB.Form frmabout 
   BackColor       =   &H008080FF&
   Caption         =   "About calculator"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3030
   ClipControls    =   0   'False
   FillColor       =   &H00FF80FF&
   ForeColor       =   &H80000017&
   LinkTopic       =   "Form1"
   MousePointer    =   4  'Icon
   ScaleHeight     =   1755
   ScaleWidth      =   3030
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Pannagesh"
      Top             =   720
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2280
      Picture         =   "frmabout.frx":0000
      Top             =   360
      Width           =   1080
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Programmer"
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub




Private Sub cmdOk_Click()
Unload Me
End Sub


