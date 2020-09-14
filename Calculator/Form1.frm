VERSION 5.00
Begin VB.Form Calculator 
   BackColor       =   &H0000C000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Simple Calculator"
   ClientHeight    =   5280
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   7530
   DrawStyle       =   6  'Inside Solid
   FontTransparent =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7530
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C000C0&
      Caption         =   "Simple calculator"
      ForeColor       =   &H80000014&
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton Command1 
         Caption         =   "About Application"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5280
         MouseIcon       =   "Form1.frx":0CCA
         TabIndex        =   29
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Frame Frame4 
         Caption         =   "Back Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   5280
         TabIndex        =   24
         Top             =   1440
         Width           =   1695
         Begin VB.OptionButton Option2 
            Caption         =   "Option2"
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   1080
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   25
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Blue"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   28
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Black"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   26
            ToolTipText     =   "You are using Black color"
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdDot 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   1320
         TabIndex        =   21
         Top             =   3840
         Width           =   500
      End
      Begin VB.CommandButton cmdx 
         Caption         =   "1/x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   3480
         TabIndex        =   20
         Top             =   3840
         Width           =   500
      End
      Begin VB.CommandButton cmdper 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   4200
         TabIndex        =   19
         Top             =   3120
         Width           =   500
      End
      Begin VB.CommandButton cmdsqr 
         Caption         =   "Sqrt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   3480
         TabIndex        =   18
         Top             =   3120
         Width           =   500
      End
      Begin VB.CommandButton Command2 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   2160
         TabIndex        =   17
         Top             =   3840
         Width           =   500
      End
      Begin VB.CommandButton cmdequals 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   4200
         TabIndex        =   16
         Top             =   3840
         Width           =   500
      End
      Begin VB.CommandButton cmdsign 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   3
         Left            =   4200
         TabIndex        =   15
         Top             =   2400
         Width           =   500
      End
      Begin VB.CommandButton cmdsign 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   2
         Left            =   3480
         TabIndex        =   14
         Top             =   2400
         Width           =   500
      End
      Begin VB.CommandButton cmdsign 
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   1
         Left            =   4200
         TabIndex        =   13
         Top             =   1680
         Width           =   500
      End
      Begin VB.CommandButton cmdsign 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   0
         Left            =   3480
         TabIndex        =   12
         Top             =   1680
         Width           =   500
      End
      Begin VB.CommandButton cmdbut 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   9
         Left            =   2160
         TabIndex        =   11
         Top             =   3120
         Width           =   500
      End
      Begin VB.CommandButton cmdbut 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   8
         Left            =   1320
         TabIndex        =   10
         Top             =   3120
         Width           =   500
      End
      Begin VB.CommandButton cmdbut 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   7
         Left            =   600
         TabIndex        =   9
         Top             =   3120
         Width           =   500
      End
      Begin VB.CommandButton cmdbut 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   6
         Left            =   2160
         TabIndex        =   8
         Top             =   2400
         Width           =   500
      End
      Begin VB.CommandButton cmdbut 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   5
         Left            =   1320
         TabIndex        =   7
         Top             =   2400
         Width           =   500
      End
      Begin VB.CommandButton cmdbut 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   4
         Left            =   600
         TabIndex        =   6
         Top             =   2400
         Width           =   500
      End
      Begin VB.CommandButton cmdbut 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   3
         Left            =   2160
         TabIndex        =   5
         Top             =   1680
         Width           =   500
      End
      Begin VB.CommandButton cmdbut 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   2
         Left            =   1320
         TabIndex        =   4
         Top             =   1680
         Width           =   500
      End
      Begin VB.CommandButton cmdbut 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   1
         Left            =   600
         TabIndex        =   3
         Top             =   1680
         Width           =   500
      End
      Begin VB.CommandButton cmdbut 
         BackColor       =   &H80000007&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   0
         Left            =   600
         TabIndex        =   2
         Top             =   3840
         Width           =   500
      End
      Begin VB.TextBox Text1 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   4335
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   3135
         Left            =   360
         TabIndex        =   22
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   3360
         TabIndex        =   23
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.Menu Exit 
      Caption         =   "E&xit Calci"
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim first, second As Double
Dim sign  As String
Dim intres, dvn As Integer

Private Sub About_Click()
frmabout.Show
End Sub

Private Sub advanced_Click()
Calculator.Show
cmdsqr.Enabled = True
cmdper.Enabled = True
cmdx.Enabled = True
End Sub

Private Sub cmdbut_Click(Index As Integer)
Text1.Text = Text1.Text & cmdbut(Index).Caption
End Sub

Private Sub cmdDot_Click()
sign = "."
Text1.Text = Text1.Text + sign
first = Text1.Text
End Sub

Private Sub cmdper_Click()
Text1.Text = Text1.Text * 1 / 100
End Sub

Private Sub cmdsign_Click(Index As Integer)
first = Text1.Text
Text1.Text = ""
sign = cmdsign(Index).Caption
Text1.Text = ""

End Sub
Private Sub cmdequals_Click()
second = Text1.Text
  If sign = "+" Then
Text1.Text = first + second
ElseIf sign = "--" Then
Text1.Text = first - second
ElseIf sign = "X" Then
Text1.Text = first * second
ElseIf sign = "/" Then
    If first = 0 And second = 0 Then
    dvn = MsgBox(" !!! Indeterminant", , "Division Error")
    ElseIf first <> 0 And second = 0 Then
    dvn = MsgBox("  !!! Infinity ", , _
    "Division Error")
    Text1.Text = ""
    Else
    Text1.Text = first / second
End If
End If
End Sub

Private Sub cmdsqr_Click()
Text1.Text = Sqr(Text1.Text)
End Sub

Private Sub cmdx_Click()
Text1.Text = 1 / (Text1.Text)
End Sub

Private Sub Command1_Click()
frmabout.Show
End Sub

Private Sub Command2_Click()
Text1.Text = ""
End Sub


Private Sub Exit_Click()
Unload Me
End Sub

Private Sub Exitcalci_Click()

End Sub

Private Sub Label1_Click(Index As Integer)
Frame1.BackColor = vbGreen
End Sub



Private Sub Option1_Click(Index As Integer)
Frame1.BackColor = vbBlack
End Sub


Private Sub Option2_Click()
Frame1.BackColor = vbBlue
End Sub

