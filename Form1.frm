VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Transparent window"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMiscOpt 
      Caption         =   "Transparent window"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton cmdTTurnOn 
         Caption         =   "Turn On"
         Height          =   255
         Left            =   4080
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin MSComctlLib.Slider sldTransparent 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         Min             =   40
         Max             =   255
         SelStart        =   230
         TickFrequency   =   13
         Value           =   230
      End
      Begin VB.Label Label1 
         Caption         =   "Coded by SonicBlue"
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblTFScale 
         Caption         =   "Transparency Frequency Scale:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblTFRate 
         Caption         =   "Rate:"
         Height          =   255
         Left            =   3480
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblRate 
         Caption         =   "230"
         Height          =   255
         Left            =   3480
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTTurnOn_Click()
If cmdTTurnOn.Caption = "Turn On" Then
    MakeTransparent Me.hWnd, 230
    MakeTransparent Form1.hWnd, 230
    cmdTTurnOn.Caption = "Turn Off"
ElseIf cmdTTurnOn.Caption = "Turn Off" Then
    MakeOpaque Form1.hWnd
    MakeOpaque Me.hWnd
    cmdTTurnOn.Caption = "Turn On"
End If
End Sub

Private Sub sldTransparent_Click()
    lblRate.Caption = sldTransparent.Value
    MakeTransparent Me.hWnd, sldTransparent.Value
    MakeTransparent Form1.hWnd, sldTransparent.Value
If sldTransparent.Value Then
cmdTTurnOn.Caption = "Turn Off"
End If
End Sub
