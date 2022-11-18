VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2670
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   2670
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2835
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.TextBox txtTotalPLBA 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton btnCalc 
         Caption         =   "&Calculate PLBA"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtPFrame 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtPSec 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtPMin 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTotalPLBA 
         Alignment       =   2  'Center
         Caption         =   "Total PLBA value:"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label lblPFrame 
         Caption         =   "PFrame:"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblPSec 
         Caption         =   "PSec:"
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblPMin 
         Caption         =   "PMin:"
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCalc_Click()

	If txtPMin.Text = "" Then
		GoTo Error1
	Else
	If txtPSec.Text = "" Then
		GoTo Error1
	Else
	If txtPFrame.Text = "" Then
		GoTo Error1
	End If

On Error GoTo Error2

	txtTotalPLBA.Enabled = True
	txtTotalPLBA.Text = Str(Val(txtPMin.Text * 60) + Val(txtPSec.Text)) * 75 + Val(txtPFrame.Text) - 150
	
	GoTo Trap

Error1:
	MsgBox "Please fill in all of the fields.", vbCritical, "Error"
	GoTo Trap

Error2:
	MsgBox "Invalid characters.", vbCritical, "Error"
	GoTo Trap

Trap:
End Sub

