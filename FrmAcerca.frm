VERSION 5.00
Begin VB.Form FrmAcerca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca de Bola Suspendida"
   ClientHeight    =   1815
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5745
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1252.745
   ScaleMode       =   0  'User
   ScaleWidth      =   5394.852
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   480
      Picture         =   "FrmAcerca.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   360
      Width           =   540
   End
   Begin VB.CommandButton CmdOk 
      Cancel          =   -1  'True
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   4200
      TabIndex        =   0
      Top             =   1320
      Width           =   1140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   776.495
      Y2              =   776.495
   End
   Begin VB.Label Lbl1 
      AutoSize        =   -1  'True
      Caption         =   "Bola Suspendida."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1320
      TabIndex        =   3
      Top             =   240
      Width           =   1485
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   786.848
      Y2              =   786.848
   End
   Begin VB.Label Lbl2 
      AutoSize        =   -1  'True
      Caption         =   "Juego de Habilidad."
      Height          =   195
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   1665
   End
   Begin VB.Label LblProg 
      AutoSize        =   -1  'True
      Caption         =   "Programador: Diego Caivano © 2005."
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1305
      Width           =   3285
   End
End
Attribute VB_Name = "FrmAcerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()
    Unload Me
End Sub
