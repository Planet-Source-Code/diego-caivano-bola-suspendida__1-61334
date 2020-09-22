VERSION 5.00
Begin VB.Form FrmBolaSusp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bola Suspendida"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   3630
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBolaSusp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "FrmBolaSusp.frx":030A
   ScaleHeight     =   8325
   ScaleWidth      =   3630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkMarcador 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1650
      TabIndex        =   8
      Top             =   7800
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "Cmd1"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   7200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton CmdAyuda 
      Caption         =   "&Ayuda"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Timer TmrDificultad 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2760
      Top             =   4320
   End
   Begin VB.Timer TmrTiempo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   4320
   End
   Begin VB.CommandButton CmdEmpezar 
      Caption         =   "&Empezar"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   6720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton CmdDetener 
      Caption         =   "&Detener"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.Timer TmrBola 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   4320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marcador Lateral"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1920
      TabIndex        =   9
      Top             =   7815
      Width           =   1440
   End
   Begin VB.Label LblMejorTpo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mejor Tiempo: 0 Seg."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1440
      TabIndex        =   6
      Top             =   3120
      Width           =   2040
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080C0FF&
      Height          =   375
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   2565
      Width           =   2055
   End
   Begin VB.Line Lne1 
      BorderColor     =   &H00C0C0C0&
      X1              =   1320
      X2              =   1320
      Y1              =   120
      Y2              =   8160
   End
   Begin VB.Label Limite 
      BackStyle       =   0  'Transparent
      Height          =   8055
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label LblTiempo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tiempo: 0 Seg."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   2640
      Width           =   2040
   End
   Begin VB.Shape Shp1 
      BorderColor     =   &H00FFFFFF&
      Height          =   8055
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape Bola 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFC0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   630
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shp2 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00C0C0C0&
      Height          =   8055
      Left            =   405
      Shape           =   4  'Rounded Rectangle
      Top             =   160
      Width           =   735
   End
   Begin VB.Line Lne2 
      BorderColor     =   &H00C0C0C0&
      X1              =   1200
      X2              =   1320
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Menu MnuNuevo 
      Caption         =   "&Nuevo"
   End
   Begin VB.Menu MnuAcerca 
      Caption         =   "A&cerca"
   End
End
Attribute VB_Name = "FrmBolaSusp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Acu, G, T, Max As Integer

Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) _
As Long 'Speaker

'Función Para Reproducción de Sonidos
Private Declare Function SndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    Const SND_SYNC = &H0
    Const SND_ASYNC = &H1
    Const SND_NODEFAULT = &H2
    Const SND_LOOP = &H8
    Const SND_NOSTOP = &H10

Private Sub ChkMarcador_Click()
    With ChkMarcador
        If .Value = 1 Then
            Lne1.Visible = True
            Lne2.Visible = True
        Else
            Lne1.Visible = False
            Lne2.Visible = False
        End If
    End With
End Sub

Private Sub Cmd1_Click()
    CmdDetener_Click
    SndPlaySound App.Path & "\ElectricTransp.Wav", SND_ASYNC Or SND_NODEFAULT
    
    'Setear Posiciones Iniciales De Los Elementos
    Bola.Top = 3960
    Shp1.Height = 8055
    Shp1.Top = 120
    Shp2.Height = 8055
    Shp2.Top = 160
    Limite.Height = 8055
    Limite.Top = 120
    
    T = 0 'Resetear Tiempo
    LblTiempo.Caption = "Tiempo: " & T & " Seg."
    LblMejorTpo.Caption = "Mejor Tiempo: " & Max & " Seg."
End Sub

Private Sub CmdEmpezar_Click()
    Acu = 80 'Impulsar La Bola Hacia Arriba
    TmrBola.Enabled = True
    TmrTiempo.Enabled = True
    TmrDificultad.Enabled = True
    
    CmdEmpezar.Enabled = False
    CmdDetener.Enabled = True
End Sub

Private Sub CmdDetener_Click()
    TmrBola.Enabled = False
    TmrTiempo.Enabled = False
    TmrDificultad.Enabled = False
    
    CmdEmpezar.Enabled = True
    CmdDetener.Enabled = False
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdAyuda_Click()
    MsgBox "Haga Click Dentro Del Rectángulo Vertical Para Comenzar El Juego e Impulsar La Bola Hacia Arriba. Intente Mantenerla Suspendida El Mayor Tiempo Posible Sin Tocar Los Límites.", _
    vbInformation, "Ayuda - Instrucciones"
End Sub

Private Sub Form_Load()
    SndPlaySound App.Path & "\ElectricTransp.Wav", SND_ASYNC Or SND_NODEFAULT
    Acu = 80
    G = 3 'Fuerza De Gravedad
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Limite_Click()
    CmdEmpezar_Click
End Sub

Private Sub MnuAcerca_Click()
    FrmAcerca.Show
End Sub

Private Sub MnuNuevo_Click()
    Max = 0 'Resetear Mejor Tiempo
    Cmd1_Click
End Sub

Private Sub TmrBola_Timer()
    Acu = Acu - G
    With Bola
        .Top = .Top - Acu
        
        'Mover Marcador Lateral
        Lne2.Y1 = .Top + .Height / 2
        Lne2.Y2 = .Top + .Height / 2
        
        Select Case .Top
            Case Is >= Limite.Top + Limite.Height - .Height 'Si Toca El Piso
                TmrDificultad.Enabled = False 'Detener El Achicamiento Del Rectángulo
                TmrTiempo.Enabled = False
                .Top = Limite.Top + Limite.Height - .Height 'Corregir Error Decimal
                MsgBox "La Bola Ha Tocado El Límite Inferior; Su Tiempo Fue De " & T & " Seg.", _
                vbInformation, "¡Ha Perdido!" 'Aviso De Pérdida
                If T > Max Then Max = T 'Guardar Mejor Tiempo
                Cmd1_Click
            Case Is <= Limite.Top 'Si Toca El Techo
                TmrDificultad.Enabled = False
                TmrTiempo.Enabled = False
                .Top = Limite.Top
                MsgBox "La Bola Ha Tocado El Límite Superior; Su Tiempo Fue De " & T & " Seg.", _
                vbInformation, "¡Ha Perdido!"
                If T > Max Then Max = T
                Cmd1_Click
        End Select
    End With
End Sub

Private Sub TmrDificultad_Timer()
    'Achicar Límite
    Shp1.Height = Shp1.Height - 4
    Shp1.Top = Shp1.Top + 2
    Shp2.Height = Shp2.Height - 4
    Shp2.Top = Shp2.Top + 2
    Limite.Height = Limite.Height - 4
    Limite.Top = Limite.Top + 2
End Sub

Private Sub TmrTiempo_Timer()
    SndPlaySound App.Path & "\MineTick.Wav", SND_ASYNC Or SND_NODEFAULT
    T = T + 1 'Contar Segundos
    LblTiempo.Caption = "Tiempo: " & T & " Seg."
End Sub
