VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmPantalla 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   10740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "frmPantalla.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10740
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMover 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2040
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   3
      Top             =   1920
      Width           =   375
      Begin VB.Shape cursor 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   2  'Oval
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Timer timerReloj 
      Interval        =   17
      Left            =   360
      Top             =   1320
   End
   Begin WMPLibCtl.WindowsMediaPlayer rep 
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1085
      _cy             =   873
   End
   Begin VB.Label lblconiguracíon 
      BackStyle       =   0  'Transparent
      Caption         =   "&Configuracíon"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label lbldigital 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "LCD"
         Size            =   369.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   7305
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1755
   End
   Begin VB.Label lbltime 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "LCD"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbldigital 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "LCD"
         Size            =   369.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   7305
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   7530
   End
End
Attribute VB_Name = "frmPantalla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    IncrementoDecremento KeyAscii, KeyAscii
End Sub

Private Sub IncrementoDecremento(ByVal incremento As Integer, ByVal decremento _
As Integer)
Me.Caption = FUNCIONES.LCDContador
'----------------------------------------------------------------------------
'mostrar Tablero LCD en el Lienso e Incrementar o Decrementar               -
'----------------------------------------------------------------------------
If FUNCIONES.LCDContador = 100 Then ' si el contador interno  es igual a 100
   FUNCIONES.LCDContador = 0        ' el contador interno se restablece a  0
End If                              ' finalisa la istrucion si
'****************************************************************************
   If FUNCIONES.TecIncremento = incremento Then 'si la tecla como resultado es
                                                'igual a la alojada en la base
                                                'se ejecuta el contador
        If FUNCIONES.LCDContador <= 9 Then      'si el conteo es menor o igual a 09
            Me.lbldigital(1).Caption = 0 _
            & FUNCIONES.LCDContador             'incrementa el contador de uno en 1
            FUNCIONES.reproducir                'reproduce el sonido del Display
        ElseIf FUNCIONES.LCDContador >= 9 Then  'si el contador en mayor a 9 +
            Me.lbldigital(1).Caption = FUNCIONES.LCDContador
            FUNCIONES.reproducir                '""
        End If                                  'fin si
      FUNCIONES.LCDContador = FUNCIONES.LCDContador + 1
     End If

If FUNCIONES.NumeroInicial = 0 Then
    FUNCIONES.NumeroInicial = FUNCIONES.LCDContador
ElseIf FUNCIONES.NumeroInicial <> 0 Then
    FUNCIONES.NumeroInicial = FUNCIONES.LCDContador - 1
End If

If FUNCIONES.TecDescremento = decremento And FUNCIONES.LCDContador > 0 Then
    
    FUNCIONES.LCDContador = FUNCIONES.LCDContador - 1
    
    If FUNCIONES.LCDContador <= 9 Then
    Me.lbldigital(1).Caption = 0 & FUNCIONES.LCDContador
    FUNCIONES.reproducir
    ElseIf FUNCIONES.LCDContador >= 9 Then
    Me.lbldigital(1).Caption = FUNCIONES.LCDContador
    FUNCIONES.reproducir
    End If
    
End If
End Sub

Private Sub Form_Load()
        FUNCIONES.nodoEstado = False ' nodo invisible
        FUNCIONES.visualizarNODO FUNCIONES.nodoEstado
        FUNCIONES.LCDContador = FUNCIONES.NumeroInicial
        ModGurdarAbrir.AbrirArchivo
        cargarDatosPrograma
        lblconiguracíon.Caption = FUNCIONES.VectorTexto(15)
        
       
       IncrementoDecremento 1, 0
       
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift _
As Integer, X As Single, Y As Single)
Me.SetFocus
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift _
As Integer, X As Single, Y As Single)
IncrementoDecremento Button, Button
End Sub

Private Sub lblconiguracíon_Click()
FUNCIONES.visualizarLCD False
PROPIEDADESLCD.Show 1
End Sub

Private Sub lbldigital_MouseUp(Index As Integer, Button _
As Integer, Shift As Integer, X As Single, Y As Single)
IncrementoDecremento Button, Button
Me.SetFocus
End Sub

Private Sub picMover_MouseMove(Button As Integer, Shift _
As Integer, X As Single, Y As Single)
'************************************************************
' Desplasa el Display LCD por toda la Superficie del Lienso *
' o formulario si el boton es mayor a 0.                    *
'************************************************************
    If Button > 0 Then
    
        picMover.Move picMover.Left + X, picMover.top + Y
        lbldigital(0).Left = picMover.Left + 270: lbldigital(0).top = _
        picMover.top - 700: lbldigital(0).Refresh
        lbldigital(1).Left = picMover.Left + 270: lbldigital(1).top = _
        picMover.top - 700: lbldigital(1).Refresh
    
    End If

End Sub

Private Sub timerReloj_Timer()
lbltime.Caption = Time
    ModGurdarAbrir.gardarArchio
End Sub

Private Sub cargarDatosPrograma()
'**********************************************************
' carga todos los datos alojados al display y propiedades *
'**********************************************************
    lbldigital(0).top = FUNCIONES.top
    lbldigital(1).top = FUNCIONES.top
    lbldigital(0).Left = FUNCIONES.Size
    lbldigital(1).Left = FUNCIONES.Size
    frmPantalla.lbldigital(0).Font.Size = CInt(FUNCIONES.tamanioDisplay)
    frmPantalla.lbldigital(1).Font.Size = CInt(FUNCIONES.tamanioDisplay)
    PROPIEDADESLCD.lblCambio.Visible = True
    frmPantalla.lbldigital(1).ForeColor = FUNCIONES.ColorIluminado
    frmPantalla.lbltime.ForeColor = FUNCIONES.ColorIluminado
    FUNCIONES.ColorIluminado = FUNCIONES.ColorIluminado
    PROPIEDADESLCD.lbldigital(1).ForeColor = FUNCIONES.ColorIluminado
    frmPantalla.lbldigital(0).ForeColor = FUNCIONES.ColorLCDInactivo
    PROPIEDADESLCD.lbldigital(0).ForeColor = FUNCIONES.ColorLCDInactivo
    PROPIEDADESLCD.Shape1.BackColor = FUNCIONES.ColorVentana
    PROPIEDADESLCD.Shape2.BackColor = FUNCIONES.ColorVentana
    PROPIEDADESLCD.Shape1.BorderColor = FUNCIONES.ColorVentana
    PROPIEDADESLCD.Shape2.BorderColor = FUNCIONES.ColorVentana
    PROPIEDADESLCD.cursor.BackColor = FUNCIONES.ColorVentana
    PROPIEDADESLCD.Label1.BackColor = FUNCIONES.ColorVentana
    PROPIEDADESLCD.Label2.BackColor = FUNCIONES.ColorVentana
    PROPIEDADESLCD.LISTPropiedades.ForeColor = FUNCIONES.ColorVentana
    PROPIEDADESLCD.Pic_Nodo.BackColor = FUNCIONES.ColorVentana
    FUNCIONES.LCDContador = CInt(NumeroInicial)
End Sub

