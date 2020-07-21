VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form propiedadesLCD 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Propiedades del LCD"
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   Icon            =   "propiedadesLCD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   3960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   3960
   End
   Begin VB.ListBox LISTPropiedades 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4830
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   4440
      ScaleHeight     =   4815
      ScaleWidth      =   4455
      TabIndex        =   2
      Top             =   360
      Width           =   4455
      Begin VB.TextBox txtDisplay 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "371"
         Top             =   480
         Width           =   4095
      End
      Begin VB.PictureBox Pic_Nodo 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   2895
         TabIndex        =   7
         Top             =   3480
         Width           =   2895
         Begin VB.Label Labnodo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SI"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   8
            Top             =   45
            Width           =   2730
         End
      End
      Begin VB.PictureBox picMover 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   67.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   3
         Top             =   1200
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
      Begin VB.Label lblDispositivo 
         BackStyle       =   0  'Transparent
         Caption         =   "El Dispoivio de entrada es:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   4080
         Width           =   4215
      End
      Begin VB.Label lblCambio 
         BackColor       =   &H0000FF00&
         Caption         =   "Los Cambios se Aplicaron con Exito"
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
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Label lbldigital 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "LCD"
            Size            =   69.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1380
         Index           =   1
         Left            =   1440
         TabIndex        =   4
         Top             =   1680
         Width           =   1410
      End
      Begin VB.Label lbldigital 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "LCD"
            Size            =   69.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1380
         Index           =   0
         Left            =   1440
         TabIndex        =   5
         Top             =   1680
         Width           =   1410
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Configuración del LCD"
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
      Left            =   4440
      TabIndex        =   6
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Propiedades LCD"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   7
      Height          =   5295
      Left            =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   20
      Height          =   5295
      Left            =   4440
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "PROPIEDADESLCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

                                   
Private Sub listado()
Dim recX As Byte
    LISTPropiedades.Clear
    With LISTPropiedades
    
     For recX = 0 To 10
        .AddItem FUNCIONES.VectorTexto(recX)
     Next recX
    
    End With
End Sub

Private Sub Form_Load()
listado
FUNCIONES.visualizarLCDPropiedades False
LISTPropiedades.ListIndex = 0
Lista
End Sub
Private Sub Form_Unload(Cancel As Integer)
FUNCIONES.visualizarLCD True
FUNCIONES.visualizarNODO FUNCIONES.nodoEstado
End Sub

Private Sub Labnodo_Click()
boton
ModGurdarAbrir.gardarArchio
lblCambio.Visible = True
End Sub

Private Sub LISTPropiedades_Click()
Lista
End Sub

Private Sub Pic_Nodo_Click()
Labnodo_Click
End Sub

Private Sub Pic_Nodo1_Click()
FUNCIONES.visualizarNODO False: picMover.Visible = False
lblCambio.Visible = True: FUNCIONES.nodoEstado = False
End Sub

Private Sub txtDisplay_Change()
    lbldigital(0).Caption = txtDisplay.Text
    lbldigital(1).Caption = txtDisplay.Text
End Sub




Private Sub txtDisplay_KeyPress(KeyAscii As Integer)

If (KeyAscii >= 97) And (KeyAscii < 122) Or (KeyAscii >= 65) And (KeyAscii < 90) Then
 KeyAscii = 8
  End If
If control = 6 Then
txtDisplay.Text = KeyAscii

ElseIf control = 7 Then
txtDisplay.Text = KeyAscii
End If
FUNCIONES.DispositoEntrada 2

End Sub

Private Sub Timer1_Timer()
lblCambio.Visible = False
End Sub


Private Sub txtDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If control = 6 Then
txtDisplay.Text = Button
ElseIf control = 7 Then
txtDisplay.Text = Button
End If
FUNCIONES.DispositoEntrada 1
End Sub

Private Sub Lista()
control = LISTPropiedades.ListIndex
Select Case LISTPropiedades.ListIndex

Case (0) 'Tamaño del display LCD

    txtDisplay.Visible = True
    txtDisplay.Text = CInt(frmPantalla.lbldigital(0).Font.Size)
    Pic_Nodo.Visible = True
    txtDisplay.MaxLength = 3
    Labnodo.Caption = FUNCIONES.VectorTexto(11)
    
Case (1) 'Color del LCD Iluminado

FUNCIONES.Display False
txtDisplay.Visible = False
Pic_Nodo.Visible = True
Labnodo.Caption = FUNCIONES.VectorTexto(12)


Case (2) 'Color del LCD Inactivo

txtDisplay.Visible = False
Pic_Nodo.Visible = True
Labnodo.Caption = FUNCIONES.VectorTexto(12)


Case (3) 'Color de la Ventana

Pic_Nodo.Visible = True
Labnodo.Caption = FUNCIONES.VectorTexto(12)
FUNCIONES.nodo False


Case (4) 'Ocultar nodo del Display LCD

FUNCIONES.nodo frmPantalla.picMover.Visible
FUNCIONES.Display False
txtDisplay.Visible = False
Pic_Nodo.Visible = True
Labnodo.Caption = FUNCIONES.VectorTexto(11)

Case (5) 'Visualizar nodo del Display LCD

FUNCIONES.nodo False
FUNCIONES.Display False
txtDisplay.Visible = False
Pic_Nodo.Visible = True
Labnodo.Caption = FUNCIONES.VectorTexto(11)

Case (6) 'Tecla de Incremento

FUNCIONES.nodo False
FUNCIONES.Display False
txtDisplay.Visible = True
Pic_Nodo.Visible = True
FUNCIONES.Display False
txtDisplay.MaxLength = 3
txtDisplay.Text = FUNCIONES.TecIncremento
Labnodo.Caption = FUNCIONES.VectorTexto(11)

Case (7) 'Tecla de Incremento

FUNCIONES.nodo False
FUNCIONES.Display False
txtDisplay.Visible = True
Pic_Nodo.Visible = True
txtDisplay.MaxLength = 3
txtDisplay.Text = FUNCIONES.TecDescremento
Labnodo.Caption = FUNCIONES.VectorTexto(11)

Case (8) 'Numero Inicial en el Display LCD

    txtDisplay.MaxLength = 2
    FUNCIONES.Display True
    lbldigital(0).Caption = txtDisplay.Text
    lbldigital(1).Caption = txtDisplay.Text
    txtDisplay.Text = FUNCIONES.NumeroInicial
    FUNCIONES.NumeroInicial = txtDisplay.Text
    Labnodo.Caption = FUNCIONES.VectorTexto(11)
    txtDisplay.Text = FUNCIONES.LCDContador - 1
    txtDisplay.Visible = True
    
Case (9) 'Cerrar Ventana Propiedades
    
    FUNCIONES.Display False
    txtDisplay.Visible = False
    Labnodo.Caption = FUNCIONES.VectorTexto(13)
    
 Case (10)
 
    FUNCIONES.Display False
    txtDisplay.Visible = False
    Labnodo.Caption = FUNCIONES.VectorTexto(14)
    
 End Select
End Sub



Private Sub boton()

Select Case LISTPropiedades.ListIndex

    Case (0) 'Tamaño del display LCD
    
        frmPantalla.lbldigital(0).Font.Size = CInt(txtDisplay.Text)
        frmPantalla.lbldigital(1).Font.Size = CInt(txtDisplay.Text)
        FUNCIONES.tamanioDisplay = CInt(txtDisplay.Text) - 1
        lblCambio.Visible = True
        
    Case (1) 'Color del LCD Iluminado
    
    FUNCIONES.Color_del_LCD_Iluminado
    
    Case (2) 'Color del LCD Inactivo
    
    FUNCIONES.Color_LCD_Inactivo
    
    Case (3) 'Color de la Ventana
    
    FUNCIONES.pintarVentana
    
    Case (4) 'Ocultar nodo del Display LCD
    
    FUNCIONES.visualizarNODO False: picMover.Visible = False
    lblCambio.Visible = True: FUNCIONES.nodoEstado = False
    
    Case (5) 'Visualizar nodo del Display LCD
    
    FUNCIONES.visualizarNODO False: picMover.Visible = True
    lblCambio.Visible = True: FUNCIONES.nodoEstado = True
    
    Case (6) 'Tecla de Incremento
    
    lblCambio.Visible = True
    FUNCIONES.TecIncremento = txtDisplay.Text

    Case (7) 'Tecla de Incremento
    
    lblCambio.Visible = True
    FUNCIONES.TecDescremento = txtDisplay.Text
    Case (8) 'Numero Inicial en el Display LCD
     NumeroInicial = txtDisplay.Text
     ModGurdarAbrir.gardarArchio
    Case (9) 'Cerrar Ventana Propiedades
    Unload Me
    Case (10) 'Cerrar Tablero
    End
End Select
End Sub


