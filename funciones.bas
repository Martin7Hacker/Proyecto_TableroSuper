Attribute VB_Name = "FUNCIONES"
'---------------------------------------------------------------------
' ARRANCAR CON WINDOWS S.O
'---------------------------------------------------------------------
'< MartinSoft@Hotmail.es >
'
'Constantes de la Rama del registro para los path de _
 las aplicaciones que inician con Windows
Const RAMA_RUN_WINDOWS As String = "SOFTWARE\Microsoft\" & _
                                   "Windows\CurrentVersion\Run"
Public VectorTexto(15) As String
Public LCDContador As String
Public nodoEstado As String
Public control As String
Public TecIncremento As String
Public TecDescremento As String
Public tamanioDisplay As String
Public ColorIluminado As String
Public ColorLCDInactivo As String
Public ColorVentana As String
Public NumeroInicial As String
Public Size As String
Public top As String

Public Function visualizarLCD(ByVal activo As Boolean)
With frmPantalla
.lbldigital(0).Visible = activo
.lbldigital(1).Visible = activo
.picMover.Visible = activo
End With

End Function

Public Function visualizarLCDPropiedades(ByVal activo As Boolean)
With PROPIEDADESLCD
.lbldigital(0).Visible = activo
.lbldigital(1).Visible = activo
.picMover.Visible = activo
End With
End Function

Public Function visualizarNODO(ByVal activo As Boolean)
With frmPantalla
.picMover.Visible = activo
End With
End Function

Public Function DispositoEntrada(ByVal dispositivo As Byte) As String
        Select Case dispositivo
        Case 0
        DispositoEntrada = ""
        Case 1
        DispositoEntrada = "El Dispoivio de entrada es:Ratón"
        Case 2
        DispositoEntrada = "El Dispoivio de entrada es:Teclado"
    End Select
     PROPIEDADESLCD.lblDispositivo = DispositoEntrada
End Function

Public Function nodo(ByVal nodoX As Boolean) As Boolean
    PROPIEDADESLCD.picMover.Visible = nodoX
End Function

Public Function pintarVentana()
With PROPIEDADESLCD.cd
    .ShowColor
    PROPIEDADESLCD.Shape1.BackColor = .Color
    PROPIEDADESLCD.Shape2.BackColor = .Color
    PROPIEDADESLCD.Shape1.BorderColor = .Color
    PROPIEDADESLCD.Shape2.BorderColor = .Color
    PROPIEDADESLCD.cursor.BackColor = .Color
    PROPIEDADESLCD.Label1.BackColor = .Color
    PROPIEDADESLCD.Label2.BackColor = .Color
    PROPIEDADESLCD.LISTPropiedades.ForeColor = .Color
    FUNCIONES.ColorVentana = .Color
    PROPIEDADESLCD.Pic_Nodo.BackColor = .Color
End With
End Function
Public Function Display(ByVal control As Boolean) As Boolean
    With PROPIEDADESLCD
        .lbldigital(0).Visible = control
        .lbldigital(1).Visible = control
    End With
End Function

Public Function Color_del_LCD_Iluminado()
    With PROPIEDADESLCD.cd
        .ShowColor
         frmPantalla.lbldigital(1).ForeColor = .Color
         frmPantalla.lbltime.ForeColor = .Color
         FUNCIONES.ColorIluminado = .Color
         PROPIEDADESLCD.lbldigital(1).ForeColor = .Color
End With
End Function

Public Function Color_LCD_Inactivo()
    With PROPIEDADESLCD.cd
        .ShowColor
        frmPantalla.lbldigital(0).ForeColor = .Color
        PROPIEDADESLCD.lbldigital(0).ForeColor = .Color
        FUNCIONES.ColorLCDInactivo = .Color
    End With
End Function

Public Sub reproducir()
 With frmPantalla.rep
    .URL = "LCD.wav"
    End With
End Sub
