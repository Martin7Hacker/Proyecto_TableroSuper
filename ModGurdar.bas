Attribute VB_Name = "ModGurdarAbrir"
'------------------------------------------
'Módulo de Código para guardar en Archivo -
'------------------------------------------

Option Explicit

Public Sub gardarArchio()
Dim recX As Byte
    Open "DisplayLCD.dbx" For Output As 1
    Size = frmPantalla.lbldigital(1).Left
    top = frmPantalla.lbldigital(1).top
    Print #1, LCDContador
    Print #1, nodoEstado
    Print #1, control
    Print #1, TecIncremento
    Print #1, TecDescremento
    Print #1, tamanioDisplay
    Print #1, ColorIluminado
    Print #1, ColorLCDInactivo
    Print #1, ColorVentana
    Print #1, NumeroInicial
    Print #1, CInt(Size)
    Print #1, CInt(top)
    
    For recX = 0 To 15
        Print #1, FUNCIONES.VectorTexto(recX)
    Next recX
 
  Close #1
End Sub

'-----------------------------------
'- Abrir Archivo                   -
'-----------------------------------

Public Sub AbrirArchivo()

Dim recX As Byte

On Error GoTo no_se

Open "DisplayLCD.dbx" For Input As 1

Do While Not EOF(1)

 Line Input #1, LCDContador
 Line Input #1, nodoEstado
 Line Input #1, control
 Line Input #1, TecIncremento
 Line Input #1, TecDescremento
 Line Input #1, tamanioDisplay
 Line Input #1, ColorIluminado
 Line Input #1, ColorLCDInactivo
 Line Input #1, ColorVentana
 Line Input #1, NumeroInicial
 Line Input #1, Size
 Line Input #1, top

 For recX = 0 To 15
    Line Input #1, FUNCIONES.VectorTexto(recX)
 Next recX
  Loop
 
 
 Close #1

 
no_se:
End Sub



