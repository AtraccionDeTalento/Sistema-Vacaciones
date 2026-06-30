Attribute VB_Name = "LanzadorVacaciones"
'==============================================================================
' Lanzador VBA del Pipeline de Vacaciones USIL  (opcion "boton dentro de Excel")
'
' COMO USARLO:
'   1. Abre Excel > Alt+F11 (editor VBA).
'   2. Archivo > Importar archivo... > elige este  Lanzador_Excel.bas
'      (puedes importarlo en PERSONAL.XLSB para tenerlo en todos los libros,
'       o en un libro .xlsm dedicado tipo "Lanzador.xlsm").
'   3. Vuelve a Excel > pestania Programador > Insertar > Boton, o una forma,
'      y asigna la macro  ActualizarVacaciones.
'   4. Click en el boton = corre el mismo proceso que ACTUALIZAR_VACACIONES.bat
'
' Si mueves la carpeta PIPELINE, actualiza la constante RUTA_BAT de abajo.
'==============================================================================
Option Explicit

Private Const RUTA_BAT As String = _
    "C:\Users\jlopezp\OneDrive - Universidad San Ignacio de Loyola\ACTIVIDADES\CESAR\SISTEMA DE VACACIONES\PIPELINE\ACTUALIZAR_VACACIONES.bat"

Public Sub ActualizarVacaciones()
    If Dir(RUTA_BAT) = "" Then
        MsgBox "No se encontro el lanzador:" & vbCrLf & RUTA_BAT & vbCrLf & vbCrLf & _
               "Edita la constante RUTA_BAT en el modulo VBA.", vbCritical, "Vacaciones USIL"
        Exit Sub
    End If

    If MsgBox("Se actualizara el Sistema de Vacaciones tomando el ultimo" & vbCrLf & _
              "VACRptMotivo_*.xlsx de la carpeta Descargas." & vbCrLf & vbCrLf & _
              "Se creara una COPIA FECHADA en la carpeta SALIDAS" & vbCrLf & _
              "(el archivo original no se toca)." & vbCrLf & vbCrLf & _
              "Continuar?", vbQuestion + vbYesNo, "Vacaciones USIL") = vbNo Then Exit Sub

    ' cmd /k deja la ventana abierta para ver el resultado; usa /c si la prefieres auto-cerrable
    Shell "cmd /k """ & RUTA_BAT & """", vbNormalFocus
End Sub
