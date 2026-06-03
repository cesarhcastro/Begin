' =============================================================================
' Módulo : RevertirProvisionMes
' Propósito: Revierte (extorna en bloque) todas las provisiones del mes actual
'            registradas en la hoja PlanPagos por la macro EjecutarProvision.
'
' Funcionamiento:
'   Lee el Expediente (CUE) e Informe definidos en la hoja Control para el
'   mes en curso, luego recorre la hoja PlanPagos y limpia las columnas
'   AA:AD (Orden, CUE, Informe, Monto) en cada fila donde coincidan ambos
'   valores, dejando esas filas listas para ser provisionadas nuevamente.
'
' Instrucciones de uso:
'   1. Abrir el Editor VBA (Alt+F11).
'   2. Insertar módulo nuevo y pegar este código (o usar Archivo > Importar).
'   3. Asegurarse de que en la hoja Control estén correctamente cargados:
'      - Celda B2  : Fecha del mes a revertir.
'      - Celda B6  : N° de Informe de la provisión ejecutada.
'      - Celda B7  : N° de Expediente (CUE) de la provisión ejecutada.
'   4. Asignar acceso directo Ctrl+Mayús+R desde:
'      Herramientas > Macros > Opciones > Tecla de método abreviado.
' =============================================================================

Sub RevertirProvisionMes()
'
' Revierte todas las provisiones del mes actual en PlanPagos
' (limpia columnas AA, AB, AC, AD: Orden, CUE, Informe y Monto sin IGV).
' Acceso directo: Ctrl+Mayús+R  (asignar manualmente en Herramientas > Macros)
'

    ' ── Hojas ────────────────────────────────────────────────────────────────
    Const CH_Control    = "Control"
    Const CH_Cronograma = "PlanPagos"

    ' ── Celdas en hoja Control (idénticas a EjecutarProvision) ───────────────
    Const CC_xMesP = 2  ' Fila  del mes a provisionar
    Const CC_yMesP = 2  ' Columna del mes a provisionar
    Const CC_xExpe = 7  ' Fila  del N° de Expediente (CUE)
    Const CC_yExpe = 2  ' Columna del N° de Expediente (CUE)
    Const CC_xInfo = 6  ' Fila  del N° de Informe
    Const CC_yInfo = 2  ' Columna del N° de Informe

    ' ── Columnas en PlanPagos escritas por EjecutarProvision ─────────────────
    Const CP_ColOrden  = 27  ' AA – N° Orden de provisión
    Const CP_ColExpSIG = 28  ' AB – CUE / Expediente SIGEDD
    Const CP_ColInfSIG = 29  ' AC – N° Informe SIGEDD
    Const CP_ColMonto  = 30  ' AD – Monto provisionado sin IGV

    Const CC_IniCronogra = 4 ' Primera fila de datos en PlanPagos

    ' ── Variables ─────────────────────────────────────────────────────────────
    Dim wsControl   As Worksheet
    Dim wsPlanPagos As Worksheet

    Dim ExpSIGEDD   As String
    Dim DocSIGEDD   As String
    Dim MesProv     As Date
    Dim mesLetras   As String

    Dim UltimaFil      As Long
    Dim FilaPlanP      As Long
    Dim ContRevertidos As Long
    Dim cExpFila       As String
    Dim cInfFila       As String

    ' ── Obtener hojas ─────────────────────────────────────────────────────────
    Set wsControl   = ThisWorkbook.Sheets(CH_Control)
    Set wsPlanPagos = ThisWorkbook.Sheets(CH_Cronograma)

    ' ── Leer parámetros del mes desde la hoja Control ─────────────────────────
    If IsEmpty(wsControl.Cells(CC_xMesP, CC_yMesP).Value) Or _
       Not IsDate(wsControl.Cells(CC_xMesP, CC_yMesP).Value) Then
        MsgBox "No hay un mes de provisión definido en la hoja Control (celda B2)." & vbNewLine & _
               "Ingrese la fecha del mes a revertir y vuelva a ejecutar.", _
               vbExclamation, "Reversión cancelada"
        Exit Sub
    End If

    MesProv   = CDate(wsControl.Cells(CC_xMesP, CC_yMesP).Value)
    ExpSIGEDD = Trim(CStr(wsControl.Cells(CC_xExpe, CC_yExpe).Value))
    DocSIGEDD = Trim(CStr(wsControl.Cells(CC_xInfo, CC_yInfo).Value))
    mesLetras = Format(MesProv, "mmmm-yyyy")

    If ExpSIGEDD = "" Or ExpSIGEDD = "0" Then
        MsgBox "No hay número de Expediente (CUE) definido en la hoja Control (celda B7)." & vbNewLine & _
               "Ingrese el expediente de la provisión a revertir.", _
               vbExclamation, "Reversión cancelada"
        Exit Sub
    End If

    ' ── Confirmación del usuario ───────────────────────────────────────────────
    Dim respuesta As Integer
    respuesta = MsgBox( _
        "¿Confirma la reversión de las provisiones del mes de " & UCase(mesLetras) & "?" & vbNewLine & vbNewLine & _
        "  Expediente (CUE) : " & ExpSIGEDD & vbNewLine & _
        "  Informe          : " & DocSIGEDD & vbNewLine & vbNewLine & _
        "Se limpiarán los campos Orden, CUE, Informe y Monto (columnas AA-AD)" & vbNewLine & _
        "de todas las filas coincidentes en la hoja PlanPagos." & vbNewLine & vbNewLine & _
        "Esta acción permite volver a ejecutar la provisión para el mismo mes.", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Confirmar Reversión de Provisiones")

    If respuesta = vbNo Then
        MsgBox "Operación cancelada por el usuario.", vbInformation, "Reversión cancelada"
        Exit Sub
    End If

    ' ── Optimizar rendimiento ─────────────────────────────────────────────────
    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual
    Application.EnableEvents   = False

    ' ── Recorrer PlanPagos y limpiar filas que coincidan ─────────────────────
    UltimaFil      = wsPlanPagos.Cells(wsPlanPagos.Rows.Count, 3).End(xlUp).Row
    ContRevertidos = 0

    For FilaPlanP = CC_IniCronogra To UltimaFil

        cExpFila = Trim(CStr(wsPlanPagos.Cells(FilaPlanP, CP_ColExpSIG).Value))
        cInfFila = Trim(CStr(wsPlanPagos.Cells(FilaPlanP, CP_ColInfSIG).Value))

        If cExpFila = ExpSIGEDD And cInfFila = DocSIGEDD Then
            wsPlanPagos.Cells(FilaPlanP, CP_ColOrden).ClearContents   ' AA – N° Orden
            wsPlanPagos.Cells(FilaPlanP, CP_ColExpSIG).ClearContents  ' AB – CUE
            wsPlanPagos.Cells(FilaPlanP, CP_ColInfSIG).ClearContents  ' AC – Informe
            wsPlanPagos.Cells(FilaPlanP, CP_ColMonto).ClearContents   ' AD – Monto sin IGV
            ContRevertidos = ContRevertidos + 1
        End If

    Next FilaPlanP

    ' ── Restaurar configuración ───────────────────────────────────────────────
    Application.ScreenUpdating = True
    Application.Calculation    = xlCalculationAutomatic
    Application.EnableEvents   = True

    ' ── Informar resultado ────────────────────────────────────────────────────
    If ContRevertidos > 0 Then
        MsgBox "Reversión completada exitosamente." & vbNewLine & vbNewLine & _
               "  Mes         : " & UCase(mesLetras) & vbNewLine & _
               "  Expediente  : " & ExpSIGEDD & vbNewLine & _
               "  Informe     : " & DocSIGEDD & vbNewLine & vbNewLine & _
               "  Provisiones revertidas: " & ContRevertidos & vbNewLine & vbNewLine & _
               "Las filas están listas para volver a ser provisionadas.", _
               vbInformation, "Reversión exitosa"
    Else
        MsgBox "No se encontraron provisiones registradas para los parámetros indicados:" & vbNewLine & vbNewLine & _
               "  Mes         : " & mesLetras & vbNewLine & _
               "  Expediente  : " & ExpSIGEDD & vbNewLine & _
               "  Informe     : " & DocSIGEDD & vbNewLine & vbNewLine & _
               "Verifique que los datos en la hoja Control correspondan" & vbNewLine & _
               "a una provisión que ya fue ejecutada.", _
               vbExclamation, "Sin resultados"
    End If

End Sub
