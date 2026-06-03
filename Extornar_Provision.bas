' =============================================================================
' Módulo : Extornar_Provision
' Propósito: Extorna (revierte en bloque) todas las provisiones del mes actual:
'            1. Limpia los campos de provisión en la hoja PlanPagos
'               (columnas AA:AD – Orden, CUE, Informe, Monto sin IGV).
'            2. Cierra el archivo Excel de provisiones si está abierto.
'            3. Elimina dicho archivo de la subcarpeta del año correspondiente.
'
' Ruta del archivo (igual que EjecutarProvision):
'   [Celda B10]\[Año]\Provisiones TIC [Año]-[mmm].xlsx
'
' Instrucciones de uso:
'   1. Abrir el Editor VBA (Alt+F11).
'   2. Insertar módulo nuevo y pegar este código.
'   3. Verificar en la hoja Control:
'      - Celda B2  : Fecha del mes a extornar.
'      - Celda B6  : N° de Informe de la provisión ejecutada.
'      - Celda B7  : N° de Expediente (CUE) de la provisión ejecutada.
'      - Celda B10 : Ruta base donde se guardan los archivos de provisión.
'   4. Asignar acceso directo Ctrl+Mayús+R:
'      Herramientas > Macros > Opciones > Tecla de método abreviado.
' =============================================================================

Sub Extornar_Provision()
'
' Extorna las provisiones del mes actual: limpia PlanPagos y elimina el archivo
' generado por EjecutarProvision en la subcarpeta del año correspondiente.
' Acceso directo: Ctrl+Mayús+R  (asignar manualmente en Herramientas > Macros)
'

    ' ── Hojas ────────────────────────────────────────────────────────────────
    Const CH_Control    = "Control"
    Const CH_Cronograma = "PlanPagos"

    ' ── Celdas en hoja Control (idénticas a EjecutarProvision) ───────────────
    Const CC_xMesP = 2   ' Fila  – Mes a provisionar
    Const CC_yMesP = 2   ' Columna – Mes a provisionar
    Const CC_xExpe = 7   ' Fila  – N° Expediente (CUE)
    Const CC_yExpe = 2   ' Columna – N° Expediente (CUE)
    Const CC_xInfo = 6   ' Fila  – N° Informe
    Const CC_yInfo = 2   ' Columna – N° Informe
    Const CC_xRuta = 10  ' Fila  – Ruta base de archivos de provisión (B10)
    Const CC_yRuta = 2   ' Columna – Ruta base de archivos de provisión

    ' ── Nombre del archivo (igual que EjecutarProvision) ─────────────────────
    Const CT_Archivo = "Provisiones TIC "   ' Prefijo del nombre de archivo

    ' ── Columnas en PlanPagos escritas por EjecutarProvision ─────────────────
    Const CP_ColOrden  = 27  ' AA – N° Orden de provisión
    Const CP_ColExpSIG = 28  ' AB – CUE / Expediente SIGEDD
    Const CP_ColInfSIG = 29  ' AC – N° Informe SIGEDD
    Const CP_ColMonto  = 30  ' AD – Monto provisionado sin IGV

    Const CC_IniCronogra = 4 ' Primera fila de datos en PlanPagos

    ' ── Variables ─────────────────────────────────────────────────────────────
    Dim wsControl   As Worksheet
    Dim wsPlanPagos As Worksheet
    Dim wb          As Workbook

    Dim ExpSIGEDD    As String
    Dim DocSIGEDD    As String
    Dim MesProv      As Date
    Dim mesLetras    As String
    Dim mesAnno      As String
    Dim AnnoProv     As Integer
    Dim RutaBase     As String
    Dim nombreLib    As String
    Dim rutaCompleta As String

    Dim UltimaFil        As Long
    Dim FilaPlanP        As Long
    Dim ContRevertidos   As Long
    Dim cExpFila         As String
    Dim cInfFila         As String
    Dim estadoArchivo    As String

    ' ── Obtener hojas ─────────────────────────────────────────────────────────
    Set wsControl   = ThisWorkbook.Sheets(CH_Control)
    Set wsPlanPagos = ThisWorkbook.Sheets(CH_Cronograma)

    ' ── Leer y validar parámetros desde la hoja Control ──────────────────────
    If IsEmpty(wsControl.Cells(CC_xMesP, CC_yMesP).Value) Or _
       Not IsDate(wsControl.Cells(CC_xMesP, CC_yMesP).Value) Then
        MsgBox "No hay un mes de provisión definido en la hoja Control (celda B2)." & vbNewLine & _
               "Ingrese la fecha del mes a extornar y vuelva a ejecutar.", _
               vbExclamation, "Extorno cancelado"
        Exit Sub
    End If

    MesProv   = CDate(wsControl.Cells(CC_xMesP, CC_yMesP).Value)
    ExpSIGEDD = Trim(CStr(wsControl.Cells(CC_xExpe, CC_yExpe).Value))
    DocSIGEDD = Trim(CStr(wsControl.Cells(CC_xInfo, CC_yInfo).Value))
    RutaBase  = Trim(CStr(wsControl.Cells(CC_xRuta, CC_yRuta).Value))

    AnnoProv  = Year(MesProv)
    mesLetras = Format(MesProv, "mmm")        ' ej: "jun"  (igual que EjecutarProvision)
    mesAnno   = Format(MesProv, "mmmm-yyyy")  ' ej: "junio-2025"  (para mensajes)

    ' Construir ruta y nombre exactamente igual que EjecutarProvision
    nombreLib    = CT_Archivo & AnnoProv & "-" & mesLetras & ".xlsx"
    rutaCompleta = RutaBase & "\" & AnnoProv & "\" & nombreLib

    If ExpSIGEDD = "" Or ExpSIGEDD = "0" Then
        MsgBox "No hay número de Expediente (CUE) definido en la hoja Control (celda B7)." & vbNewLine & _
               "Ingrese el expediente de la provisión a extornar.", _
               vbExclamation, "Extorno cancelado"
        Exit Sub
    End If

    ' ── Confirmar con el usuario ───────────────────────────────────────────────
    Dim respuesta As Integer
    respuesta = MsgBox( _
        "¿Confirma el extorno de las provisiones del mes de " & UCase(mesAnno) & "?" & vbNewLine & vbNewLine & _
        "  Expediente (CUE) : " & ExpSIGEDD & vbNewLine & _
        "  Informe          : " & DocSIGEDD & vbNewLine & vbNewLine & _
        "Se realizarán las siguientes acciones:" & vbNewLine & _
        "  1. Limpiar columnas AA-AD en PlanPagos (Orden, CUE, Informe, Monto)." & vbNewLine & _
        "  2. Cerrar el archivo de provisión si está abierto en Excel." & vbNewLine & _
        "  3. Eliminar el archivo:" & vbNewLine & _
        "     " & rutaCompleta, _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Confirmar Extorno de Provisiones")

    If respuesta = vbNo Then
        MsgBox "Operación cancelada por el usuario.", vbInformation, "Extorno cancelado"
        Exit Sub
    End If

    ' ── Optimizar rendimiento ─────────────────────────────────────────────────
    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual
    Application.EnableEvents   = False

    ' ── 1. Cerrar el archivo de provisión si está abierto en Excel ───────────
    estadoArchivo = "no abierto"
    On Error Resume Next
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, nombreLib, vbTextCompare) = 0 Then
            wb.Close SaveChanges:=False
            estadoArchivo = "cerrado"
            Exit For
        End If
    Next wb
    On Error GoTo 0

    ' ── 2. Eliminar el archivo del disco si existe ────────────────────────────
    If Dir(rutaCompleta) <> "" Then
        On Error Resume Next
        Kill rutaCompleta
        If Err.Number <> 0 Then
            Application.ScreenUpdating = True
            Application.Calculation    = xlCalculationAutomatic
            Application.EnableEvents   = True
            MsgBox "No se pudo eliminar el archivo:" & vbNewLine & rutaCompleta & vbNewLine & vbNewLine & _
                   "Error: " & Err.Description & vbNewLine & vbNewLine & _
                   "Verifique que el archivo no esté bloqueado y vuelva a intentarlo.", _
                   vbCritical, "Error al eliminar archivo"
            Exit Sub
        End If
        On Error GoTo 0
        estadoArchivo = estadoArchivo & " y eliminado"
    Else
        estadoArchivo = estadoArchivo & " | archivo no encontrado en disco"
    End If

    ' ── 3. Limpiar columnas de provisión en PlanPagos ────────────────────────
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
        MsgBox "Extorno completado exitosamente." & vbNewLine & vbNewLine & _
               "  Mes              : " & UCase(mesAnno) & vbNewLine & _
               "  Expediente (CUE) : " & ExpSIGEDD & vbNewLine & _
               "  Informe          : " & DocSIGEDD & vbNewLine & vbNewLine & _
               "  Filas extornadas en PlanPagos : " & ContRevertidos & vbNewLine & _
               "  Archivo           : " & estadoArchivo & vbNewLine & _
               "  Ruta              : " & rutaCompleta & vbNewLine & vbNewLine & _
               "Las filas están listas para volver a ser provisionadas.", _
               vbInformation, "Extorno exitoso"
    Else
        MsgBox "No se encontraron provisiones en PlanPagos para:" & vbNewLine & vbNewLine & _
               "  Expediente (CUE) : " & ExpSIGEDD & vbNewLine & _
               "  Informe          : " & DocSIGEDD & vbNewLine & vbNewLine & _
               "  Archivo : " & estadoArchivo & vbNewLine & _
               "  Ruta    : " & rutaCompleta & vbNewLine & vbNewLine & _
               "Verifique que los datos en la hoja Control correspondan" & vbNewLine & _
               "a una provisión ya ejecutada.", _
               vbExclamation, "Extorno sin resultados en PlanPagos"
    End If

End Sub
