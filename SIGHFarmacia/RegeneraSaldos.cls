VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegeneraSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para Regenerar Saldos, Actualizar Precios
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Formulario As New HerrRegeneraSaldos
Dim mrs_Tmp As New ADODB.Recordset
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_idUsuario As Long
Dim mo_lcNombrePc  As String
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let idUsuario(lIdValue As Long)
    ml_idUsuario = lIdValue
End Property

Sub MostrarFormulario()
    mo_Formulario.lcNombrePc = mo_lcNombrePc
    mo_Formulario.idUsuario = ml_idUsuario
    mo_Formulario.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    mo_Formulario.Show 1
    Set mo_Formulario = Nothing
End Sub

'Sub CrearReporte_excel(lbUsandoProcesoEnServidor As Boolean, lnHwnd As Long)
'Dim rsReporte As New Recordset
'Dim rsReporte1 As New Recordset
'Dim iFila As Long
'Dim lnTotal As Double
'Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
'Dim mo_ReporteUtil As New sighentidades.ReporteUtil
'Dim oConexion As New ADODB.Connection
'Dim lcSql As String, lcHoraInicio As String, lcHoraFinal As String
'Dim lbEsOpenOffice As Boolean
'
'lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
'    On Error GoTo ManejadorErrorExcel
'
'    If lbEsOpenOffice = True Then
'        Dim ServiceManager As Object
'        Dim Desktop As Object
'        Dim Document As Object
'        Dim Feuille As Object
'        Dim Plage As Object
'        Dim args()
'        Dim Chemin As String
'        Dim Fichier As String
'        Dim lcArchivoExcel As String
'        Dim PrintArea(0)
'        Dim Style As Object
'        Dim Border As Object
'        'encabezado
'        Dim PageStyles As Object
'        Dim Sheet As Object
'        Dim StyleFamilies As Object
'        Dim DefPage As Object
'        Dim Htext As Object
'        Dim Hcontent As Object
'        Dim ret As Long
'    Else
'        Dim oExcel As Excel.Application
'        Dim oWorkBookPlantilla As Workbook
'        Dim oWorkBook As Workbook
'        Dim oWorkSheet As Worksheet
'    End If
'    oConexion.Open sighentidades.CadenaConexion
'    oConexion.CursorLocation = adUseClient
'    GenerarRecordsetTemporal
'    '***************Error- Cabecera sin  Detalle
'    lcHoraInicio = lcBuscaParametro.RetornaHoraServidorSQL1
'    '**********Errores de Stock con Cantidades negativas
'    Set rsReporte = mo_ReglasFarmacia.SaldosNegativos
'    If rsReporte.RecordCount > 0 Then
'        rsReporte.MoveFirst
'        Do While Not rsReporte.EOF
'            mrs_Tmp.AddNew
'            mrs_Tmp.Fields("codigo").Value = "Existe LOTE con Saldos NEGATIVOS"
'            mrs_Tmp.Fields!descrip = "Almacen: " & Trim(rsReporte.Fields!descripcion) & "    Producto: " & rsReporte.Fields!codigo & " " & Trim(rsReporte.Fields!Nombre) & "    Cantidad: " & rsReporte.Fields!cantidad
'            mrs_Tmp.Fields!solucion = "Chequee los Movimientos de E/S (KARDEX), debe faltar algun Ingreso"
'            mrs_Tmp.Update
'            rsReporte.MoveNext
'        Loop
'    End If
'    rsReporte.Close
'
'        Set rsReporte = FarmNiNsConCabeceraSinDetalle(oConexion)
'        rsReporte.Filter = "cantidad=0"
'        If rsReporte.RecordCount > 0 Then
'            rsReporte.MoveFirst
'            Do While Not rsReporte.EOF
'                  mrs_Tmp.AddNew
'                  mrs_Tmp.Fields!codigo = "Cabecera sin Detalle"
'                  mrs_Tmp.Fields!descrip = IIf(rsReporte.Fields!MovTipo = "E", "Nota Ingreso: ", "Nota Salida: ") & rsReporte.Fields!movNumero & "    Fecha: " & rsReporte.Fields!FechaCreacion
'                  mrs_Tmp.Fields!solucion = "Fijese si el Dcto existe fisicamente, Anulelo o  Registre los Productos"
'                  mrs_Tmp.Update
'               rsReporte.MoveNext
'            Loop
'        End If
'        rsReporte.Close
''
'    lcHoraFinal = lcBuscaParametro.RetornaHoraServidorSQL1
'    If mrs_Tmp.RecordCount > 0 Then
'        If lbEsOpenOffice = True Then
'            'Abre el archivo ExcelOpenOffice
'            lcArchivoExcel = App.Path + "\Plantillas\farmErrores.ods"
''            FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
''            Chemin = "file:///" & App.Path & "\Plantillas\"
''            Chemin = Replace(Chemin, "\", "/")
''            Fichier = Chemin & "/OpenOffice.ods"
'            Fichier = Format(Time, "hhmmss") & ".ods"
'            FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
'            lcArchivoExcel = Fichier
'            Chemin = "file:///" & App.Path & "\Plantillas\"
'            Chemin = Replace(Chemin, "\", "/")
'            Fichier = Chemin & "/" & lcArchivoExcel
'            '
'            Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
'            Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
'            Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
'            Set Feuille = Document.getSheets().getByIndex(0)
'            'Encabezado de Pagina
'            mo_CabeceraReportes.CabeceraReportes Document, True
'            ' Pone la ventana en primer plano, pasándole el Hwnd
'            ret = SetForegroundWindow(lnHwnd)
'        Else
'            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
'            'Crea nueva hoja
'            Set oWorkBook = oExcel.Workbooks.Add
'            'Abre, copia y cierra la plantilla
'            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\farmErrores.xls")
'            oWorkBookPlantilla.Worksheets("farmErrores").Copy Before:=oWorkBook.Sheets(1)
'            oWorkBookPlantilla.Close
'            'Activa la primera hoja
'            Set oWorkSheet = oWorkBook.Sheets(1)
'            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
'        End If
'        iFila = 6: lnTotal = 0
'        mrs_Tmp.MoveFirst
'        Do While Not mrs_Tmp.EOF
'            If lbEsOpenOffice = True Then
'                Call Feuille.getcellbyposition(1, iFila - 1).setFormula(mrs_Tmp.Fields("codigo").Value)
'                Call Feuille.getcellbyposition(2, iFila - 1).setFormula(mrs_Tmp.Fields("descrip").Value)
'                Call Feuille.getcellbyposition(3, iFila - 1).setFormula(mrs_Tmp.Fields("solucion").Value)
'            Else
'                oWorkSheet.Cells(iFila, 2).Value = mrs_Tmp.Fields("codigo").Value
'                oWorkSheet.Cells(iFila, 3).Value = mrs_Tmp.Fields("descrip").Value
'                oWorkSheet.Cells(iFila, 4).Value = mrs_Tmp.Fields("solucion").Value
'            End If
'           iFila = iFila + 1
'           mrs_Tmp.MoveNext
'        Loop
'        If lbEsOpenOffice = True Then
'            Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":D" & CStr(iFila))
'            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
'            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Nº Errores: " + Trim(Str(mrs_Tmp.RecordCount)))
'        Else
'            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 4
'            oWorkSheet.Cells(iFila, 2).Value = "Nº Errores: " + Trim(Str(mrs_Tmp.RecordCount))
'        End If
'        If lbEsOpenOffice = True Then
'            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
'            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
'        Else
'            oExcel.Visible = True
'            oWorkSheet.PrintPreview
'            'oWorkSheet.PrintOut
'        End If
'    End If
'    If lbEsOpenOffice = True Then
'        'Liberar Memoria
'        Set Plage = Nothing
'        Set Feuille = Nothing
'        Set Document = Nothing
'        Set Desktop = Nothing
'        Set ServiceManager = Nothing
'        Set Style = Nothing
'        Set Border = Nothing
'        'encabezado de pagina
'        Set PageStyles = Nothing
'        Set Sheet = Nothing
'        Set StyleFamilies = Nothing
'        Set DefPage = Nothing
'        Set Htext = Nothing
'        Set Hcontent = Nothing
'    Else
'        'Liberar memoria
'        Set oExcel = Nothing
'        Set oWorkBookPlantilla = Nothing
'        Set oWorkBook = Nothing
'        Set oWorkSheet = Nothing
'    End If
'Exit Sub
'ManejadorErrorExcel:
'    Select Case Err.Number
'    Case 1004
'        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
'    Case Else
'        MsgBox Err.Description
'    End Select
'    Exit Sub
'Resume
'End Sub


'Sub GenerarRecordsetTemporal()
'    With mrs_Tmp
'          .Fields.Append "codigo", adVarChar, 50, adFldIsNullable
'          .Fields.Append "Descrip", adVarChar, 200, adFldIsNullable
'          .Fields.Append "Solucion", adVarChar, 200, adFldIsNullable
'          .LockType = adLockOptimistic
'          .Open
'    End With
'End Sub

'Function FarmNiNsConCabeceraSinDetalle(mo_Conexion As Connection) As ADODB.Recordset
'On Error GoTo ManejadorDeError
'Dim oRecordset As New ADODB.Recordset
'Dim oCommand As New ADODB.Command
'Dim oParameter As ADODB.Parameter
'Dim ms_MensajeError As String
'Dim oConexion As New ADODB.Connection
'    oConexion.Open sighentidades.CadenaConexion
'    oConexion.CursorLocation = adUseClient
'    Set FarmNiNsConCabeceraSinDetalle = Nothing
'    ms_MensajeError = ""
'    With oCommand
'        .CommandType = adCmdStoredProc
'        Set .ActiveConnection = oConexion
'        .CommandTimeout = 150
'        .CommandText = "FarmNiNsConCabeceraSinDetalle"
'        Set oRecordset = .Execute
'   End With
'   Set FarmNiNsConCabeceraSinDetalle = oRecordset
'Exit Function
'ManejadorDeError:
'   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
'Exit Function
'End Function

