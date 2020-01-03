VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clRptHisDxOmitidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Diagnósticos Omitidos (HIS)
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
'*************** Franklin Cachay **************
Option Explicit
Dim mrs_Tmp As New ADODB.Recordset
Dim ml_Mes As Integer
Dim ml_Anio As Integer
Dim mc_Texto As String
Dim mo_Conexion As ADODB.Connection
Dim mo_ReporteUtil As New ReporteUtil
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let Mes(lValue As Integer)
    ml_Mes = lValue
End Property
Property Let Anio(lValue As Integer)
    ml_Anio = lValue
End Property
Property Let Texto(lValue As String)
    mc_Texto = lValue
End Property
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New RptHISDxOmitidos
    oFormulario.Show 1
End Sub

Sub CrearReporte_excel(lnHwnd As Long)
Dim rsReporte As New Recordset
Dim iFila As Long
Dim lnNumTotal As Long
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String
Dim lnCantidadAcumulada As Integer

lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
On Error GoTo ManejadorErrorExcel

    If lbEsOpenOffice = True Then
        Dim ServiceManager As Object
        Dim Desktop As Object
        Dim Document As Object
        Dim Feuille As Object
        Dim Plage As Object
        Dim args()
        Dim Chemin As String
        Dim Fichier As String
        Dim lcArchivoExcel As String
        Dim PrintArea(0)
        Dim Style As Object
        Dim Border As Object
        'encabezado
        Dim PageStyles As Object
        Dim Sheet As Object
        Dim StyleFamilies As Object
        Dim DefPage As Object
        Dim Htext As Object
        Dim Hcontent As Object
        Dim ret As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
    
    'Filtra los Datos
    Set rsReporte = ConsultarDiagnosticosOmitidosHIS(ml_Anio, ml_Mes)
    If rsReporte.RecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        GenerarRecordsetTemporal
        rsReporte.MoveFirst
             If lbEsOpenOffice = True Then
                'Abre el archivo ExcelOpenOffice
                lcArchivoExcel = App.Path + "\Plantillas\RptHisDxOmitidos.ods"
'                FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'                Chemin = "file:///" & App.Path & "\Plantillas\"
'                Chemin = Replace(Chemin, "\", "/")
'                Fichier = Chemin & "/OpenOffice.ods"
'
                Fichier = Format(Time, "hhmmss") & ".ods"
                FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
                lcArchivoExcel = Fichier
                Chemin = "file:///" & App.Path & "\Plantillas\"
                Chemin = Replace(Chemin, "\", "/")
                Fichier = Chemin & "/" & lcArchivoExcel
                '
                Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
                Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
                Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
                Set Feuille = Document.getSheets().getByIndex(0)
                'Encabezado de Pagina
                mo_CabeceraReportes.CabeceraReportes Document, True
                ' Pone la ventana en primer plano, pasándole el Hwnd
                ret = SetForegroundWindow(lnHwnd)
            Else
                Set oExcel = GalenhosExcelApplication()  'New Excel.Application
                'Crea nueva hoja
                Set oWorkBook = oExcel.Workbooks.Add
                'Abre, copia y cierra la plantilla
                Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\RptHisDxOmitidos.xls")
                oWorkBookPlantilla.Worksheets("hoja_Libre").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Close
                'Activa la primera hoja
                Set oWorkSheet = oWorkBook.Sheets(1)
                mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            End If
            
            iFila = 6
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":B" & CStr(iFila)): mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Set Plage = Feuille.getCellRangeByName("C" & CStr(iFila) & ":C" & CStr(iFila)): mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":D" & CStr(iFila)): mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Set Plage = Feuille.getCellRangeByName("E" & CStr(iFila) & ":E" & CStr(iFila)): mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Set Plage = Feuille.getCellRangeByName("F" & CStr(iFila) & ":F" & CStr(iFila)): mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 2
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, 3
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 4, iFila, 4
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 5, iFila, 5
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 6, iFila, 6
            End If
            
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, 3).setFormula(mc_Texto)
                lnNumTotal = 0
                iFila = iFila + 1
            Else
                oWorkSheet.Cells(4, 1).Value = mc_Texto
                lnNumTotal = 0
                iFila = iFila + 1
            End If
            
            lnCantidadAcumulada = 0
            Do While Not rsReporte.EOF
                lnCantidadAcumulada = lnCantidadAcumulada + Val(rsReporte.Fields("CANTIDAD").Value)
                
                If lbEsOpenOffice = True Then
                    Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":B" & CStr(iFila)): mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    Set Plage = Feuille.getCellRangeByName("C" & CStr(iFila) & ":C" & CStr(iFila)): mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":D" & CStr(iFila)): mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    Set Plage = Feuille.getCellRangeByName("E" & CStr(iFila) & ":E" & CStr(iFila)): mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    Set Plage = Feuille.getCellRangeByName("F" & CStr(iFila) & ":F" & CStr(iFila)): mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsReporte.Fields("CODVALIDACION").Value)
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula(rsReporte.Fields("TXTDESCRIPCION").Value)
                    Call Feuille.getcellbyposition(3, iFila - 1).setFormula(rsReporte.Fields("TXTVALIDACION").Value)
                    Call Feuille.getcellbyposition(4, iFila - 1).setFormula(rsReporte.Fields("CANTIDAD").Value)
                    Call Feuille.getcellbyposition(5, iFila - 1).setFormula(rsReporte.Fields("PORCENTAJE").Value)
                    iFila = iFila + 1
                Else
                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 2
                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 3, iFila, 3
                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 4, iFila, 4
                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 5, iFila, 5
                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 6, iFila, 6
                
                    oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields("CODVALIDACION").Value
                    oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields("TXTDESCRIPCION").Value
                    oWorkSheet.Cells(iFila, 4).Value = rsReporte.Fields("TXTVALIDACION").Value
                    oWorkSheet.Cells(iFila, 5).Value = rsReporte.Fields("CANTIDAD").Value
                    oWorkSheet.Cells(iFila, 6).Value = rsReporte.Fields("PORCENTAJE").Value
                    iFila = iFila + 1
                End If
                rsReporte.MoveNext
            Loop
            
            If lbEsOpenOffice = True Then
                 Set Plage = Feuille.getCellRangeByName("E" & CStr(iFila) & ":E" & CStr(iFila)): mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                 Set Plage = Feuille.getCellRangeByName("F" & CStr(iFila) & ":F" & CStr(iFila)): mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            
                 Call Feuille.getcellbyposition(4, iFila - 1).setFormula(lnCantidadAcumulada)
                 Call Feuille.getcellbyposition(5, iFila - 1).setFormula(1)
            Else
                 mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 5, iFila, 5
                 mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 6, iFila, 6
                 
                 oWorkSheet.Cells(iFila, 5).Value = lnCantidadAcumulada
                 oWorkSheet.Cells(iFila, 6).Value = 1
            End If
            
            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 1
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 22
                PrintArea(0).EndRow = iFila
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
                MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
            Else
                'Falta que salga el nombre del responsable
                If oWorkSheet.PageSetup.PrintArea <> "" Then
                   oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                End If
                oExcel.Visible = True
                oWorkSheet.PrintPreview
            End If
        End If
        
        If lbEsOpenOffice = True Then
            'Liberar Memoria
            Set Plage = Nothing
            Set Feuille = Nothing
            Set Document = Nothing
            Set Desktop = Nothing
            Set ServiceManager = Nothing
            Set Style = Nothing
            Set Border = Nothing
            'encabezado de pagina
            Set PageStyles = Nothing
            Set Sheet = Nothing
            Set StyleFamilies = Nothing
            Set DefPage = Nothing
            Set Htext = Nothing
            Set Hcontent = Nothing
        Else
            'Liberar memoria
            Set oExcel = Nothing
            Set oWorkBookPlantilla = Nothing
            Set oWorkBook = Nothing
            Set oWorkSheet = Nothing
        End If
   
Exit Sub
ManejadorErrorExcel:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub
End Sub

Sub GenerarRecordsetTemporal()
    With mrs_Tmp
          .Fields.Append "CODVALIDACION", adInteger, 4, adFldIsNullable
          .Fields.Append "TXTDESCRIPCION", adVarChar, 160, adFldIsNullable
          .Fields.Append "TXTVALIDACION", adVarChar, 50, adFldIsNullable
          .Fields.Append "CANTIDAD", adDate, , adFldIsNullable
          .Fields.Append "PORCENTAJE", adDate, , adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
End Sub

Function ConsultarDiagnosticosOmitidosHIS(lnAnio As Integer, lnMes As Integer) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
        Set ConsultarDiagnosticosOmitidosHIS = Nothing
        ms_MensajeError = ""
        oConexion.Open SIGHEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "HIS_ConsultarDiagnosticosOmitidosHIS"
            Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, lnAnio): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Mes", adInteger, adParamInput, 0, lnMes): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set ConsultarDiagnosticosOmitidosHIS = oRecordset
    Exit Function
ManejadorDeError:
       ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
    Exit Function
End Function


