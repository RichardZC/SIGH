VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptEFospolis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Fospolis
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReporteUtil As New ReporteUtil
Dim mda_FechaInicio As Date
Dim mda_FechaFin As Date
Dim ml_TextoDelFiltro  As String
Dim rsCostos As New ADODB.Recordset
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let TextoDelFiltro(lValue As String)
    ml_TextoDelFiltro = lValue
End Property

Property Let FechaInicio(daValue As Date)
    mda_FechaInicio = daValue
End Property
Property Let FechaFin(daValue As Date)
    mda_FechaFin = daValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New EFospolis
    oFormulario.Show 1
End Sub

Sub CrearReporte_excel(lnHwnd As Long)
Dim mrs_Tmp1 As New Recordset
Dim mrs_Tmp As New Recordset
Dim iFila As Long: Dim iColumna As Integer: Dim iColInicial As Integer
Dim lnGasto As Double: Dim lnNroHistorias As Long: Dim lnIdPaciente As Long
Dim lnTotPaciente As Double: Dim lnTotal As Double: Dim lcCodCCosto As String
Dim lnColumna As Integer: Dim lcDias As String: Dim lnNumDias As Integer
Dim lnPrecio As Double: Dim lnGastoCE As Double
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String

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
    Set mrs_Tmp1 = ml_ReglasFacturacion.ServiciosConsumoConveniosPorFechas(mda_FechaInicio, mda_FechaFin + 1, 1)
    If mrs_Tmp1.RecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        'Carga Centro Costos
        GenerarRecordsetTemporal
        CargaCentroCostos mda_FechaInicio, mda_FechaFin
        
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\Efospolis.ods"
'            FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'            Chemin = "file:///" & App.Path & "\Plantillas\"
'            Chemin = Replace(Chemin, "\", "/")
'            Fichier = Chemin & "/OpenOffice.ods"
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
            'Muestra Reporte
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            'Crea nueva hoja
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\Efospolis.xls")
            oWorkBookPlantilla.Worksheets("eFospolis").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, 2).setFormula(ml_TextoDelFiltro)
        Else
            oWorkSheet.Cells(3, 2).Value = ml_TextoDelFiltro
        End If
        
        iColInicial = 7
        'Imprime Cabecera de Columas
        rsCostos.MoveFirst
        Do While Not rsCostos.EOF
           lnColumna = iColInicial + rsCostos.Fields!columna
           ml_TextoDelFiltro = rsCostos.Fields!dcosto
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(lnColumna - 1, 4).setFormula(ml_TextoDelFiltro)
            Else
                oWorkSheet.Cells(5, lnColumna).Value = ml_TextoDelFiltro
            End If
           rsCostos.MoveNext
        Loop
        lnColumna = iColInicial + rsCostos.RecordCount + 1
        ml_TextoDelFiltro = "SALDO PENDIENTE"
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(lnColumna - 1, 4).setFormula(ml_TextoDelFiltro)
            Else
                oWorkSheet.Cells(5, lnColumna).Value = ml_TextoDelFiltro
            End If
        'Imprime Datos
        iFila = 7: lnNroHistorias = 0: lnTotal = 0
        mrs_Tmp1.MoveFirst
        Do While Not mrs_Tmp1.EOF
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula(Trim(mrs_Tmp1.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp1.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp1.Fields!PrimerNombre))
            Else
                oWorkSheet.Cells(iFila, 2).Value = Trim(mrs_Tmp1.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp1.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp1.Fields!PrimerNombre)
            End If
           'Consulta Externa por 15 minutos
           Set mrs_Tmp = ml_ReglasFacturacion.ServiciosConsumoConvenioConsultaExternaPorFechas(mda_FechaInicio, mda_FechaFin + 1, mrs_Tmp1.Fields!IdPaciente)
           lcDias = "": lnNumDias = 0: lnPrecio = 0: lnGasto = 0
           
           Do While Not mrs_Tmp.EOF
              lnPrecio = mrs_Tmp.Fields!PrecioUnitario
              lnNumDias = lnNumDias + 1
              lcDias = lcDias & Day(mrs_Tmp.Fields!FechaIngreso) & ", "
              lnGasto = lnGasto + mrs_Tmp.Fields!totalPorPagar
              mrs_Tmp.MoveNext
           Loop
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(3, iFila - 1).setFormula(lcDias)
                Call Feuille.getcellbyposition(4, iFila - 1).setFormula(lnNumDias)
                Call Feuille.getcellbyposition(5, iFila - 1).setFormula(lnPrecio)
                Call Feuille.getcellbyposition(6, iFila - 1).setFormula(lnGasto)
            Else
                oWorkSheet.Cells(iFila, 4) = lcDias
                oWorkSheet.Cells(iFila, 5) = lnNumDias
                oWorkSheet.Cells(iFila, 6) = lnPrecio
                oWorkSheet.Cells(iFila, 7) = lnGasto
            End If
           lnGastoCE = lnGastoCE + lnGasto
           lnTotal = lnTotal + lnGasto
           mrs_Tmp.Close
           
           'Gastos en Servicios
           lnTotPaciente = lnGasto
           lnIdPaciente = mrs_Tmp1.Fields!IdPaciente
           Do While Not mrs_Tmp1.EOF And lnIdPaciente = mrs_Tmp1.Fields!IdPaciente
              lcCodCCosto = mrs_Tmp1.Fields!codCCosto
              lnGasto = 0
              rsCostos.MoveFirst
              rsCostos.Find "cCosto='" & lcCodCCosto & "'"
              'Gastos en Servicios
              lnColumna = rsCostos.Fields!columna
              Do While Not mrs_Tmp1.EOF And lnIdPaciente = mrs_Tmp1.Fields!IdPaciente And lcCodCCosto = mrs_Tmp1.Fields!codCCosto
                    lnGasto = lnGasto + mrs_Tmp1.Fields!TotalFinanciado
                    mrs_Tmp1.MoveNext
                    If mrs_Tmp1.EOF Then
                      Exit Do
                    End If
              Loop
              rsCostos.Fields!Importe = rsCostos.Fields!Importe + lnGasto
              rsCostos.Update
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(iColInicial - 1 + lnColumna - 1, iFila - 1).setFormula(lnGasto)
                Else
                    oWorkSheet.Cells(iFila, iColInicial + lnColumna).Value = lnGasto
                End If
              lnTotPaciente = lnTotPaciente + lnGasto
              lnTotal = lnTotal + lnGasto
              If mrs_Tmp1.EOF Then
                   Exit Do
              End If
           Loop
           lnColumna = iColInicial + rsCostos.RecordCount + 1
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(iColInicial - 1 + iFila - 1).setFormula(lnTotPaciente)
                Else
                    oWorkSheet.Cells(iFila, lnColumna).Value = lnTotPaciente
                End If
           lnNroHistorias = lnNroHistorias + 1
           iFila = iFila + 1
           If mrs_Tmp1.EOF Then
                 Exit Do
           End If
        Loop
        iFila = iFila + 1
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & iColInicial + rsCostos.RecordCount + 1 & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Nº Pacientes: " + Trim(Str(lnNroHistorias)))
            Call Feuille.getcellbyposition(6, iFila - 1).setFormula(lnGastoCE)
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, iColInicial + rsCostos.RecordCount + 1
            oWorkSheet.Cells(iFila, 2).Value = "Nº Pacientes: " + Trim(Str(lnNroHistorias))
            'Imprime Cabecera de Columas
            oWorkSheet.Cells(iFila, 7).Value = lnGastoCE
        End If
        rsCostos.MoveFirst
        Do While Not rsCostos.EOF
           lnColumna = iColInicial + rsCostos.Fields!columna
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(lnColumna - 1, iFila - 1).setFormula(rsCostos.Fields!Importe)
            Else
                oWorkSheet.Cells(iFila, lnColumna).Value = rsCostos.Fields!Importe
           End If
           rsCostos.MoveNext
        Loop
        lnColumna = iColInicial + rsCostos.RecordCount + 1
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(lnColumna - 1, iFila - 1).setFormula(lnTotal)
        Else
            oWorkSheet.Cells(iFila, lnColumna).Value = lnTotal
        End If
        
        If lbEsOpenOffice = True Then
            Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
            PrintArea(0).Sheet = 0
            PrintArea(0).startcolumn = 1
            PrintArea(0).StartRow = 0
            PrintArea(0).EndColumn = 51
            PrintArea(0).EndRow = iFila
            Call Feuille.SetPrintAreas(PrintArea())
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
            oWorkSheet.PageSetup.PrintTitleRows = "$1:$6"
            If oWorkSheet.PageSetup.PrintArea <> "" Then
               oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
            End If
            oExcel.Visible = True
            oWorkSheet.PrintPreview
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
    With rsCostos
          .Fields.Append "CCosto", adBSTR, 20
          .Fields.Append "dCosto", adBSTR, 100
          .Fields.Append "Columna", adInteger
          .Fields.Append "Importe", adDouble
          .LockType = adLockOptimistic
          .Open
    End With
End Sub








Sub CargaCentroCostos(ldFechaIni As Date, ldFechaFin As Date)
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim sSQL1 As String: Dim sSQL2 As String
Dim lnColumna As Integer: Dim lcCodCosto As String
    Set oRecordset = ml_ReglasFacturacion.ServiciosConsumoConveniosPorFechas(ldFechaIni, ldFechaFin, 1)
    If oRecordset.RecordCount > 0 Then
        lnColumna = 1
        oRecordset.MoveFirst
        Do While Not oRecordset.EOF
           lcCodCosto = oRecordset.Fields!codCCosto
           rsCostos.AddNew
           rsCostos.Fields!cCosto = lcCodCosto
           rsCostos.Fields!dcosto = oRecordset.Fields!dCCosto
           rsCostos.Fields!columna = lnColumna
           rsCostos.Fields!Importe = 0
           rsCostos.Update
           
           Do While Not oRecordset.EOF And lcCodCosto = oRecordset.Fields!codCCosto
              oRecordset.MoveNext
              If oRecordset.EOF Then
                 Exit Do
              End If
           Loop
           lnColumna = lnColumna + 1
        Loop
    End If
    Exit Sub
ManejadorDeError:
   MsgBox Err.Description
   Exit Sub
End Sub




