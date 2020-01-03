VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptCEgastosDePacientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Frecuencia de Gastos de pacientes
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mrs_Tmp As New ADODB.Recordset
Dim ml_IdServicioCE As Long
Dim ml_IdResponsable As Long
Dim mo_Conexion As ADODB.Connection
Dim mo_ReporteUtil As New ReporteUtil
Dim ml_OrdenFiltro As String
Dim mb_SonDatosDePruebas As Boolean
Dim ml_TextoDelFiltro As String
Dim ml_TextoDelFiltro1 As String
Dim mda_FechaInicio As String
Dim mda_FechaFin As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let FechaInicio(daValue As String)
    mda_FechaInicio = daValue
End Property
Property Let FechaFin(daValue As String)
    mda_FechaFin = daValue
End Property
Property Let TextoDelFiltro(lValue As String)
    ml_TextoDelFiltro = lValue
End Property
Property Let TextoDelFiltro1(lValue As String)
    ml_TextoDelFiltro1 = lValue
End Property
Property Let SonDatosDePruebas(lValue As Boolean)
    mb_SonDatosDePruebas = lValue
End Property

Property Let IdServicioCE(lValue As Long)
    ml_IdServicioCE = lValue
End Property
Property Let IdResponsable(lValue As Long)
    ml_IdResponsable = lValue
End Property
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Let OrdenFiltro(lValue As String)
    ml_OrdenFiltro = lValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New CEGastosDePacientes
    oFormulario.Show 1
End Sub

Sub CrearReporte_excel(lnHwnd As Long)
Dim rsReporte As New Recordset
Dim iFila As Long
Dim lnNumHistorias As Long
Dim lnIdPaciente As Long
Dim lcPaciente As String
Dim lnNumTotal As Long
Dim oDOPaciente As New doPaciente
Dim oDOAtencion As New DOAtencion
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim lnLineas As Integer: Dim lnIdAtencion  As Long
Dim lcProcedencia As String
Dim lnGastosConsulta As Double: Dim lnTotGastosConsulta As Double
Dim lnGastosServicios As Double: Dim lnTotGastosServicios As Double
Dim lnGastosFarmacia As Double: Dim lnTotGastosFarmacia As Double
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
    Set mrs_Tmp = ml_ReglasFacturacion.AtencionesConsultaExternaPorFechas(CDate(mda_FechaInicio), CDate(mda_FechaFin) + 1)
    If mrs_Tmp.RecordCount = 0 Then
            MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\CEatencionesGastos.ods"
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
            
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            'Crea nueva hoja
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEatencionesGastos.xls")
            oWorkBookPlantilla.Worksheets("ceAtencionGastos").Copy Before:=oWorkBook.Sheets(1)
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
            iFila = 6
            lnNumHistorias = 0
            lnTotGastosServicios = 0: lnTotGastosConsulta = 0: lnTotGastosFarmacia = 0
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
                If lbEsOpenOffice = True Then
                   Call Feuille.getcellbyposition(1, 2).setFormula(mrs_Tmp.Fields("nroHistoriaClinica").Value)
                   Call Feuille.getcellbyposition(1, 2).setFormula(Trim(mrs_Tmp.Fields("apellidoPaterno").Value) & " " & Trim(mrs_Tmp.Fields("apellidoMaterno").Value) & " " & Trim(mrs_Tmp.Fields("PrimerNombre").Value))
                Else
                    oWorkSheet.Cells(iFila, 2).Value = mrs_Tmp.Fields("nroHistoriaClinica").Value
                    oWorkSheet.Cells(iFila, 3).Value = Trim(mrs_Tmp.Fields("apellidoPaterno").Value) & " " & Trim(mrs_Tmp.Fields("apellidoMaterno").Value) & " " & Trim(mrs_Tmp.Fields("PrimerNombre").Value)
                End If
                
               lnNumHistorias = lnNumHistorias + 1
               lnGastosServicios = 0: lnGastosConsulta = 0
               'Gastos en Servicios
               'Set rsReporte = BuscaGastosEnServicios(mrs_Tmp.Fields!idPaciente, mrs_Tmp.Fields!idAtencion)
               Set rsReporte = ml_ReglasFacturacion.ServicioBoletasPorIdPacienteIdCuenta(mrs_Tmp.Fields!IdPaciente, mrs_Tmp.Fields!idCuentaAtencion)
               If rsReporte.RecordCount > 0 Then
                    rsReporte.MoveFirst
                    Do While Not rsReporte.EOF
                       If rsReporte.Fields!idProducto = 4688 Or rsReporte.Fields!idProducto = 4690 Or rsReporte.Fields!idProducto = 4583 Or rsReporte.Fields!idProducto = 4727 Then   'Carnet/folder/consulta
                          lnGastosConsulta = lnGastosConsulta + rsReporte.Fields!totalPorPagar
                       Else
                          lnGastosServicios = lnGastosServicios + rsReporte.Fields!totalPorPagar
                       End If
                       rsReporte.MoveNext
                    Loop
               End If
               rsReporte.Close
               'Gastos en Farmacia
               lnGastosFarmacia = 0
               Set rsReporte = ml_ReglasFacturacion.FarmaciaBoletasPorIdPacienteIdCuenta(mrs_Tmp.Fields!IdPaciente, mrs_Tmp.Fields!idCuentaAtencion)
               If rsReporte.RecordCount > 0 Then
                    rsReporte.MoveFirst
                    Do While Not rsReporte.EOF
                       lnGastosFarmacia = lnGastosFarmacia + rsReporte.Fields!TotalPagar
                       rsReporte.MoveNext
                    Loop
               End If
               rsReporte.Close

                If lbEsOpenOffice = True Then
                   Call Feuille.getcellbyposition(3, iFila - 1).setFormula(lnGastosConsulta)
                   Call Feuille.getcellbyposition(4, iFila - 1).setFormula(lnGastosServicios)
                   Call Feuille.getcellbyposition(5, iFila - 1).setFormula(lnGastosFarmacia)
                Else

                   oWorkSheet.Cells(iFila, 4).Value = lnGastosConsulta ' Format(lnGastosConsulta, "##,###,###.##")
                   oWorkSheet.Cells(iFila, 5).Value = lnGastosServicios 'Format(lnGastosServicios, "##,###,###.##")
                   oWorkSheet.Cells(iFila, 6).Value = lnGastosFarmacia 'Format(lnGastosFarmacia, "##,###,###.##")
                End If
               mrs_Tmp.MoveNext
               lnTotGastosServicios = lnTotGastosServicios + lnGastosServicios
               lnTotGastosConsulta = lnTotGastosConsulta + lnGastosConsulta
               lnTotGastosFarmacia = lnTotGastosFarmacia + lnGastosFarmacia
               iFila = iFila + 1
            Loop
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":F" & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Nro Historias: " & Trim(Str(lnNumHistorias)))
                Call Feuille.getcellbyposition(3, iFila - 1).setFormula(lnTotGastosConsulta)
                Call Feuille.getcellbyposition(4, iFila - 1).setFormula(lnTotGastosServicios)
                Call Feuille.getcellbyposition(5, iFila - 1).setFormula(lnTotGastosFarmacia)
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 6
                oWorkSheet.Cells(iFila, 2).Value = "Nro Historias: " & Trim(Str(lnNumHistorias))
                oWorkSheet.Cells(iFila, 4).Value = lnTotGastosConsulta
                oWorkSheet.Cells(iFila, 5).Value = lnTotGastosServicios
                oWorkSheet.Cells(iFila, 6).Value = lnTotGastosFarmacia
            End If
            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 1
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 6
                PrintArea(0).EndRow = iFila
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
            Else
                oWorkSheet.PageSetup.PrintTitleRows = "$1:$5"
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





