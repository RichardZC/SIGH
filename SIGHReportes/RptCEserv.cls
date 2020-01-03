VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptCEservi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Citas por Consultorios
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
    Dim oFormulario As New CEservicios
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
Dim lcCie10 As String: Dim lcDx As String
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
    Set rsReporte = ml_ReglasFacturacion.ServicioBoletasPorFechas(CDate(mda_FechaInicio), CDate(mda_FechaFin))
    If ml_IdServicioCE > 0 Then
       rsReporte.Filter = "IdEspecialidad=" & ml_IdServicioCE
    End If
    lnNumHistorias = rsReporte.RecordCount
    If lnNumHistorias = 0 Then
            MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
            GenerarRecordsetTemporal
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
               If rsReporte.Fields!idProducto = 4688 Or rsReporte.Fields!idProducto = 4690 Or rsReporte.Fields!idProducto = 4583 Then   'Carnet/folder/consulta
                    rsReporte.MoveNext
               Else
                    lcCie10 = rsReporte.Fields!Codigo
                    lcDx = rsReporte.Fields!Nombre
                    lnGastosConsulta = 0: lnTotGastosConsulta = 0
                    Do While Not rsReporte.EOF And lcCie10 = rsReporte.Fields!Codigo
                         lnGastosConsulta = lnGastosConsulta + 1
                         lnTotGastosConsulta = lnTotGastosConsulta + rsReporte.Fields!totalPorPagar
                         rsReporte.MoveNext
                         If rsReporte.EOF Then
                            Exit Do
                         End If
                    Loop
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!Codigo = lcCie10
                    mrs_Tmp.Fields!Servicio = lcDx
                    mrs_Tmp.Fields!Nro = lnGastosConsulta
                    mrs_Tmp.Fields!Porc = Round((lnGastosConsulta * 100) / lnNumHistorias, 2)
                    mrs_Tmp.Fields!Importe = lnTotGastosConsulta
                    mrs_Tmp.Update
               End If
            Loop
            If mrs_Tmp.RecordCount = 0 Then
               MsgBox "No existe Información con esos datos", vbInformation, "Mensaje"
            Else
            mrs_Tmp.Sort = "nro desc"
            
                If lbEsOpenOffice = True Then
                    'Abre el archivo ExcelOpenOffice
                    lcArchivoExcel = App.Path + "\Plantillas\CEservicios.ods"
'                    FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'                    Chemin = "file:///" & App.Path & "\Plantillas\"
'                    Chemin = Replace(Chemin, "\", "/")
'                    Fichier = Chemin & "/OpenOffice.ods"
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
                    Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEservicios.xls")
                    oWorkBookPlantilla.Worksheets("ceServicios").Copy Before:=oWorkBook.Sheets(1)
                    oWorkBookPlantilla.Close
                    'Activa la primera hoja
                    Set oWorkSheet = oWorkBook.Sheets(1)
                    mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
                End If
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, 1).setFormula("FRECUENCIA DE CONSUMO DE SERVICIOS (CPT) DE PACIENTES ATENDIDOS EN C.E.")
                    Call Feuille.getcellbyposition(1, 2).setFormula(ml_TextoDelFiltro)
                    Call Feuille.getcellbyposition(1, 4).setFormula("Procedimientos (Cpt)")
                Else
                    oWorkSheet.Cells(2, 2).Value = "FRECUENCIA DE CONSUMO DE SERVICIOS (CPT) DE PACIENTES ATENDIDOS EN C.E."
                    oWorkSheet.Cells(3, 2).Value = ml_TextoDelFiltro
                    oWorkSheet.Cells(5, 2).Value = "Procedimientos (CPT)"
                End If
            iFila = 6
            lnNumTotal = 0: lnTotGastosConsulta = 0: lnNumHistorias = 0
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(mrs_Tmp.Fields("codigo").Value & " " & mrs_Tmp.Fields("servicio").Value)
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula(mrs_Tmp.Fields("importe").Value)
                    Call Feuille.getcellbyposition(3, iFila - 1).setFormula(mrs_Tmp.Fields("nro").Value)
                    Call Feuille.getcellbyposition(4, iFila - 1).setFormula(mrs_Tmp.Fields("Porc").Value)
                Else
                    oWorkSheet.Cells(iFila, 2).Value = mrs_Tmp.Fields("codigo").Value & " " & mrs_Tmp.Fields("servicio").Value
                    oWorkSheet.Cells(iFila, 3).Value = mrs_Tmp.Fields("importe").Value
                    oWorkSheet.Cells(iFila, 4).Value = mrs_Tmp.Fields("nro").Value
                    oWorkSheet.Cells(iFila, 5).Value = mrs_Tmp.Fields("Porc").Value
                End If
               lnTotGastosConsulta = lnTotGastosConsulta + mrs_Tmp.Fields("importe").Value
               lnNumHistorias = lnNumHistorias + mrs_Tmp.Fields("nro").Value
               mrs_Tmp.MoveNext
               iFila = iFila + 1
               lnNumTotal = lnNumTotal + 1
            Loop
                If lbEsOpenOffice = True Then
                    Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":E" & CStr(iFila))
                    mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Nro Servicios: " & Trim(Str(lnNumTotal)))
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula(lnTotGastosConsulta)
                    Call Feuille.getcellbyposition(3, iFila - 1).setFormula(lnNumHistorias)
                Else
                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 5
                    oWorkSheet.Cells(iFila, 2).Value = "Nro Servicios: " & Trim(Str(lnNumTotal))
                    oWorkSheet.Cells(iFila, 3).Value = lnTotGastosConsulta
                    oWorkSheet.Cells(iFila, 4).Value = lnNumHistorias
                End If
                If lbEsOpenOffice = True Then
                    Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                    PrintArea(0).Sheet = 0
                    PrintArea(0).startcolumn = 1
                    PrintArea(0).StartRow = 0
                    PrintArea(0).EndColumn = 5
                    PrintArea(0).EndRow = iFila
                    Call Feuille.SetPrintAreas(PrintArea())
                    Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                    MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
                Else
                    oWorkSheet.PageSetup.PrintTitleRows = "$1:$5"
                    If oWorkSheet.PageSetup.PrintArea <> "" Then
                       oWorkSheet.PageSetup.PrintArea = sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                    End If
                    oExcel.Visible = True
                    oWorkSheet.PrintPreview
                    'oWorkSheet.PrintOut
                End If
            End If
            If lbEsOpenOffice = True Then
                Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
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
                'liberar memoria
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
    'Resume
    Exit Sub
End Sub






Sub GenerarRecordsetTemporal()
    With mrs_Tmp
          .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "Servicio", adVarChar, 250, adFldIsNullable
          .Fields.Append "Nro", adInteger
          .Fields.Append "Porc", adDouble
          .Fields.Append "importe", adDouble
          .LockType = adLockOptimistic
          .Open
    End With
End Sub

