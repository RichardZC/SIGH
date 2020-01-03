VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptHProcedimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para procedimientos
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mrs_Tmp As New ADODB.Recordset
Dim mrs_Tmp1 As New ADODB.Recordset
Dim mo_ReporteUtil As New ReporteUtil
Dim ml_TipoReporte As Integer
Dim ml_Titulo As String
Dim ml_TextoDelFiltro As String
Dim ml_idDepartamento1 As Long
Dim ml_idEspecialidad1 As Long
Dim ml_idServicio1 As Long
Dim ml_idDepartamento2 As Long
Dim ml_idEspecialidad2 As Long
Dim ml_idServicio2 As Long
Dim ml_IdDistrito As Long
Dim ml_idDiagnostico As Long
Dim mda_FechaInicio As Date
Dim mda_FechaFin As Date
Dim ml_idTipoSexo As Integer
Dim ml_idTipoDiagnostico As Integer
Dim ml_AnioCol11 As Double
Dim ml_AnioCol12 As Double
Dim ml_AnioCol21 As Double
Dim ml_AnioCol22 As Double
Dim ml_AnioCol31 As Double
Dim ml_AnioCol32 As Double
Dim ml_AnioCol41 As Double
Dim ml_AnioCol42 As Double
Dim ml_AnioCol51 As Double
Dim ml_AnioCol52 As Double
Dim ml_AnioCol61 As Double
Dim ml_AnioCol62 As Double
Dim ml_AnioCol71 As Double
Dim ml_AnioCol72 As Double
Dim ml_AnioCol81 As Double
Dim ml_AnioCol82 As Double
Dim ml_DetallaHC As Boolean
Dim ml_TituloCol1 As String
Dim ml_TituloCol2 As String
Dim ml_TituloCol3 As String
Dim ml_TituloCol4 As String
Dim ml_TituloCol5 As String
Dim ml_TituloCol6 As String
Dim ml_TituloCol7 As String
Dim ml_TituloCol8 As String
Dim ml_PuntoCarga As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lcNombre As String, lcSql As String
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim ml_idTipoServicio As Long
Property Let PuntoCarga(lValue As Long)
    ml_PuntoCarga = lValue
End Property
Property Let TituloCol1(lValue As String)
    ml_TituloCol1 = lValue
End Property
Property Let TituloCol2(lValue As String)
    ml_TituloCol2 = lValue
End Property
Property Let TituloCol3(lValue As String)
    ml_TituloCol3 = lValue
End Property
Property Let TituloCol4(lValue As String)
    ml_TituloCol4 = lValue
End Property
Property Let TituloCol5(lValue As String)
    ml_TituloCol5 = lValue
End Property
Property Let TituloCol6(lValue As String)
    ml_TituloCol6 = lValue
End Property
Property Let TituloCol7(lValue As String)
    ml_TituloCol7 = lValue
End Property
Property Let TituloCol8(lValue As String)
    ml_TituloCol8 = lValue
End Property
Property Let DetallaHC(lValue As Boolean)
    ml_DetallaHC = lValue
End Property
Property Let AnioCol11(lValue As Double)
    ml_AnioCol11 = lValue
End Property
Property Let AnioCol12(lValue As Double)
    ml_AnioCol12 = lValue
End Property
Property Let AnioCol21(lValue As Double)
    ml_AnioCol21 = lValue
End Property
Property Let AnioCol22(lValue As Double)
    ml_AnioCol22 = lValue
End Property
Property Let AnioCol31(lValue As Double)
    ml_AnioCol31 = lValue
End Property
Property Let AnioCol32(lValue As Double)
    ml_AnioCol32 = lValue
End Property
Property Let AnioCol41(lValue As Double)
    ml_AnioCol41 = lValue
End Property
Property Let AnioCol42(lValue As Double)
    ml_AnioCol42 = lValue
End Property
Property Let AnioCol51(lValue As Double)
    ml_AnioCol51 = lValue
End Property
Property Let AnioCol52(lValue As Double)
    ml_AnioCol52 = lValue
End Property
Property Let AnioCol61(lValue As Double)
    ml_AnioCol61 = lValue
End Property
Property Let AnioCol62(lValue As Double)
    ml_AnioCol62 = lValue
End Property
Property Let AnioCol71(lValue As Double)
    ml_AnioCol71 = lValue
End Property
Property Let AnioCol72(lValue As Double)
    ml_AnioCol72 = lValue
End Property
Property Let AnioCol81(lValue As Double)
    ml_AnioCol81 = lValue
End Property
Property Let AnioCol82(lValue As Double)
    ml_AnioCol82 = lValue
End Property
Property Let idTipoDiagnostico(lValue As Integer)
    ml_idTipoDiagnostico = lValue
End Property
Property Let idTipoSexo(lValue As Integer)
    ml_idTipoSexo = lValue
End Property
Property Let FechaInicio(daValue As Date)
    mda_FechaInicio = daValue
End Property
Property Let FechaFin(daValue As Date)
    mda_FechaFin = daValue
End Property
Property Let idDepartamento2(lValue As Long)
    ml_idDepartamento2 = lValue
End Property
Property Let idServicio2(lValue As Long)
    ml_idServicio2 = lValue
End Property
Property Let idEspecialidad2(lValue As Long)
    ml_idEspecialidad2 = lValue
End Property
Property Let idDepartamento1(lValue As Long)
    ml_idDepartamento1 = lValue
End Property
Property Let idServicio1(lValue As Long)
    ml_idServicio1 = lValue
End Property
Property Let idEspecialidad1(lValue As Long)
    ml_idEspecialidad1 = lValue
End Property
Property Let IdDistrito(lValue As Long)
    ml_IdDistrito = lValue
End Property
Property Let idDiagnostico(lValue As Long)
    ml_idDiagnostico = lValue
End Property
Property Let Titulo(lValue As String)
    ml_Titulo = lValue
End Property
Property Let TextoDelFiltro(lValue As String)
    ml_TextoDelFiltro = lValue
End Property
Property Let TipoReporte(lValue As Integer)
    ml_TipoReporte = lValue
End Property
Property Let idTipoServicio(lValue As Long)
    ml_idTipoServicio = lValue
End Property
Sub EjecutaFormulario()
    Dim oFormulario As New HProcedimientos
    oFormulario.Show 1
End Sub

Sub CrearReporte(lnHwnd As Long, lbEnCantidades As Boolean)
    GenerarRecordsetTemporal
    CrearReporteDetallado lnHwnd, lbEnCantidades
End Sub

Sub CrearReporteDetallado(lnHwnd As Long, lbEnCantidades As Boolean)
Dim rsReporte As New Recordset: Dim rsTmpReporte As New Recordset: Dim rsTmp1 As New Recordset
Dim iFila As Long: Dim iCol As Integer
Dim lnTotNro As Integer: Dim lnTotCol1 As Integer: Dim lnTotCol2 As Integer: Dim lnTotCol3 As Integer: Dim lnTotCol4 As Integer: Dim lnTotCol5 As Integer: Dim lnTotCol6 As Integer: Dim lnTotCol7 As Integer: Dim lnTotCol8 As Integer
Dim lnTotDNro As Integer: Dim lnTotDCol1 As Integer: Dim lnTotDCol2 As Integer: Dim lnTotDCol3 As Integer: Dim lnTotDCol4 As Integer: Dim lnTotDCol5 As Integer: Dim lnTotDCol6 As Integer: Dim lnTotDCol7 As Integer: Dim lnTotDCol8 As Integer
Dim lnNumero As Integer: Dim lnEdadAnios As Double
Dim ldDiagnostico As String: Dim lnDiagnostico As Long: Dim lcCie10 As String
Dim lcFiltro As String
Dim lcDescripcion As String: Dim lnProducto As Long: Dim lcCodigoCpt As String
Dim ml_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim lbEsOpenOffice As Boolean


lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
On Error GoTo ManejadorError


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
    
    Set rsReporte = FacturacionServiciosSeleccionarCPT(ml_idTipoSexo, mda_FechaInicio, mda_FechaFin)
    If rsReporte.RecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        'Filtros,Totales para Porcentaje y Temporal
        lnTotNro = 0: lnTotCol1 = 0: lnTotCol2 = 0: lnTotCol3 = 0: lnTotCol4 = 0: lnTotCol5 = 0: lnTotCol6 = 0: lnTotCol7 = 0: lnTotCol8 = 0
        lcFiltro = "idestadofacturacion<>9 and idPuntoCarga=" & ml_PuntoCarga & " and idTipoServicio=" & ml_idTipoServicio
        Select Case ml_TipoReporte
        Case 1   'Por un Departamento
            If lcFiltro = "" Then
               lcFiltro = "CodDpto=" & ml_idDepartamento1
            Else
               lcFiltro = lcFiltro & " and CodDpto=" & ml_idDepartamento1
            End If
        Case 2   'Por un Servicio
            If lcFiltro = "" Then
               lcFiltro = "CodEspecialidad=" & ml_idEspecialidad1
            Else
               lcFiltro = lcFiltro & " and CodEspecialidad=" & ml_idEspecialidad1
            End If
        Case 3   'Por dos Especialidades
            If lcFiltro = "" Then
               lcFiltro = "(idServicioEgreso=" & ml_idServicio1 & " or idServicioEgreso=" & ml_idServicio2 & ")"
            Else
               lcFiltro = lcFiltro & " and (idServicioEgreso=" & ml_idServicio1 & " or idServicioEgreso=" & ml_idServicio2 & ")"
            End If
        Case 4    'por un Servicio
            If lcFiltro = "" Then
               If ml_idTipoServicio <> 1 Then
                  lcFiltro = "idServicioEgreso=" & ml_idServicio1
               Else
                  lcFiltro = "idServicioIngreso=" & ml_idServicio1
               End If
            Else
               If ml_idTipoServicio <> 1 Then
                    lcFiltro = lcFiltro & " and idServicioEgreso=" & ml_idServicio1
               Else
                    lcFiltro = lcFiltro & " and idServicioIngreso=" & ml_idServicio1
               End If
            End If
        End Select
        If lcFiltro <> "" Then
           rsReporte.Filter = lcFiltro
        End If
        If rsReporte.RecordCount = 0 Then
           MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
        Else
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
                lnTotDNro = 0: lnTotDCol1 = 0: lnTotDCol2 = 0: lnTotDCol3 = 0: lnTotDCol4 = 0: lnTotDCol5 = 0: lnTotDCol6 = 0: lnTotDCol7 = 0: lnTotDCol8 = 0
                lnNumero = lnNumero + 1
                lcDescripcion = Left(rsReporte.Fields("NombreMINSA").Value & "-" & rsReporte.Fields("grupoCpt"), 150)
                lnProducto = rsReporte.Fields("IdProducto").Value
                lcCodigoCpt = rsReporte.Fields("CodigoCpt").Value
                Do While Not rsReporte.EOF And lnProducto = rsReporte.Fields("IdProducto").Value
                    mrs_Tmp1.AddNew
                    mrs_Tmp1.Fields("Codigo").Value = lcCodigoCpt
                    mrs_Tmp1.Fields("NroHistoria").Value = rsReporte.Fields("nroHistoriaClinica").Value
                    mrs_Tmp1.Fields("Fecha").Value = rsReporte.Fields("FechaOrden").Value
                    mrs_Tmp1.Fields("Edad").Value = Trim(Str(rsReporte.Fields("EdadNum").Value)) & IIf(rsReporte.Fields("idTipoEdad").Value = 1, " Años", IIf(rsReporte.Fields("idTipoEdad").Value = 2, " Meses", " Días"))
                    mrs_Tmp1.Update
                    lnEdadAnios = rsReporte.Fields("EdadNum").Value
                    If rsReporte.Fields("idTipoEdad").Value = 2 Then   'Meses
                       lnEdadAnios = sighentidades.ConvierteEnAnioUnMesOdia(rsReporte.Fields("EdadNum").Value, 1)
                    ElseIf rsReporte.Fields("idTipoEdad").Value = 3 Then  'Dias
                       lnEdadAnios = sighentidades.ConvierteEnAnioUnMesOdia(rsReporte.Fields("EdadNum").Value, 0)
                    ElseIf rsReporte.Fields("idTipoEdad").Value = 4 Then  'Horas
                       lnEdadAnios = sighentidades.ConvierteEnAnioUnMesOdia(rsReporte.Fields("EdadNum").Value, 4)
                    End If
                    If lnEdadAnios >= ml_AnioCol11 And lnEdadAnios <= ml_AnioCol12 Then
                       lnTotDCol1 = lnTotDCol1 + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotCol1 = lnTotCol1 + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotDNro = lnTotDNro + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotNro = lnTotNro + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                    End If
                    If lnEdadAnios >= ml_AnioCol21 And lnEdadAnios <= ml_AnioCol22 Then
                       lnTotDCol2 = lnTotDCol2 + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotCol2 = lnTotCol2 + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotDNro = lnTotDNro + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotNro = lnTotNro + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                    End If
                    If lnEdadAnios >= ml_AnioCol31 And lnEdadAnios <= ml_AnioCol32 Then
                       lnTotDCol3 = lnTotDCol3 + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotCol3 = lnTotCol3 + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotDNro = lnTotDNro + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotNro = lnTotNro + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                    End If
                    If lnEdadAnios >= ml_AnioCol41 And lnEdadAnios <= ml_AnioCol42 Then
                       lnTotDCol4 = lnTotDCol4 + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotCol4 = lnTotCol4 + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotDNro = lnTotDNro + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotNro = lnTotNro + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                    End If
                    If lnEdadAnios >= ml_AnioCol51 And lnEdadAnios <= ml_AnioCol52 Then
                       lnTotDCol5 = lnTotDCol5 + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotCol5 = lnTotCol5 + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotDNro = lnTotDNro + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotNro = lnTotNro + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                    End If
                    If lnEdadAnios >= ml_AnioCol61 And lnEdadAnios <= ml_AnioCol62 Then
                       lnTotDCol6 = lnTotDCol6 + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotCol6 = lnTotCol6 + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotDNro = lnTotDNro + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotNro = lnTotNro + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                    End If
                    If lnEdadAnios >= ml_AnioCol71 And lnEdadAnios <= ml_AnioCol72 Then
                       lnTotDCol7 = lnTotDCol7 + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotCol7 = lnTotCol7 + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotDNro = lnTotDNro + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotNro = lnTotNro + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                    End If
                    If lnEdadAnios >= ml_AnioCol81 And lnEdadAnios <= ml_AnioCol82 Then
                       lnTotDCol8 = lnTotDCol8 + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotCol8 = lnTotCol8 + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotDNro = lnTotDNro + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                       lnTotNro = lnTotNro + IIf(lbEnCantidades = True, rsReporte!Cantidad, 1)
                    End If
                    rsReporte.MoveNext
                    If rsReporte.EOF Then
                       Exit Do
                    End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields("Codigo").Value = lcCodigoCpt
                mrs_Tmp.Fields("Descripcion").Value = lcDescripcion
                mrs_Tmp.Fields("Nro").Value = lnTotDNro
                mrs_Tmp.Fields("Col1").Value = lnTotDCol1
                mrs_Tmp.Fields("Col2").Value = lnTotDCol2
                mrs_Tmp.Fields("Col3").Value = lnTotDCol3
                mrs_Tmp.Fields("Col4").Value = lnTotDCol4
                mrs_Tmp.Fields("Col5").Value = lnTotDCol5
                mrs_Tmp.Fields("Col6").Value = lnTotDCol6
                mrs_Tmp.Fields("Col7").Value = lnTotDCol7
                mrs_Tmp.Fields("Col8").Value = lnTotDCol8
                mrs_Tmp.Update
            Loop
            mrs_Tmp.Sort = "nro desc,Descripcion"
            If lbEsOpenOffice = True Then
                'Abre el archivo ExcelOpenOffice
                lcArchivoExcel = App.Path + "\Plantillas\Hprocedimientos.ods"
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
                'Crea nueva hoja
                Set oExcel = GalenhosExcelApplication()  'New Excel.Application
                Set oWorkBook = oExcel.Workbooks.Add
                'Abre, copia y cierra la plantilla
                Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\Hprocedimientos.xls")
                oWorkBookPlantilla.Worksheets("Hprocedimientos").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Close
                'Activa la primera hoja
                Set oWorkSheet = oWorkBook.Sheets(1)
                mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            End If
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(2, 1).setFormula(ml_Titulo)
                Call Feuille.getcellbyposition(2, 2).setFormula(ml_TextoDelFiltro)
            Else
                'Inicio de Impresion
                oWorkSheet.Cells(2, 3).Value = ml_Titulo
                oWorkSheet.Cells(3, 3).Value = ml_TextoDelFiltro
            End If
            iFila = 7
            iCol = 7
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(iCol + 0, iFila - 2).setFormula(ml_TituloCol1)
                Call Feuille.getcellbyposition(iCol + 1, iFila - 2).setFormula(ml_TituloCol2)
                Call Feuille.getcellbyposition(iCol + 2, iFila - 2).setFormula(ml_TituloCol3)
                Call Feuille.getcellbyposition(iCol + 3, iFila - 2).setFormula(ml_TituloCol4)
                Call Feuille.getcellbyposition(iCol + 4, iFila - 2).setFormula(ml_TituloCol5)
                Call Feuille.getcellbyposition(iCol + 5, iFila - 2).setFormula(ml_TituloCol6)
                Call Feuille.getcellbyposition(iCol + 6, iFila - 2).setFormula(ml_TituloCol7)
                Call Feuille.getcellbyposition(iCol + 7, iFila - 2).setFormula(ml_TituloCol8)
            Else
                oWorkSheet.Cells(iFila - 1, iCol + 1).Value = ml_TituloCol1
                oWorkSheet.Cells(iFila - 1, iCol + 2).Value = ml_TituloCol2
                oWorkSheet.Cells(iFila - 1, iCol + 3).Value = ml_TituloCol3
                oWorkSheet.Cells(iFila - 1, iCol + 4).Value = ml_TituloCol4
                oWorkSheet.Cells(iFila - 1, iCol + 5).Value = ml_TituloCol5
                oWorkSheet.Cells(iFila - 1, iCol + 6).Value = ml_TituloCol6
                oWorkSheet.Cells(iFila - 1, iCol + 7).Value = ml_TituloCol7
                oWorkSheet.Cells(iFila - 1, iCol + 8).Value = ml_TituloCol8
            End If
            lnNumero = 1
            
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(lnNumero)
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula(mrs_Tmp.Fields("Codigo").Value)
                    Call Feuille.getcellbyposition(3, iFila - 1).setFormula(ml_ReglasComunes.OPCsDevuelveCodigoOPCporCodigoCPT(mrs_Tmp.Fields("Codigo").Value))
                    Call Feuille.getcellbyposition(4, iFila - 1).setFormula(mrs_Tmp.Fields("Descripcion").Value)
                    Call Feuille.getcellbyposition(iCol - 2, iFila - 1).setFormula(mrs_Tmp.Fields("nro").Value)
                    Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(IIf(mrs_Tmp.Fields("nro").Value > 0, Round(mrs_Tmp.Fields("nro").Value * 100 / lnTotNro, 2), 0))
                    Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(mrs_Tmp.Fields("Col1").Value)
                    Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(mrs_Tmp.Fields("Col2").Value)
                    Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(mrs_Tmp.Fields("Col3").Value)
                    Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(mrs_Tmp.Fields("Col4").Value)
                    Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(mrs_Tmp.Fields("Col5").Value)
                    Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(mrs_Tmp.Fields("Col6").Value)
                    Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(mrs_Tmp.Fields("Col7").Value)
                    Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(mrs_Tmp.Fields("Col8").Value)
                Else
                    oWorkSheet.Cells(iFila, 2).Value = lnNumero
                    oWorkSheet.Cells(iFila, 3).Value = mrs_Tmp.Fields("Codigo").Value
                    oWorkSheet.Cells(iFila, 4).Value = ml_ReglasComunes.OPCsDevuelveCodigoOPCporCodigoCPT(mrs_Tmp.Fields("Codigo").Value)
                    oWorkSheet.Cells(iFila, 5).Value = mrs_Tmp.Fields("Descripcion").Value
                    oWorkSheet.Cells(iFila, iCol - 1).Value = mrs_Tmp.Fields("nro").Value
                    oWorkSheet.Cells(iFila, iCol).Value = IIf(mrs_Tmp.Fields("nro").Value > 0, Round(mrs_Tmp.Fields("nro").Value * 100 / lnTotNro, 2), 0)
                    oWorkSheet.Cells(iFila, iCol + 1).Value = mrs_Tmp.Fields("Col1").Value
                    oWorkSheet.Cells(iFila, iCol + 2).Value = mrs_Tmp.Fields("Col2").Value
                    oWorkSheet.Cells(iFila, iCol + 3).Value = mrs_Tmp.Fields("Col3").Value
                    oWorkSheet.Cells(iFila, iCol + 4).Value = mrs_Tmp.Fields("Col4").Value
                    oWorkSheet.Cells(iFila, iCol + 5).Value = mrs_Tmp.Fields("Col5").Value
                    oWorkSheet.Cells(iFila, iCol + 6).Value = mrs_Tmp.Fields("Col6").Value
                    oWorkSheet.Cells(iFila, iCol + 7).Value = mrs_Tmp.Fields("Col7").Value
                    oWorkSheet.Cells(iFila, iCol + 8).Value = mrs_Tmp.Fields("Col8").Value
                End If
                If ml_DetallaHC = True Then
                    mrs_Tmp1.Filter = "Codigo='" & mrs_Tmp.Fields("Codigo").Value & "'"
                    mrs_Tmp1.MoveFirst
                    lnTotDNro = 1
                    Do While Not mrs_Tmp1.EOF
                       iFila = iFila + 1
                       If lbEsOpenOffice = True Then
                          Call Feuille.getcellbyposition(1, iFila - 1).setFormula(lnTotDNro & ")    Nº H.C.: " & mrs_Tmp1.Fields("nroHistoria").Value & "   F.Procedim: " & mrs_Tmp1.Fields("fecha").Value & "  Edad: " & mrs_Tmp1.Fields("Edad").Value)
                       Else
                           oWorkSheet.Cells(iFila, iCol + 1).Value = lnTotDNro & ")    Nº H.C.: " & mrs_Tmp1.Fields("nroHistoria").Value & "   F.Procedim: " & mrs_Tmp1.Fields("fecha").Value & "  Edad: " & mrs_Tmp1.Fields("Edad").Value
                       End If
                       lnTotDNro = lnTotDNro + 1
                       mrs_Tmp1.MoveNext
                    Loop
                End If
                iFila = iFila + 1
                lnNumero = lnNumero + 1
                mrs_Tmp.MoveNext
            Loop
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":O" & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Total: ")
                Call Feuille.getcellbyposition(iCol - 2, iFila - 1).setFormula(lnTotNro)
                Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(lnTotCol1)
                Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(lnTotCol2)
                Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(lnTotCol3)
                Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(lnTotCol4)
                Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(lnTotCol5)
                Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(lnTotCol6)
                Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(lnTotCol7)
                Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(lnTotCol8)
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, iCol + 8
                oWorkSheet.Cells(iFila, 2).Value = "Total: "
                oWorkSheet.Cells(iFila, iCol - 1).Value = lnTotNro
                oWorkSheet.Cells(iFila, iCol + 1).Value = lnTotCol1
                oWorkSheet.Cells(iFila, iCol + 2).Value = lnTotCol2
                oWorkSheet.Cells(iFila, iCol + 3).Value = lnTotCol3
                oWorkSheet.Cells(iFila, iCol + 4).Value = lnTotCol4
                oWorkSheet.Cells(iFila, iCol + 5).Value = lnTotCol5
                oWorkSheet.Cells(iFila, iCol + 6).Value = lnTotCol6
                oWorkSheet.Cells(iFila, iCol + 7).Value = lnTotCol7
                oWorkSheet.Cells(iFila, iCol + 8).Value = lnTotCol8
            End If
            Set mrs_Tmp = Nothing
            
            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 1
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 15
                PrintArea(0).EndRow = iFila
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
            Else
                oWorkSheet.PageSetup.PrintTitleRows = "$1:$6"
                    If oWorkSheet.PageSetup.PrintArea <> "" Then
                       oWorkSheet.PageSetup.PrintArea = sighentidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                    End If
                oExcel.Visible = True
                oWorkSheet.PrintPreview
            End If
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
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub
Resume
End Sub


Function FacturacionServiciosSeleccionarCPT(lnTipoSexo As Integer, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set FacturacionServiciosSeleccionarCPT = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServiciosSeleccionarCPT"
        Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@TipoSexo", adInteger, adParamInput, 0, lnTipoSexo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionServiciosSeleccionarCPT = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Sub GenerarRecordsetTemporal()
    With mrs_Tmp
          .Fields.Append "Ord", adInteger, 4, adFldIsNullable
          .Fields.Append "Codigo", adVarChar, 10, adFldIsNullable
          .Fields.Append "Descripcion", adVarChar, 255, adFldIsNullable
          .Fields.Append "Nro", adInteger
          .Fields.Append "Porc", adDouble
          .Fields.Append "Col1", adInteger
          .Fields.Append "Col2", adInteger
          .Fields.Append "Col3", adInteger
          .Fields.Append "Col4", adInteger
          .Fields.Append "Col5", adInteger
          .Fields.Append "Col6", adInteger
          .Fields.Append "Col7", adInteger
          .Fields.Append "Col8", adInteger
          .LockType = adLockOptimistic
          .Open
    End With
    With mrs_Tmp1
          .Fields.Append "Codigo", adVarChar, 10, adFldIsNullable
          .Fields.Append "NroHistoria", adInteger, 4, adFldIsNullable
          .Fields.Append "Fecha", adDate
          .Fields.Append "Edad", adVarChar, 50, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
End Sub



