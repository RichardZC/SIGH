VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptFarICI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************daniel barrantes**************
'***************Filtro de Datos, Configuración y Emisión del Reporte
'***************Historias Clinicas solicitadas por Médico
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
    Dim oFormulario As New FarICI
    oFormulario.Show 1
End Sub

Sub CrearReporte_excel()
Dim oExcel As Excel.Application
Dim oWorkBookPlantilla As Workbook
Dim oWorkBook As Workbook
Dim oWorkSheet As Worksheet
Dim rsReporte As New Recordset
Dim oBuscaPrecio As New Recordset
Dim iFila As Integer
Dim lnNumTotal As Long
Dim lcCodigo As String: Dim lcNombre As String
Dim lnContado As Integer: Dim lnSIS  As Integer: Dim lnSOAT  As Integer: Dim lnConvenios  As Integer: Dim lnCreditoH  As Integer
Dim lnDefensaN As Integer: Dim lnOSDevol  As Integer: Dim lnOSVencim  As Integer: Dim lnExoneracion   As Integer: Dim LnMerma  As Integer: Dim lnIntervS  As Integer
Dim lnPrecio As Double
On Error GoTo ManejadorErrorExcel
    'Filtra los Datos
    Set rsReporte = BuscaDespachosEnFarmacia(ml_IdResponsable, mda_FechaInicio, mda_FechaFin)
    If rsReporte.RecordCount = 0 Then
            MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
            GenerarRecordsetTemporal
            'Carga Datos al TEMPORAL
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
               lnPrecio = 0
               oBuscaPrecio.Open "select PrecioUnitario from FactCatalogoBienesInsumosHosp where idTipoFinanciamiento=1 and idProducto=" & rsReporte.Fields!IdProducto, SIGHComun.CadenaConexion, adOpenKeyset, adLockOptimistic
               If oBuscaPrecio.RecordCount > 0 Then
                  lnPrecio = oBuscaPrecio.Fields!precioUnitario
               End If
               oBuscaPrecio.Close
               lcCodigo = rsReporte.Fields!Codigo
               lcNombre = rsReporte.Fields!nombre
               lnContado = 0: lnSIS = 0: lnSOAT = 0: lnConvenios = 0: lnCreditoH = 0
               lnDefensaN = 0: lnOSDevol = 0: lnOSVencim = 0: lnExoneracion = 0: LnMerma = 0: lnIntervS = 0
               Do While Not rsReporte.EOF And lcCodigo = rsReporte.Fields!Codigo
                    Select Case rsReporte.Fields!IdFormaPago
                    Case 1
                        'lnContado = lnContado + rsReporte.Fields!cantidad
                        lnContado = lnContado + rsReporte.Fields!cantidadPagar
                    Case 2
                        'lnSIS = lnSIS + rsReporte.Fields!cantidadSis
                        lnSIS = lnSIS + rsReporte.Fields!cantidadPagar
                    Case 3
                        'lnSOAT = lnSOAT + rsReporte.Fields!cantidadSoat
                        lnSOAT = lnSOAT + rsReporte.Fields!cantidadPagar
                    Case 4
                        'lnConvenios = lnConvenios + rsReporte.Fields!cantidadConv
                        lnConvenios = lnConvenios + rsReporte.Fields!cantidadPagar
                    Case 5
                        lnCreditoH = lnCreditoH + rsReporte.Fields!cantidadPagar
                    Case 6
                        lnDefensaN = lnDefensaN + rsReporte.Fields!cantidadPagar
                    Case 7
                        lnOSDevol = lnOSDevol + rsReporte.Fields!cantidadPagar
                    Case 8
                        lnOSVencim = lnOSVencim + rsReporte.Fields!cantidadPagar
                    Case 9
                        lnExoneracion = lnExoneracion + rsReporte.Fields!cantidadPagar
                    Case 10
                        LnMerma = LnMerma + rsReporte.Fields!cantidadPagar
                    Case 11
                        lnIntervS = lnIntervS + rsReporte.Fields!cantidadPagar
                    End Select
                    rsReporte.MoveNext
                    If rsReporte.EOF Then
                       Exit Do
                    End If
               Loop
               mrs_Tmp.AddNew
               mrs_Tmp.Fields!Codigo = lcCodigo
               mrs_Tmp.Fields!nombre = lcNombre
               mrs_Tmp.Fields!Precio = lnPrecio
               mrs_Tmp.Fields!Contado = lnContado
               mrs_Tmp.Fields!SIS = lnSIS
               mrs_Tmp.Fields!SOAT = lnSOAT
               mrs_Tmp.Fields!Convenios = lnConvenios
               mrs_Tmp.Fields!CreditoH = lnCreditoH
               mrs_Tmp.Fields!DefensaN = lnDefensaN
               mrs_Tmp.Fields!OSDevol = lnOSDevol
               mrs_Tmp.Fields!OSVencim = lnOSVencim
               mrs_Tmp.Fields!Exoneracion = lnExoneracion
               mrs_Tmp.Fields!OSMerma = LnMerma
               mrs_Tmp.Fields!intervS = lnIntervS
               mrs_Tmp.Update
            Loop
            mrs_Tmp.Sort = "NOMBRE"
             
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            'Crea nueva hoja
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\GalenHos.xls")
            oWorkBookPlantilla.Worksheets("ICI").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
            oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\logotipo.jpg"
            oWorkSheet.Cells(2, 3).Value = ml_TextoDelFiltro
            iFila = 6
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
               oWorkSheet.Cells(iFila, 2).Value = mrs_Tmp.Fields!Codigo
               oWorkSheet.Cells(iFila, 3).Value = mrs_Tmp.Fields!nombre
               oWorkSheet.Cells(iFila, 5).Value = mrs_Tmp.Fields!Precio
               oWorkSheet.Cells(iFila, 6).Value = mrs_Tmp.Fields!Contado
               oWorkSheet.Cells(iFila, 7).Value = mrs_Tmp.Fields!SIS
               oWorkSheet.Cells(iFila, 8).Value = mrs_Tmp.Fields!SOAT
               oWorkSheet.Cells(iFila, 9).Value = mrs_Tmp.Fields!Convenios
               oWorkSheet.Cells(iFila, 10).Value = mrs_Tmp.Fields!CreditoH
               oWorkSheet.Cells(iFila, 11).Value = mrs_Tmp.Fields!DefensaN
               oWorkSheet.Cells(iFila, 12).Value = mrs_Tmp.Fields!OSDevol
               oWorkSheet.Cells(iFila, 13).Value = mrs_Tmp.Fields!OSVencim
               oWorkSheet.Cells(iFila, 14).Value = mrs_Tmp.Fields!OSMerma
               oWorkSheet.Cells(iFila, 15).Value = mrs_Tmp.Fields!Exoneracion
               oWorkSheet.Cells(iFila, 16).Value = mrs_Tmp.Fields!intervS
               mrs_Tmp.MoveNext
               iFila = iFila + 1
            Loop
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 16
            oWorkSheet.Cells(iFila, 3).Value = "Nº Medicamentos: " & Trim(Str(mrs_Tmp.RecordCount))
            oExcel.Visible = True
            oWorkSheet.PrintPreview
            'oWorkSheet.PrintOut
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
          .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "Nombre", adVarChar, 150, adFldIsNullable
          .Fields.Append "Precio", adDouble
          .Fields.Append "Contado", adInteger, 4, adFldIsNullable
          .Fields.Append "SIS", adInteger, 4, adFldIsNullable
          .Fields.Append "SOAT", adInteger, 4, adFldIsNullable
          .Fields.Append "Convenios", adInteger, 4, adFldIsNullable
          .Fields.Append "CreditoH", adInteger, 4, adFldIsNullable
          .Fields.Append "DefensaN", adInteger, 4, adFldIsNullable
          .Fields.Append "OSDevol", adInteger, 4, adFldIsNullable
          .Fields.Append "OSVencim", adInteger, 4, adFldIsNullable
          .Fields.Append "Exoneracion", adInteger, 4, adFldIsNullable
          .Fields.Append "OSMerma", adInteger, 4, adFldIsNullable
          .Fields.Append "IntervS", adInteger, 4, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
End Sub



Function BuscaDespachosEnFarmacia(idFarmacia As Long, lcFechaDespachoInicial As String, lcFechaDespachoFin As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL1 As String: Dim sSQL2 As String
Dim oConexion As New ADODB.Connection
        oConexion.Open SIGHComun.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set BuscaDespachosEnFarmacia = Nothing
        sSQL1 = "SELECT   dbo.FactOrdenesBienesInsumo.IdOrden, dbo.FactOrdenesBienesInsumo.FechaOrden, dbo.FacturacionBienesInsumos.idProducto,dbo.FactOrdenesBienesInsumo.IdFormaPago, " & _
                      " dbo.FactOrdenesBienesInsumo.IdEstadoOrden, dbo.FacturacionBienesInsumos.Cantidad, dbo.FacturacionBienesInsumos.IdTipoFinanciamiento," & _
                      " dbo.FactCatalogoBienesInsumos.Codigo, dbo.FactCatalogoBienesInsumos.Nombre, dbo.FacturacionBienesInsumos.cantidadSIS," & _
                      " dbo.FacturacionBienesInsumos.precioSIS, dbo.FacturacionBienesInsumos.cantidadSOAT, dbo.FacturacionBienesInsumos.precioSOAT," & _
                      " dbo.FacturacionBienesInsumos.cantidadConv, dbo.FacturacionBienesInsumos.precConv, dbo.FacturacionBienesInsumos.ImporteEXO," & _
                      " dbo.FactOrdenesBienesInsumo.IdFormaPago , dbo.FactOrdenesBienesInsumo.idFarmacia,dbo.FacturacionBienesInsumos.cantidadPagar" & _
                " FROM  dbo.FactCatalogoBienesInsumos RIGHT OUTER JOIN" & _
                      " dbo.FacturacionBienesInsumos ON dbo.FactCatalogoBienesInsumos.IdProducto = dbo.FacturacionBienesInsumos.IdProducto LEFT OUTER JOIN" & _
                      " dbo.FactOrdenesBienesInsumo ON dbo.FacturacionBienesInsumos.IdOrden = dbo.FactOrdenesBienesInsumo.IdOrden" & _
                " WHERE dbo.FactOrdenesBienesInsumo.idFarmacia=" & idFarmacia & _
                      " and dbo.FactOrdenesBienesInsumo.FechaOrden Between ('" & lcFechaDespachoInicial & " 00:00:00') and ('" & lcFechaDespachoFin & " 23:59:59')" & _
                " ORDER BY dbo.FactCatalogoBienesInsumos.Codigo"
        With oCommand
        .CommandType = adCmdText
        .CommandTimeout = 150
        Set .ActiveConnection = oConexion
        .CommandText = sSQL1
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    Set BuscaDespachosEnFarmacia = oRecordset
    oConexion.Close
    Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

