VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ms_MensajeError As String
Dim mo_ReporteUtil As New sighEntidades.ReporteUtil
Dim ml_IdGestionCaja As Long
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_AdminServComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mda_FechaInicio As String
Dim mda_FechaFin As String
Dim ml_IdTurno As Long
Dim ml_IdCajero As Long
Dim rsRsTmp As New ADODB.Recordset
Dim rsRsTmp1 As New ADODB.Recordset
Dim ml_TextoDelFiltro  As String
Dim lnIdPagosACuenta As Long
Dim ml_IdComprobantePago As Long
Dim ml_IdTipoFinanciamiento As Long
Dim ml_lnHwnd As Long
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision 'Actualizado 29092014
Dim lcBuscaParametro As New SIGHDatos.Parametros 'Actualizado 29092014
Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos 'Actualizado 29092014

Property Let lnHwnd(lValue As Long)
    ml_lnHwnd = lValue
End Property

Property Let idTipoFinanciamiento(lValue As Long)
    ml_IdTipoFinanciamiento = lValue
End Property

Property Let IdComprobantePago(lValue As Long)
    ml_IdComprobantePago = lValue
End Property

Property Let TextoDelFiltro(lValue As String)
    ml_TextoDelFiltro = lValue
End Property

Property Let IdCajero(lValue As Long)
    ml_IdCajero = lValue
End Property
Property Let IdTurno(lValue As Long)
    ml_IdTurno = lValue
End Property
Property Let FechaInicio(daValue As String)
    mda_FechaInicio = daValue
End Property
Property Let FechaFin(daValue As String)
    mda_FechaFin = daValue
End Property


Property Let IdGestionCaja(lValue As Long)
    ml_IdGestionCaja = lValue
End Property

Property Get IdGestionCaja() As Long
    IdGestionCaja = ml_IdGestionCaja
End Property

Sub ImprimirComprobantePagoServicios(oDOPaciente As doPaciente, oDOAtencion As DOAtencion, oDOFactOrdenServicio As DOFactOrdenServicio, oDOComprobantePago As DOCajaComprobantesPago, rsFacturacionProductos As Recordset)
Dim oRecibos As New RecibosBoleta
    
    oRecibos.Imprimir oDOPaciente, oDOAtencion, oDOComprobantePago, rsFacturacionProductos
    oRecibos.Show 1
    Unload oRecibos
    
End Sub

Sub ImprimirComprobantePagoBienesInsumos(oDOPaciente As doPaciente, oDOAtencion As DOAtencion, oDOFactOrdenBienInsumo As DOFactOrdenBienInsumo, oDOComprobantePago As DOCajaComprobantesPago, rsFacturacionProductos As Recordset)
Dim oRecibos As New RecibosBoleta
    
    oRecibos.Imprimir oDOPaciente, oDOAtencion, oDOComprobantePago, rsFacturacionProductos
    Unload oRecibos
End Sub


'***************daniel barrantes**************
'***************reporte PARTE DIARIO-CAJA
'***************
Sub CrearParteDiario(lnHwnd As Long)
Dim rsreporte As New Recordset
Dim rsReporte1 As New Recordset
Dim lnImpSubTot As Double: Dim lntImpSubTot As Double
Dim lnImpAnul As Double: Dim lntImpAnul As Double
Dim lnImpExo As Double: Dim lntImpExo As Double
Dim lnImpDevol As Double: Dim lntImpDevol As Double
Dim lnImpPagCta As Double: Dim lntImpPagCta As Double
Dim lnImpTot As Double: Dim lntImpTot As Double
Dim lnIdPartida As Long: Dim lcPartida As String: Dim lcDpartida As String
Dim iFila As Long
Dim lRecordCount As Long
Dim lbNuevo As Boolean
Dim lbSeguir As Boolean: Dim lcDocum As String: Dim lcBuscar As String: Dim lnDctos As Double
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lcSql As String, lbEsOpenOffice As Boolean
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
        Dim oRange As range
        Dim range As Excel.range
        Dim borders As Excel.borders
    End If
    
    Const COL_CODIGO = 4
    Const COL_DESCRIPCION = 10
    Const COL_SUBTOTAL = 37
    Const COL_ANULADO = 43
    Const COL_EXONERADO = 49
    Const COL_DEVOLUCION = 55
    Const COL_PAGOCTA = 61
    Const COL_TOTAL = 67
    '
    lnIdPagosACuenta = Val(lcBuscaParametro.SeleccionaFilaParametro(245))

        GenerarRecordsetTemporal
        GenerarRecordsetDctos
        
        lRecordCount = 0: lntImpSubTot = 0: lntImpAnul = 0: lntImpExo = 0: lntImpDevol = 0: lntImpPagCta = 0: lntImpTot = 0
        'Servicios
        'Set rsReporte = ParteDiarioServiciosDEBB(ml_IdGestionCaja, mda_FechaInicio, mda_FechaFin, ml_IdTurno, ml_IdCajero)
        Set rsreporte = mo_ReglasFacturacion.ServicioParteDiario(ml_IdGestionCaja, CDate(mda_FechaInicio), CDate(mda_FechaFin), ml_IdTurno, ml_IdCajero)
        If rsreporte.RecordCount > 0 Then
            rsreporte.MoveFirst
            Do While Not rsreporte.EOF
               lnImpAnul = 0: lnImpDevol = 0
               lnImpTot = 0
               lnImpExo = 0:  lnImpPagCta = 0
               lnIdPartida = rsreporte!IdPartidaPresupuestal
               lcPartida = rsreporte!partida
               lcDpartida = rsreporte!dpartida
               Do While Not rsreporte.EOF And lnIdPartida = rsreporte!IdPartidaPresupuestal
                        'lnImpExo = lnImpExo + IIf(IsNull(rsReporte!importeEXO), 0, rsReporte!importeEXO)             'Exonerado
                        
                        Select Case rsreporte!idEstadoComprobante
                        Case 6                                               'Devolucion
                           lnImpDevol = lnImpDevol + rsreporte!Total
                        Case 9                                               'Anulado
                           lnImpAnul = lnImpAnul + rsreporte!Total
                        Case 4
                           If rsreporte!idProducto = lnIdPagosACuenta Then                 'Pago a cuenta
                              lnImpPagCta = lnImpPagCta + rsreporte!Total
                              lnImpTot = lnImpTot + rsreporte!Total
                           Else
                              lnImpTot = lnImpTot + rsreporte!Total     'solo pago
                           End If
                        End Select
                        
                        
                        'Suma Descuentos de la CABECERA - COMPROBANTE PAGO
                        If Not IsNull(rsreporte.Fields!Adelantos) Then
                            lbSeguir = True
                            lcDocum = rsreporte.Fields!nroSerie + rsreporte.Fields!nrodocumento & "'"
                            If rsRsTmp1.RecordCount > 0 Then
                               rsRsTmp1.MoveFirst
                               Do While Not rsRsTmp1.EOF
                                    If rsRsTmp1.Fields!Docume = lcDocum Then
                                       Exit Do
                                    End If
                                    rsRsTmp1.MoveNext
                               Loop
                               lbSeguir = IIf(rsRsTmp1.EOF, True, False)
                            End If
                            If lbSeguir Then
                                rsRsTmp1.AddNew
                                rsRsTmp1.Fields!Docume = lcDocum
                                rsRsTmp1.Fields!Adelantos = rsreporte.Fields!Adelantos
                                rsRsTmp1.Update
                            End If
                        End If
                        
                        rsreporte.MoveNext
                        If rsreporte.EOF Then
                           Exit Do
                        End If
               Loop
               lnImpSubTot = lnImpTot + lnImpAnul + lnImpDevol
               
               lntImpSubTot = lntImpSubTot + lnImpSubTot
               lntImpExo = lntImpExo + lnImpExo
               lntImpAnul = lntImpAnul + lnImpAnul
               lntImpDevol = lntImpDevol + lnImpDevol
               lntImpTot = lntImpTot + lnImpTot
               lntImpPagCta = lntImpPagCta + lnImpPagCta
               
               rsRsTmp.AddNew
               rsRsTmp.Fields("idPartida").Value = lnIdPartida
               rsRsTmp.Fields("Partida").Value = lcPartida
               rsRsTmp.Fields("dPartida").Value = lcDpartida
               rsRsTmp.Fields("SubTotal").Value = lnImpSubTot
               rsRsTmp.Fields("Exonerado").Value = lnImpExo
               rsRsTmp.Fields("Anulado").Value = lnImpAnul
               rsRsTmp.Fields("Devolucion").Value = lnImpDevol
               rsRsTmp.Fields("PagoCta").Value = lnImpPagCta
               rsRsTmp.Fields("Total").Value = lnImpTot
               rsRsTmp.Update
            Loop
        End If
        'Farmacia
        'Set rsReporte = ParteDiarioFarmaciaDEBB(ml_IdGestionCaja, mda_FechaInicio, mda_FechaFin, ml_IdTurno, ml_IdCajero)
        Set rsreporte = mo_ReglasFacturacion.FarmaciaParteDiario(ml_IdGestionCaja, CDate(mda_FechaInicio), CDate(mda_FechaFin), ml_IdTurno, ml_IdCajero)
        If rsreporte.RecordCount > 0 Then
            
            rsreporte.MoveFirst
            Do While Not rsreporte.EOF
               lnImpAnul = 0: lnImpDevol = 0
               lnImpTot = 0
               lnImpExo = 0:  lnImpPagCta = 0
               lnIdPartida = rsreporte!IdPartidaPresupuestal
               lcPartida = rsreporte!partida
               lcDpartida = rsreporte!dpartida
               Do While Not rsreporte.EOF And lnIdPartida = rsreporte!IdPartidaPresupuestal
                        'lnImpExo = lnImpExo + IIf(IsNull(rsReporte!importeEXO), 0, rsReporte!importeEXO)             'Exonerado
                        
                        Select Case rsreporte!idEstadoComprobante
                        Case 6                                               'Devolucion
                           lnImpDevol = lnImpDevol + rsreporte!TotalPagar
                        Case 9                                               'Anulado
                           lnImpAnul = lnImpAnul + rsreporte!TotalPagar
                        Case 4
                           If rsreporte!idProducto = lnIdPagosACuenta Then                 'Pago a cuenta
                              lnImpPagCta = lnImpPagCta + rsreporte!TotalPagar
                              lnImpTot = lnImpTot + rsreporte!TotalPagar
                           Else
                              lnImpTot = lnImpTot + rsreporte!TotalPagar     'solo pago
                           End If
                        End Select
                        
                        rsreporte.MoveNext
                        If rsreporte.EOF Then
                           Exit Do
                        End If
               Loop
               lnImpSubTot = lnImpTot + lnImpAnul + lnImpDevol
               
               lntImpSubTot = lntImpSubTot + lnImpSubTot
               lntImpExo = lntImpExo + lnImpExo
               lntImpAnul = lntImpAnul + lnImpAnul
               lntImpDevol = lntImpDevol + lnImpDevol
               lntImpTot = lntImpTot + lnImpTot
               lntImpPagCta = lntImpPagCta + lnImpPagCta
               lbNuevo = True
               If rsRsTmp.RecordCount > 0 Then
                  rsRsTmp.MoveFirst
                  rsRsTmp.Find "idPartida=" & lnIdPartida
                  If Not rsRsTmp.EOF Then
                     lbNuevo = False
                  End If
               End If
               If lbNuevo Then
                    rsRsTmp.AddNew
                    rsRsTmp.Fields("idPartida").Value = lnIdPartida
                    rsRsTmp.Fields("Partida").Value = lcPartida
                    rsRsTmp.Fields("dPartida").Value = lcDpartida
               End If
               rsRsTmp.Fields("SubTotal").Value = rsRsTmp.Fields("SubTotal").Value + lnImpSubTot
               rsRsTmp.Fields("Exonerado").Value = rsRsTmp.Fields("Exonerado").Value + lnImpExo
               rsRsTmp.Fields("Anulado").Value = rsRsTmp.Fields("Anulado").Value + lnImpAnul
               rsRsTmp.Fields("Devolucion").Value = rsRsTmp.Fields("Devolucion").Value + lnImpDevol
               rsRsTmp.Fields("PagoCta").Value = rsRsTmp.Fields("PagoCta").Value + lnImpPagCta
               rsRsTmp.Fields("Total").Value = rsRsTmp.Fields("Total").Value + lnImpTot
               rsRsTmp.Update
            Loop
        End If
        
        If rsRsTmp.RecordCount = 0 Then
            MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
        Else
            'Acumula descuentos
            lnDctos = 0
            If rsRsTmp1.RecordCount > 0 Then
               rsRsTmp1.MoveFirst
               Do While Not rsRsTmp1.EOF
                  lnDctos = lnDctos + rsRsTmp1.Fields!Adelantos
                  rsRsTmp1.MoveNext
               Loop
            End If
        
        
            rsRsTmp.Sort = "dpartida"
            
            
            If lbEsOpenOffice = True Then
                'Abre el archivo ExcelOpenOffice
                lcArchivoExcel = App.Path + "\Plantillas\ECajaParteDiario.ods"
        '        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
        '        Chemin = "file:///" & App.Path & "\Plantillas\"
        '        Chemin = Replace(Chemin, "\", "/")
        '        Fichier = Chemin & "/OpenOffice.ods"
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
                Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\ECajaParteDiario.xls")
                oWorkBookPlantilla.Worksheets("CajaParteDiario").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Close
                'Activa la primera hoja
                Set oWorkSheet = oWorkBook.Sheets(1)
                mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            End If
            
            If lbEsOpenOffice = True Then
               Call Feuille.getcellbyposition(11, 3).setFormula(ml_TextoDelFiltro)
            Else
               oWorkSheet.Cells(3, 11).Value = ml_TextoDelFiltro
            End If
            lRecordCount = rsRsTmp.RecordCount
            iFila = 7
            rsRsTmp.MoveFirst
            Do While Not rsRsTmp.EOF
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(COL_CODIGO, iFila - 1).setFormula(rsRsTmp!partida)
                    Call Feuille.getcellbyposition(COL_DESCRIPCION, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsRsTmp!dpartida))
                    Call Feuille.getcellbyposition(COL_SUBTOTAL, iFila - 1).setFormula(rsRsTmp!Subtotal)
                    Call Feuille.getcellbyposition(COL_EXONERADO, iFila - 1).setFormula(rsRsTmp!exonerado)
                    Call Feuille.getcellbyposition(COL_ANULADO, iFila - 1).setFormula(rsRsTmp!anulado)
                    Call Feuille.getcellbyposition(COL_DEVOLUCION, iFila - 1).setFormula(rsRsTmp!Devolucion)
                    Call Feuille.getcellbyposition(COL_PAGOCTA, iFila - 1).setFormula(rsRsTmp!PagoCta)
                    Call Feuille.getcellbyposition(COL_TOTAL, iFila - 1).setFormula(rsRsTmp!Total)
                Else
                    oWorkSheet.Cells(iFila, COL_CODIGO).Value = rsRsTmp!partida
                    oWorkSheet.Cells(iFila, COL_DESCRIPCION).Value = mo_ReporteUtil.NullToVacio(rsRsTmp!dpartida)
                    oWorkSheet.Cells(iFila, COL_SUBTOTAL).Value = rsRsTmp!Subtotal
                    oWorkSheet.Cells(iFila, COL_EXONERADO).Value = rsRsTmp!exonerado
                    oWorkSheet.Cells(iFila, COL_ANULADO).Value = rsRsTmp!anulado
                    oWorkSheet.Cells(iFila, COL_DEVOLUCION).Value = rsRsTmp!Devolucion
                    oWorkSheet.Cells(iFila, COL_PAGOCTA).Value = rsRsTmp!PagoCta
                    oWorkSheet.Cells(iFila, COL_TOTAL).Value = rsRsTmp!Total
                End If
                iFila = iFila + 1
                rsRsTmp.MoveNext
            Loop
            iFila = iFila + 1
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName(mo_AdminReportes.BuscaNombreColumna(COL_CODIGO) & CStr(iFila) & ":" & mo_AdminReportes.BuscaNombreColumna(COL_TOTAL + 5) & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            
                Call Feuille.getcellbyposition(COL_DESCRIPCION, iFila - 1).setFormula("Cantidad de Partidas: " + Trim(Str(lRecordCount)))
                Call Feuille.getcellbyposition(COL_SUBTOTAL, iFila - 1).setFormula(lntImpSubTot)
                Call Feuille.getcellbyposition(COL_EXONERADO, iFila - 1).setFormula(lntImpExo)
                Call Feuille.getcellbyposition(COL_ANULADO, iFila - 1).setFormula(lntImpAnul)
                Call Feuille.getcellbyposition(COL_DEVOLUCION, iFila - 1).setFormula(lntImpDevol)
                Call Feuille.getcellbyposition(COL_PAGOCTA, iFila - 1).setFormula(lntImpPagCta)
                Call Feuille.getcellbyposition(COL_TOTAL, iFila - 1).setFormula(lntImpTot)
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, COL_CODIGO, iFila, COL_TOTAL + 5
                oWorkSheet.Cells(iFila, COL_DESCRIPCION).Value = "Cantidad de Partidas: " + Trim(Str(lRecordCount))
                oWorkSheet.Cells(iFila, COL_SUBTOTAL).Value = lntImpSubTot
                oWorkSheet.Cells(iFila, COL_EXONERADO).Value = lntImpExo
                oWorkSheet.Cells(iFila, COL_ANULADO).Value = lntImpAnul
                oWorkSheet.Cells(iFila, COL_DEVOLUCION).Value = lntImpDevol
                oWorkSheet.Cells(iFila, COL_PAGOCTA).Value = lntImpPagCta
                oWorkSheet.Cells(iFila, COL_TOTAL).Value = lntImpTot
            End If
            iFila = iFila + 1
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(COL_PAGOCTA, iFila - 1).setFormula("Dctos")
                Call Feuille.getcellbyposition(COL_TOTAL, iFila - 1).setFormula(lnDctos)
            Else
                oWorkSheet.Cells(iFila, COL_PAGOCTA).Value = "Dctos"
                oWorkSheet.Cells(iFila, COL_TOTAL).Value = lnDctos
            End If
            iFila = iFila + 1
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("B" & CStr(8) & ":G" & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Call Feuille.getcellbyposition(COL_TOTAL, iFila - 1).setFormula(lntImpTot - lnDctos)
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, COL_PAGOCTA, iFila, COL_TOTAL + 5
                oWorkSheet.Cells(iFila, COL_TOTAL).Value = lntImpTot - lnDctos
            End If
            '
            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 1
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 67
                PrintArea(0).EndRow = iFila
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.GetFrame.getContainerWindow.SetVisible(True)
                Call Document.getCurrentController.GetFrame.getComponentWindow.SetVisible(True)
                MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
            Else
                oWorkSheet.PageSetup.PrintTitleRows = "$2:$6"
                If oWorkSheet.PageSetup.PrintArea <> "" Then
                   oWorkSheet.PageSetup.PrintArea = sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                End If
                oExcel.Visible = True
                oWorkSheet.PrintPreview
                'oWorkSheet.PrintOut
                'oWorkBook.Close SaveChanges:=False
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
            'liberar memoria
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
End Sub

'***************daniel barrantes**************
'***************Reporte: CONSOLIDADO DE RECAUDACION
'***************
Sub CrearConsolidadoRecaudacion()
Dim oExcel As Excel.Application
Dim oWorkBookPlantilla As Workbook
Dim oWorkBook As Workbook
Dim oWorkSheet As Worksheet
Dim rsreporte As New Recordset
Dim rsReporte1 As New Recordset
Dim lnImpSubTot As Double: Dim lntImpSubTot As Double
Dim lnImpAnul As Double: Dim lntImpAnul As Double
Dim lnImpExo As Double: Dim lntImpExo As Double
Dim lnImpDevol As Double: Dim lntImpDevol As Double
Dim lnImpPagCta As Double: Dim lntImpPagCta As Double
Dim lnImpTot As Double: Dim lntImpTot As Double
Dim lnIdPartida As Long: Dim lcPartida As String: Dim lcDpartida As String
Dim iFila As Long
Dim lRecordCount As Long
Dim lbNuevo As Boolean
Dim lbSeguir As Boolean: Dim lcDocum As String: Dim lcBuscar As String: Dim lnDctos As Double
Dim lcBuscaParametro As New SIGHDatos.Parametros
On Error GoTo ManejadorError
    
'        Const COL_CODIGO = 4
'        Const COL_DESCRIPCION = 10
'        Const COL_SUBTOTAL = 37
'        Const COL_ANULADO = 43
'        Const COL_EXONERADO = 49
'        Const COL_DEVOLUCION = 55
'        Const COL_PAGOCTA = 61
'        'Const COL_TOTAL = 67
'        Const COL_TOTAL = 61
        Const COL_CODIGO = 4
        Const COL_DESCRIPCION = 10
        Const COL_SUBTOTAL = 37
        Const COL_ANULADO = 43
        Const COL_EXONERADO = 49
        Const COL_DEVOLUCION = 55
        Const COL_PAGOCTA = 61
        Const COL_TOTAL = 67
        
        GenerarRecordsetTemporal
        GenerarRecordsetDctos

        lRecordCount = 0: lntImpSubTot = 0: lntImpAnul = 0: lntImpExo = 0: lntImpDevol = 0: lntImpPagCta = 0: lntImpTot = 0
        'Servicios
        'Set rsReporte = ParteDiarioServiciosDEBB(ml_IdGestionCaja, mda_FechaInicio, mda_FechaFin, ml_IdTurno, ml_IdCajero)
        Set rsreporte = mo_ReglasFacturacion.ServicioParteDiario(ml_IdGestionCaja, CDate(mda_FechaInicio), CDate(mda_FechaFin), ml_IdTurno, ml_IdCajero)
        If rsreporte.RecordCount > 0 Then
            rsreporte.MoveFirst
            Do While Not rsreporte.EOF
               lnImpAnul = 0: lnImpDevol = 0
               lnImpTot = 0
               lnImpExo = 0:  lnImpPagCta = 0
               lnIdPartida = rsreporte!IdPartidaPresupuestal
               lcPartida = rsreporte!partida
               lcDpartida = rsreporte!dpartida
               Do While Not rsreporte.EOF And lnIdPartida = rsreporte!IdPartidaPresupuestal
                        'lnImpExo = lnImpExo + IIf(IsNull(rsReporte!importeEXO), 0, rsReporte!importeEXO)             'Exonerado
                        
                        Select Case rsreporte!idEstadoComprobante
                        Case 6                                               'Devolucion
                           lnImpDevol = lnImpDevol + rsreporte!Total
                        Case 9                                               'Anulado
                           lnImpAnul = lnImpAnul + rsreporte!Total
                        Case 4
                           If rsreporte!idProducto = lnIdPagosACuenta Then                 'Pago a cuenta
                              lnImpPagCta = lnImpPagCta + rsreporte!Total
                              lnImpTot = lnImpTot + rsreporte!Total
                           Else
                              lnImpTot = lnImpTot + rsreporte!Total     'solo pago
                           End If
                        End Select
                        
                        
                        'Suma Descuentos de la CABECERA (COMPROBANTE PAGO)
                        If Not IsNull(rsreporte.Fields!Adelantos) Then
                            lbSeguir = True
                            lcDocum = rsreporte.Fields!nroSerie + rsreporte.Fields!nrodocumento & "'"
                            If rsRsTmp1.RecordCount > 0 Then
                               rsRsTmp1.MoveFirst
                               Do While Not rsRsTmp1.EOF
                                    If rsRsTmp1.Fields!Docume = lcDocum Then
                                       Exit Do
                                    End If
                                    rsRsTmp1.MoveNext
                               Loop
                               lbSeguir = IIf(rsRsTmp1.EOF, True, False)
                            End If
                            If lbSeguir Then
                                rsRsTmp1.AddNew
                                rsRsTmp1.Fields!Docume = lcDocum
                                rsRsTmp1.Fields!Adelantos = rsreporte.Fields!Adelantos
                                rsRsTmp1.Update
                            End If
                        End If
                        
                        
                        rsreporte.MoveNext
                        If rsreporte.EOF Then
                           Exit Do
                        End If
               Loop
               lnImpSubTot = lnImpTot + lnImpAnul + lnImpDevol
               
               lntImpSubTot = lntImpSubTot + lnImpSubTot
               lntImpExo = lntImpExo + lnImpExo
               lntImpAnul = lntImpAnul + lnImpAnul
               lntImpDevol = lntImpDevol + lnImpDevol
               lntImpTot = lntImpTot + lnImpTot
               lntImpPagCta = lntImpPagCta + lnImpPagCta
               
               rsRsTmp.AddNew
               rsRsTmp.Fields("idPartida").Value = lnIdPartida
               rsRsTmp.Fields("Partida").Value = lcPartida
               rsRsTmp.Fields("dPartida").Value = lcDpartida
               rsRsTmp.Fields("SubTotal").Value = lnImpSubTot
               rsRsTmp.Fields("Exonerado").Value = lnImpExo
               rsRsTmp.Fields("Anulado").Value = lnImpAnul
               rsRsTmp.Fields("Devolucion").Value = lnImpDevol
               rsRsTmp.Fields("PagoCta").Value = lnImpPagCta
               rsRsTmp.Fields("Total").Value = lnImpTot
               rsRsTmp.Update
            Loop
        End If
        'Farmacia
        'Set rsReporte = ParteDiarioFarmaciaDEBB(ml_IdGestionCaja, mda_FechaInicio, mda_FechaFin, ml_IdTurno, ml_IdCajero)
        Set rsreporte = mo_ReglasFacturacion.FarmaciaParteDiario(ml_IdGestionCaja, CDate(mda_FechaInicio), CDate(mda_FechaFin), ml_IdTurno, ml_IdCajero)
        If rsreporte.RecordCount > 0 Then
            rsreporte.MoveFirst
            Do While Not rsreporte.EOF
               lnImpAnul = 0: lnImpDevol = 0
               lnImpTot = 0
               lnImpExo = 0:  lnImpPagCta = 0
               lnIdPartida = rsreporte!IdPartidaPresupuestal
               lcPartida = rsreporte!partida
               lcDpartida = rsreporte!dpartida
               Do While Not rsreporte.EOF And lnIdPartida = rsreporte!IdPartidaPresupuestal
                        'lnImpExo = lnImpExo + IIf(IsNull(rsReporte!importeEXO), 0, rsReporte!importeEXO)             'Exonerado
                        
                        Select Case rsreporte!idEstadoComprobante
                        Case 6                                               'Devolucion
                           lnImpDevol = lnImpDevol + rsreporte!TotalPagar
                        Case 9                                               'Anulado
                           lnImpAnul = lnImpAnul + rsreporte!TotalPagar
                        Case 4
                           If rsreporte!idProducto = lnIdPagosACuenta Then                 'Pago a cuenta
                              lnImpPagCta = lnImpPagCta + rsreporte!TotalPagar
                              lnImpTot = lnImpTot + rsreporte!TotalPagar
                           Else
                              lnImpTot = lnImpTot + rsreporte!TotalPagar     'solo pago
                           End If
                        End Select
                        
                        rsreporte.MoveNext
                        If rsreporte.EOF Then
                           Exit Do
                        End If
               Loop
               lnImpSubTot = lnImpTot + lnImpAnul + lnImpDevol
               
               lntImpSubTot = lntImpSubTot + lnImpSubTot
               lntImpExo = lntImpExo + lnImpExo
               lntImpAnul = lntImpAnul + lnImpAnul
               lntImpDevol = lntImpDevol + lnImpDevol
               lntImpTot = lntImpTot + lnImpTot
               lntImpPagCta = lntImpPagCta + lnImpPagCta
               lbNuevo = True
               If rsRsTmp.RecordCount > 0 Then
                  rsRsTmp.MoveFirst
                  rsRsTmp.Find "idPartida=" & lnIdPartida
                  If Not rsRsTmp.EOF Then
                     lbNuevo = False
                  End If
               End If
               If lbNuevo Then
                    rsRsTmp.AddNew
                    rsRsTmp.Fields("idPartida").Value = lnIdPartida
                    rsRsTmp.Fields("Partida").Value = lcPartida
                    rsRsTmp.Fields("dPartida").Value = lcDpartida
               End If
               rsRsTmp.Fields("SubTotal").Value = rsRsTmp.Fields("SubTotal").Value + lnImpSubTot
               rsRsTmp.Fields("Exonerado").Value = rsRsTmp.Fields("Exonerado").Value + lnImpExo
               rsRsTmp.Fields("Anulado").Value = rsRsTmp.Fields("Anulado").Value + lnImpAnul
               rsRsTmp.Fields("Devolucion").Value = rsRsTmp.Fields("Devolucion").Value + lnImpDevol
               rsRsTmp.Fields("PagoCta").Value = rsRsTmp.Fields("PagoCta").Value + lnImpPagCta
               rsRsTmp.Fields("Total").Value = rsRsTmp.Fields("Total").Value + lnImpTot
               rsRsTmp.Update
            Loop
        End If
        
        If rsRsTmp.RecordCount = 0 Then
           MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
        Else
            'Acumula descuentos
            lnDctos = 0
            If rsRsTmp1.RecordCount > 0 Then
               rsRsTmp1.MoveFirst
               Do While Not rsRsTmp1.EOF
                  lnDctos = lnDctos + rsRsTmp1.Fields!Adelantos
                  rsRsTmp1.MoveNext
               Loop
            End If
        
            rsRsTmp.Sort = "dpartida"
            'Crea nueva hoja
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\ECajaParteDiario.xls")
            oWorkBookPlantilla.Worksheets("CajaParteDiario").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            
'            'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
'            oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"

            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            
            oWorkSheet.Cells(2, 23).Value = "CONSOLIDADO DE RECAUDACION"
            oWorkSheet.Cells(3, 11).Value = ml_TextoDelFiltro
            
            lRecordCount = rsRsTmp.RecordCount
            iFila = 7
            rsRsTmp.MoveFirst
            Do While Not rsRsTmp.EOF
                oWorkSheet.Cells(iFila, COL_CODIGO).Value = rsRsTmp!partida
                oWorkSheet.Cells(iFila, COL_DESCRIPCION).Value = mo_ReporteUtil.NullToVacio(rsRsTmp!dpartida)
                oWorkSheet.Cells(iFila, COL_SUBTOTAL).Value = rsRsTmp!Subtotal
                oWorkSheet.Cells(iFila, COL_EXONERADO).Value = rsRsTmp!exonerado
                oWorkSheet.Cells(iFila, COL_ANULADO).Value = rsRsTmp!anulado
                oWorkSheet.Cells(iFila, COL_DEVOLUCION).Value = rsRsTmp!Devolucion
                oWorkSheet.Cells(iFila, COL_PAGOCTA).Value = rsRsTmp!PagoCta
                oWorkSheet.Cells(iFila, COL_TOTAL).Value = rsRsTmp!Total
                iFila = iFila + 1
                rsRsTmp.MoveNext
            Loop
            iFila = iFila + 1
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, COL_CODIGO, iFila, COL_TOTAL + 5
            oWorkSheet.Cells(iFila, COL_DESCRIPCION).Value = "Cantidad de Partidas: " + Trim(Str(lRecordCount))
            oWorkSheet.Cells(iFila, COL_SUBTOTAL).Value = lntImpSubTot
            oWorkSheet.Cells(iFila, COL_EXONERADO).Value = lntImpExo
            oWorkSheet.Cells(iFila, COL_ANULADO).Value = lntImpAnul
            oWorkSheet.Cells(iFila, COL_DEVOLUCION).Value = lntImpDevol
            oWorkSheet.Cells(iFila, COL_PAGOCTA).Value = lntImpPagCta
            oWorkSheet.Cells(iFila, COL_TOTAL).Value = lntImpTot
            iFila = iFila + 1
            oWorkSheet.Cells(iFila, COL_DESCRIPCION).Value = "Dctos"
            oWorkSheet.Cells(iFila, COL_TOTAL).Value = lnDctos
            iFila = iFila + 1
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, COL_PAGOCTA, iFila, COL_TOTAL + 5
            oWorkSheet.Cells(iFila, COL_TOTAL).Value = lntImpTot - lnDctos
            '
            oWorkSheet.PageSetup.PrintTitleRows = "$2:$6"
            If oWorkSheet.PageSetup.PrintArea <> "" Then
               oWorkSheet.PageSetup.PrintArea = sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
            End If
            oExcel.Visible = True
            oWorkSheet.PrintPreview
            'oWorkSheet.PrintOut
            'oWorkBook.Close SaveChanges:=False
        End If
    Set oWorkSheet = Nothing
    Set oExcel = Nothing
Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub
End Sub


'***************daniel barrantes**************
'***************Reporte: CONSOLIADADO DE SERVICIO
'***************
Sub CrearReporteConsolidadoServicios(lbSoloCreditos As Boolean)
Dim lcLlave As String, lcFilter As String
Dim lnImpSubTot As Double: Dim lntImpSubTot As Double
Dim lnImpAnul As Double: Dim lntImpAnul As Double
Dim lnImpExo As Double: Dim lntImpExo As Double
Dim lnImpDevol As Double: Dim lntImpDevol As Double
Dim lnImpPagCta As Double: Dim lntImpPagCta As Double
Dim lnImpTot As Double: Dim lntImpTot As Double: Dim lnDctos As Double
Dim iFila As Long, lnImpRedondeo As Double, lntImpRedondeo As Double
Dim lRecordCount As Long, lcCadenaConexion As String
Dim iFilaEmpiezaNota As Long

Dim rsreporte As New Recordset
Dim rsNotasCredito As New Recordset
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasReportes As New ReglasReportes
Dim lnNotaCredTot As Double

Dim CantidadSOAT As Long: Dim PrecioSOAT As Double
Dim lbEsOpenOffice As Boolean
Dim lcSql As String
Dim oConexion As New Connection, lbTienePagoAcuenta As Boolean, lnRedondeo As Double
Dim lbVerResumenNotaCredito As Boolean
oConexion.CommandTimeout = 300
oConexion.CursorLocation = adUseClient
oConexion.Open sighEntidades.CadenaConexion

lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)

'On Error GoTo ManejadorError
    
    Const COL_FECHA = 2
    Const COL_USUARIO = 3
    Const COL_BOLETA = 4
    Const COL_NRO_HISTORIA = 5
    Const COL_RAZON_SOCIAL = 6
    Const COL_SUBTOTAL = 7
    Const COL_REDONDEO = 8
    Const COL_ANULADO = 9
    Const COL_EXONERADO = 10
    Const COL_DEVOLUCION = 11
    Const COL_PAGOCTA = 12
    Const COL_TOTAL = 13
    
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
        Dim lnHwnd As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
        Dim oRange As range
        Dim range As Excel.range
        Dim borders As Excel.borders
        Dim range_row As Excel.range
    End If
        
        lnIdPagosACuenta = Val(lcBuscaParametro.SeleccionaFilaParametro(245))
        lcCadenaConexion = sighEntidades.CadenaConexion
    
        GenerarRecordsetTemporal1
        
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\ECajaConsolidadoServ.ods"
    '        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
    '        Chemin = "file:///" & App.Path & "\Plantillas\"
    '        Chemin = Replace(Chemin, "\", "/")
    '        Fichier = Chemin & "/OpenOffice.ods"
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
            Call Feuille.getcellbyposition(4, 1).setFormula("CONSOLIDADO DE SERVICIOS")
            'Call Feuille.getcellbyposition(11, 1).setFormula(lcBuscaParametro.RetornaFechaHoraServidorSQL)
            Call Feuille.getcellbyposition(2, 1).setFormula(ml_TextoDelFiltro)
        Else
            'Crea nueva hoja
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\ECajaConsolidadoServ.xls")
            oWorkBookPlantilla.Worksheets("CajaConsolidadoServ").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            oWorkSheet.Cells(2, 5).Value = "CONSOLIDADO DE SERVICIOS"
            'oWorkSheet.Cells(2, 12).Value = lcBuscaParametro.RetornaFechaHoraServidorSQL
            oWorkSheet.Cells(4, 3).Value = ml_TextoDelFiltro
        End If
        'SERVICIOS emitidos en CAJA SERVICIO
        Set rsreporte = mo_ReglasFacturacion.ServicioConsolidado(ml_IdGestionCaja, CDate(mda_FechaInicio), CDate(mda_FechaFin), ml_IdTurno, ml_IdCajero)
        
        lcFilter = ""
        If ml_IdComprobantePago > 0 Then
            lcFilter = "IdTipoComprobante=" & ml_IdComprobantePago & " and "
        End If
        If ml_IdTipoFinanciamiento > 0 Then
           lcFilter = lcFilter & "idTipoFinanciamiento=" & ml_IdTipoFinanciamiento & " and "
        End If
        lcFilter = lcFilter & IIf(lbSoloCreditos = True, "TieneCredito<>null", "TieneCredito=null") & " and "
        If lcFilter <> "" Then
           lcFilter = Left(lcFilter, Len(lcFilter) - 4)
           rsreporte.Filter = lcFilter
        End If
        If rsreporte.RecordCount > 0 Then
            With rsRsTmp
                 rsreporte.MoveFirst
                 Do While Not rsreporte.EOF
'If Val(rsReporte!nroDocumento) = 247266 Then
'lbTienePagoAcuenta = False
'End If
                    lbTienePagoAcuenta = mo_ReglasFacturacion.ChequeaSiEsPagosAcuenta(rsreporte!IdComprobantePago, _
                                                                oConexion, rsreporte!Adelantos, lnRedondeo, _
                                                                rsreporte!IdTipoOrden, rsreporte!exoneraciones, _
                                                                rsreporte!Adelantos, rsreporte!idEstadoComprobante, _
                                                                rsreporte!TotalPagado)
                    
                     If lnRedondeo <> 9999 Then
                        .AddNew
                        .Fields!nroSerie = rsreporte.Fields!nroSerie
                        .Fields!nrodocumento = rsreporte.Fields!nrodocumento
                        .Fields!NroHistoriaClinica = mo_ReporteUtil.NullToVacio(rsreporte.Fields!NroHistoriaClinica)
                        .Fields!razonSocial = mo_ReporteUtil.NullToVacio(rsreporte.Fields!razonSocial)
                        If rsreporte.Fields!Adelantos > 0 Then
                           .Fields!Adelantos = rsreporte.Fields!Adelantos
                        ElseIf rsreporte!idProducto = lnIdPagosACuenta Then
                           .Fields!Adelantos = rsreporte.Fields!Subtotal
                        ElseIf rsreporte!idProducto Then
                           .Fields!Devoluciones = rsreporte.Fields!Subtotal
                        Else
                           .Fields!Adelantos = 0
                        End If
                        If rsreporte.Fields!idEstadoComprobante = 9 Then
                           .Fields!importeEXO = 0
                        Else
                           .Fields!importeEXO = IIf(IsNull(rsreporte.Fields!exoneraciones), 0, rsreporte.Fields!exoneraciones)
                        End If
                        .Fields!idEstadoComprobante = rsreporte.Fields!idEstadoComprobante
                        .Fields!TotalPorPagar = IIf(IsNull(rsreporte.Fields!Subtotal), 0, rsreporte.Fields!Subtotal)
                        .Fields!idProducto = rsreporte.Fields!idProducto
                        .Fields!TotalPagado = rsreporte.Fields!TotalPagado
                        .Fields!cajero = mo_ReglasCaja.SeleccionaDatosCajero(rsreporte.Fields!IdCajero, sghIniciales)
                        .Fields!FechaCobranza = rsreporte.Fields!FechaCobranza
                        .Fields!redondeo = lnRedondeo
                        .Update
                    End If
                    rsreporte.MoveNext
                 Loop
            End With
        End If
        rsreporte.Close
        'MEDICAMENTOS emitidos en CAJA SERVICIO
        Set rsreporte = mo_ReglasFacturacion.FarmaciaConsolidado(ml_IdGestionCaja, CDate(mda_FechaInicio), CDate(mda_FechaFin), ml_IdTurno, ml_IdCajero)
        lcFilter = ""
        If ml_IdComprobantePago > 0 Then
            lcFilter = "IdTipoComprobante=" & ml_IdComprobantePago & " and "
        End If
        If ml_IdTipoFinanciamiento > 0 Then
           lcFilter = lcFilter & "idTipoFinanciamiento=" & ml_IdTipoFinanciamiento & " and "
        End If
        lcFilter = lcFilter & IIf(lbSoloCreditos = True, "TieneCredito<>null", "TieneCredito=null") & " and "
        
        If lcFilter <> "" Then
           lcFilter = Left(lcFilter, Len(lcFilter) - 4)
           rsreporte.Filter = lcFilter
        End If
        If rsreporte.RecordCount > 0 Then
            With rsRsTmp
                 rsreporte.MoveFirst
                 Do While Not rsreporte.EOF
                    lbTienePagoAcuenta = mo_ReglasFacturacion.ChequeaSiEsPagosAcuenta(rsreporte!IdComprobantePago, _
                                                                oConexion, rsreporte!Adelantos, lnRedondeo, _
                                                                rsreporte!IdTipoOrden, rsreporte!exoneraciones, _
                                                                rsreporte!Adelantos, rsreporte!idEstadoComprobante, _
                                                                rsreporte!TotalPagado)
                     If lnRedondeo <> 9999 Then
                        .AddNew
                        .Fields!nroSerie = rsreporte.Fields!nroSerie
                        .Fields!nrodocumento = rsreporte.Fields!nrodocumento
                        .Fields!NroHistoriaClinica = mo_ReporteUtil.NullToVacio(rsreporte.Fields!NroHistoriaClinica)
                        .Fields!razonSocial = mo_ReporteUtil.NullToVacio(rsreporte.Fields!razonSocial)
                        .Fields!Adelantos = IIf(IsNull(rsreporte.Fields!Adelantos), 0, rsreporte.Fields!Adelantos)
                        .Fields!importeEXO = IIf(IsNull(rsreporte.Fields!exoneraciones), 0, rsreporte.Fields!exoneraciones)
                        .Fields!idEstadoComprobante = rsreporte.Fields!idEstadoComprobante
                        .Fields!TotalPorPagar = IIf(IsNull(rsreporte.Fields!Subtotal), 0, rsreporte.Fields!Subtotal)
                        .Fields!idProducto = rsreporte.Fields!idProducto
                        .Fields!TotalPagado = rsreporte.Fields!TotalPagado
                        '.Fields!redondeo = mo_ReglasFacturacion.DevuelveRedondeoEnCadaBoletaFarmacia(rsReporte!totalPagado, rsReporte!exoneraciones + rsReporte!Dctos, rsReporte!IdComprobantePago, rsReporte!idEstadoComprobante, lcCadenaConexion)
                        .Fields!redondeo = lnRedondeo
                        .Fields!cajero = mo_ReglasCaja.SeleccionaDatosCajero(rsreporte.Fields!IdCajero, sghIniciales)
                        .Fields!FechaCobranza = rsreporte.Fields!FechaCobranza
                        .Update
                    End If
                    rsreporte.MoveNext
                 Loop
            End With
        End If
        rsreporte.Close
        
        Set rsreporte = rsRsTmp.Clone
        
        'NOTAS DE CREDITO pagados en CAJA SERVICIO FRANK
        lbVerResumenNotaCredito = True
        Set rsNotasCredito = mo_AdminCaja.NotaCreditoDevueltosPorNumYFecha("", "", CDate(mda_FechaInicio), CDate(mda_FechaFin))
        lcFilter = ""
        If ml_IdGestionCaja > 0 Then
            lcFilter = "IdCaja=" & ml_IdGestionCaja & " and "
        End If
        If ml_IdTurno > 0 Then
            lcFilter = "idTurno=" & ml_IdTurno & " and "
        End If
        If ml_IdCajero > 0 Then
            lcFilter = "IdCajero=" & ml_IdCajero & " and "
        End If
        If ml_IdComprobantePago > 0 Then
            lbVerResumenNotaCredito = False
        End If
        If ml_IdTipoFinanciamiento > 0 Then
           lbVerResumenNotaCredito = False
        End If
        If lcFilter <> "" Then
           lcFilter = Left(lcFilter, Len(lcFilter) - 4)
           rsNotasCredito.Filter = lcFilter
        End If
        If rsNotasCredito.RecordCount = 0 Then
            lbVerResumenNotaCredito = False
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        iFila = 6
        If rsreporte.RecordCount = 0 Then
            If lbVerResumenNotaCredito = False Then
                MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
            End If
        Else
            rsreporte.Sort = "nroSerie,nroDocumento"
            lRecordCount = 0: lntImpSubTot = 0: lntImpAnul = 0: lntImpExo = 0: lntImpDevol = 0
            lntImpPagCta = 0: lntImpTot = 0: lntImpRedondeo = 0
            
            Set range_row = oWorkSheet.range("B" & CStr(iFila), "M" & CStr(iFila))
            range_row.Interior.Color = RGB(173, 234, 234)
            Set borders = range_row.borders
            borders.LineStyle = Excel.XlLineStyle.xlContinuous
            
            iFila = iFila + 1
            rsreporte.MoveFirst
            Do While Not rsreporte.EOF
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(COL_FECHA - 1, iFila - 1).setFormula("'" & Format(rsreporte.Fields!FechaCobranza, sighEntidades.DevuelveFechaSoloFormato_DMY))
                    Call Feuille.getcellbyposition(COL_USUARIO - 1, iFila - 1).setFormula(rsreporte!cajero)
                    Call Feuille.getcellbyposition(COL_BOLETA - 1, iFila - 1).setFormula(rsreporte!nroSerie + " - " + rsreporte!nrodocumento)
                    Call Feuille.getcellbyposition(COL_NRO_HISTORIA - 1, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsreporte!NroHistoriaClinica))
                    Call Feuille.getcellbyposition(COL_RAZON_SOCIAL - 1, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsreporte!razonSocial))
                Else
                    oWorkSheet.Cells(iFila, COL_FECHA).Value = "'" & Format(rsreporte.Fields!FechaCobranza, sighEntidades.DevuelveFechaSoloFormato_DMY)
                    oWorkSheet.Cells(iFila, COL_USUARIO).Value = rsreporte!cajero
                    oWorkSheet.Cells(iFila, COL_BOLETA).Value = rsreporte!nroSerie + " - " + rsreporte!nrodocumento
                    oWorkSheet.Cells(iFila, COL_NRO_HISTORIA).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(rsreporte!NroHistoriaClinica, True)
                    oWorkSheet.Cells(iFila, COL_RAZON_SOCIAL).Value = mo_ReporteUtil.NullToVacio(rsreporte!razonSocial)
                End If
                
                lnImpSubTot = rsreporte!TotalPorPagar
                lnImpExo = rsreporte!importeEXO
                lnImpAnul = IIf(rsreporte!idEstadoComprobante = 9, rsreporte!TotalPagado, 0)
                lnImpPagCta = rsreporte!Adelantos
                lnImpTot = rsreporte!TotalPagado
                lnImpRedondeo = rsreporte!redondeo
                lnImpDevol = IIf(rsreporte!idEstadoComprobante = 6, rsreporte!TotalPagado, 0)
                
                lRecordCount = lRecordCount + 1
                lcLlave = (rsreporte!nroSerie + rsreporte!nrodocumento)
                'lnImpAnul = 0: lnImpDevol = 0
                'lnImpExo = 0:  lnImpPagCta = 0: lnImpTot = 0
                'lnDctos = IIf(IsNull(rsReporte!adelantos), 0, rsReporte!adelantos) 'Descuentos en el COMPROBANTE PAGO
                Do While Not rsreporte.EOF And lcLlave = (rsreporte!nroSerie + rsreporte!nrodocumento)
                    
                    Select Case rsreporte!idEstadoComprobante
                    Case 6                                               'Devolucion
                    Case 9                                               'Anulado
                    Case 4
                       If rsreporte!idProducto = lnIdPagosACuenta Then                 'Pago a cuenta
                       Else
                       End If
                    End Select
                    
                   rsreporte.MoveNext
                   If rsreporte.EOF Then
                      Exit Do
                   End If
                Loop
                If lnImpAnul > 0 Then
                   lnImpTot = 0
                ElseIf lnImpDevol > 0 Then
                   lnImpTot = -lnImpDevol
                End If
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(COL_SUBTOTAL - 1, iFila - 1).setFormula(lnImpSubTot)
                    Call Feuille.getcellbyposition(COL_REDONDEO - 1, iFila - 1).setFormula(lnImpRedondeo)
                    Call Feuille.getcellbyposition(COL_EXONERADO - 1, iFila - 1).setFormula(lnImpExo)
                    Call Feuille.getcellbyposition(COL_ANULADO - 1, iFila - 1).setFormula(lnImpAnul)
                    Call Feuille.getcellbyposition(COL_DEVOLUCION - 1, iFila - 1).setFormula(lnImpDevol)
                    Call Feuille.getcellbyposition(COL_PAGOCTA - 1, iFila - 1).setFormula(lnImpPagCta)
                    Call Feuille.getcellbyposition(COL_TOTAL - 1, iFila - 1).setFormula(lnImpTot)
                Else
                    oWorkSheet.Cells(iFila, COL_SUBTOTAL).Value = lnImpSubTot
                    oWorkSheet.Cells(iFila, COL_REDONDEO).Value = lnImpRedondeo
                    oWorkSheet.Cells(iFila, COL_EXONERADO).Value = lnImpExo
                    oWorkSheet.Cells(iFila, COL_ANULADO).Value = lnImpAnul
                    oWorkSheet.Cells(iFila, COL_DEVOLUCION).Value = lnImpDevol
                    oWorkSheet.Cells(iFila, COL_PAGOCTA).Value = lnImpPagCta
                    oWorkSheet.Cells(iFila, COL_TOTAL).Value = lnImpTot
                End If
                
                lntImpSubTot = lntImpSubTot + lnImpSubTot
                lntImpExo = lntImpExo + lnImpExo
                lntImpAnul = lntImpAnul + lnImpAnul
                lntImpDevol = lntImpDevol + lnImpDevol
                lntImpTot = lntImpTot + lnImpTot
                If lnImpAnul = 0 Then lntImpPagCta = lntImpPagCta + lnImpPagCta
                iFila = iFila + 1
                lntImpRedondeo = lntImpRedondeo + lnImpRedondeo
            Loop
            iFila = iFila + 1
            
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(2) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(COL_TOTAL) & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Cantidad de Documentos: " + Trim(Str(lRecordCount)))
                Call Feuille.getcellbyposition(COL_SUBTOTAL - 1, iFila - 1).setFormula(lntImpSubTot)
                Call Feuille.getcellbyposition(COL_REDONDEO - 1, iFila - 1).setFormula(lntImpRedondeo)
                Call Feuille.getcellbyposition(COL_EXONERADO - 1, iFila - 1).setFormula(lntImpExo)
                Call Feuille.getcellbyposition(COL_ANULADO - 1, iFila - 1).setFormula(lntImpAnul)
                Call Feuille.getcellbyposition(COL_DEVOLUCION - 1, iFila - 1).setFormula(lntImpDevol)
                Call Feuille.getcellbyposition(COL_PAGOCTA - 1, iFila - 1).setFormula(lntImpPagCta)
                Call Feuille.getcellbyposition(COL_TOTAL - 1, iFila - 1).setFormula(lntImpTot)
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, COL_TOTAL
                oWorkSheet.Cells(iFila, 2).Value = "Cantidad de Documentos: " + Trim(Str(lRecordCount))
                oWorkSheet.Cells(iFila, COL_SUBTOTAL).Value = lntImpSubTot
                oWorkSheet.Cells(iFila, COL_REDONDEO).Value = lntImpRedondeo
                oWorkSheet.Cells(iFila, COL_EXONERADO).Value = lntImpExo
                oWorkSheet.Cells(iFila, COL_ANULADO).Value = lntImpAnul
                oWorkSheet.Cells(iFila, COL_DEVOLUCION).Value = lntImpDevol
                oWorkSheet.Cells(iFila, COL_PAGOCTA).Value = lntImpPagCta
                oWorkSheet.Cells(iFila, COL_TOTAL).Value = lntImpTot
                '
            End If
        End If
        
        If rsreporte.RecordCount > 0 Or (lbVerResumenNotaCredito = True And rsNotasCredito.RecordCount > 0) Then
            If rsNotasCredito.RecordCount > 0 Then
                iFila = iFila + 2
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(COL_FECHA - 1, iFila - 1).setFormula("Notas de crédito cobradas")
                    Call Feuille.getcellbyposition(COL_ANULADO - 1, iFila - 1).setFormula("Resumen")
                    Call Feuille.getcellbyposition(COL_DEVOLUCION - 1, iFila - 1).setFormula("")
                Else
                    oWorkSheet.Cells(iFila, COL_FECHA).Value = "Notas de crédito cobradas"
                    oWorkSheet.Cells(iFila, COL_ANULADO).Value = "Resumen"
                    oWorkSheet.Cells(iFila, COL_DEVOLUCION).Value = ""
                End If
                iFila = iFila + 1
                iFilaEmpiezaNota = iFila
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(COL_FECHA - 1, iFila - 1).setFormula("Fecha_Emisión")
                    Call Feuille.getcellbyposition(COL_USUARIO - 1, iFila - 1).setFormula("Cajero")
                    Call Feuille.getcellbyposition(COL_BOLETA - 1, iFila - 1).setFormula("N° NotaCredito")
                    Call Feuille.getcellbyposition(COL_NRO_HISTORIA - 1, iFila - 1).setFormula("N° Historia")
                    Call Feuille.getcellbyposition(COL_RAZON_SOCIAL - 1, iFila - 1).setFormula("Nombres y Apellidos / Razon Social")
                    Call Feuille.getcellbyposition(COL_SUBTOTAL - 1, iFila - 1).setFormula("Total")
                Else
                    oWorkSheet.Cells(iFila, COL_FECHA).Value = "Fecha_Emisión"
                    oWorkSheet.Cells(iFila, COL_USUARIO).Value = "Cajero"
                    oWorkSheet.Cells(iFila, COL_BOLETA).Value = "N° NotaCredito"
                    oWorkSheet.Cells(iFila, COL_NRO_HISTORIA).Value = "N° Historia"
                    oWorkSheet.Cells(iFila, COL_RAZON_SOCIAL).Value = "Nombres y Apellidos / Razon Social"
                    oWorkSheet.Cells(iFila, COL_SUBTOTAL).Value = "Total"
                    
                    Set range_row = oWorkSheet.range("B" & CStr(iFila), "G" & CStr(iFila))
                    range_row.Interior.Color = RGB(173, 234, 234)
                    Set borders = range_row.borders
                    borders.LineStyle = Excel.XlLineStyle.xlContinuous
                End If
                iFila = iFila + 1
                
                rsNotasCredito.MoveFirst
                Do While Not rsNotasCredito.EOF
                    lnNotaCredTot = lnNotaCredTot + rsNotasCredito!Total
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(COL_FECHA - 1, iFila - 1).setFormula("'" & Format(rsNotasCredito.Fields!fecha, sighEntidades.DevuelveFechaSoloFormato_DMY))
                        Call Feuille.getcellbyposition(COL_USUARIO - 1, iFila - 1).setFormula(mo_ReglasCaja.SeleccionaDatosCajero(rsNotasCredito.Fields!IdCajero, sghIniciales))
                        Call Feuille.getcellbyposition(COL_BOLETA - 1, iFila - 1).setFormula(rsNotasCredito!nroSerie + " - " + rsNotasCredito!nrodocumento)
                        Call Feuille.getcellbyposition(COL_NRO_HISTORIA - 1, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsNotasCredito!NroHistoriaClinica))
                        Call Feuille.getcellbyposition(COL_RAZON_SOCIAL - 1, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsNotasCredito!razonSocial))
                        Call Feuille.getcellbyposition(COL_SUBTOTAL - 1, iFila - 1).setFormula(rsNotasCredito!Total)
                    Else
                        oWorkSheet.Cells(iFila, COL_FECHA).Value = "'" & Format(rsNotasCredito.Fields!fecha, sighEntidades.DevuelveFechaSoloFormato_DMY)
                        oWorkSheet.Cells(iFila, COL_USUARIO).Value = mo_ReglasCaja.SeleccionaDatosCajero(rsNotasCredito.Fields!IdCajero, sghIniciales)
                        oWorkSheet.Cells(iFila, COL_BOLETA).Value = rsNotasCredito!nroSerie + " - " + rsNotasCredito!nrodocumento
                        oWorkSheet.Cells(iFila, COL_NRO_HISTORIA).Value = mo_ReporteUtil.NullToVacio(rsNotasCredito!NroHistoriaClinica)
                        oWorkSheet.Cells(iFila, COL_RAZON_SOCIAL).Value = mo_ReporteUtil.NullToVacio(rsNotasCredito!razonSocial)
                        oWorkSheet.Cells(iFila, COL_SUBTOTAL).Value = rsNotasCredito!Total
                    End If
                    iFila = iFila + 1
                    rsNotasCredito.MoveNext
                Loop
                
                If lbEsOpenOffice = True Then
                    Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(2) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(COL_SUBTOTAL) & CStr(iFila))
                    mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Cantidad de Documentos: " + Trim(Str(rsNotasCredito.RecordCount)))
                    Call Feuille.getcellbyposition(COL_SUBTOTAL - 1, iFila - 1).setFormula(lnNotaCredTot)
                Else
                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, COL_SUBTOTAL
                    oWorkSheet.Cells(iFila, 2).Value = "Cantidad de Documentos: " + Trim(Str(rsNotasCredito.RecordCount))
                    oWorkSheet.Cells(iFila, COL_SUBTOTAL).Value = lnNotaCredTot
                End If
                
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(COL_ANULADO - 1, iFilaEmpiezaNota - 1).setFormula("Documentos")
                    Call Feuille.getcellbyposition(COL_EXONERADO - 1, iFilaEmpiezaNota - 1).setFormula("Cantidad")
                    Call Feuille.getcellbyposition(COL_DEVOLUCION - 1, iFilaEmpiezaNota - 1).setFormula("Total")
                    iFilaEmpiezaNota = iFilaEmpiezaNota + 1
                    Call Feuille.getcellbyposition(COL_ANULADO - 1, iFilaEmpiezaNota - 1).setFormula("Boletas")
                    Call Feuille.getcellbyposition(COL_EXONERADO - 1, iFilaEmpiezaNota - 1).setFormula("'" + Trim(Str(lRecordCount)))
                    Call Feuille.getcellbyposition(COL_DEVOLUCION - 1, iFilaEmpiezaNota - 1).setFormula(Trim(Str(lntImpTot)))
                    iFilaEmpiezaNota = iFilaEmpiezaNota + 1
                    Call Feuille.getcellbyposition(COL_ANULADO - 1, iFilaEmpiezaNota - 1).setFormula("Notas Crédito")
                    Call Feuille.getcellbyposition(COL_EXONERADO - 1, iFilaEmpiezaNota - 1).setFormula("'" + Trim(Str(rsNotasCredito.RecordCount)))
                    Call Feuille.getcellbyposition(COL_DEVOLUCION - 1, iFilaEmpiezaNota - 1).setFormula(Trim(Str(lnNotaCredTot)))
                    iFilaEmpiezaNota = iFilaEmpiezaNota + 1
                    Call Feuille.getcellbyposition(COL_EXONERADO - 1, iFilaEmpiezaNota - 1).setFormula("TOTAL")
                    Call Feuille.getcellbyposition(COL_DEVOLUCION - 1, iFilaEmpiezaNota - 1).setFormula(Trim(Str(lntImpTot - lnNotaCredTot)))
                    Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(COL_ANULADO) & CStr(iFilaEmpiezaNota) & ":" & mo_ReglasReportes.BuscaNombreColumna(COL_DEVOLUCION) & CStr(iFilaEmpiezaNota))
                    mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                Else
                    oWorkSheet.Cells(iFilaEmpiezaNota, COL_ANULADO).Value = "Documentos"
                    oWorkSheet.Cells(iFilaEmpiezaNota, COL_EXONERADO).Value = "Cantidad"
                    oWorkSheet.Cells(iFilaEmpiezaNota, COL_DEVOLUCION).Value = "Total"
                    
                    Set range_row = oWorkSheet.range("I" & CStr(iFilaEmpiezaNota), "K" & CStr(iFilaEmpiezaNota))
                    range_row.Interior.Color = RGB(173, 234, 234)
                    Set borders = range_row.borders
                    borders.LineStyle = Excel.XlLineStyle.xlContinuous
                    
                    iFilaEmpiezaNota = iFilaEmpiezaNota + 1
                    oWorkSheet.Cells(iFilaEmpiezaNota, COL_ANULADO).Value = "Boletas"
                    oWorkSheet.Cells(iFilaEmpiezaNota, COL_EXONERADO).Value = "'" + Trim(Str(lRecordCount))
                    oWorkSheet.Cells(iFilaEmpiezaNota, COL_DEVOLUCION).Value = Trim(Str(lntImpTot))
                    iFilaEmpiezaNota = iFilaEmpiezaNota + 1
                    Set range_row = oWorkSheet.range("I" & CStr(iFilaEmpiezaNota), "K" & CStr(iFilaEmpiezaNota))
                    range_row.Font.Color = RGB(255, 0, 0)
                    oWorkSheet.Cells(iFilaEmpiezaNota, COL_ANULADO).Value = "Notas Crédito"
                    oWorkSheet.Cells(iFilaEmpiezaNota, COL_EXONERADO).Value = "'" + Trim(Str(rsNotasCredito.RecordCount))
                    oWorkSheet.Cells(iFilaEmpiezaNota, COL_DEVOLUCION).Value = Trim(Str(lnNotaCredTot))
                    iFilaEmpiezaNota = iFilaEmpiezaNota + 1
                    oWorkSheet.Cells(iFilaEmpiezaNota, COL_EXONERADO).Value = "TOTAL"
                    oWorkSheet.Cells(iFilaEmpiezaNota, COL_DEVOLUCION).Value = Trim(Str(lntImpTot - lnNotaCredTot))
                    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFilaEmpiezaNota, COL_ANULADO, iFilaEmpiezaNota, COL_DEVOLUCION
                End If
                
                
            End If
            
            
            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 0
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 16
                PrintArea(0).EndRow = iFila
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.GetFrame.getContainerWindow.SetVisible(True)
                MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
            Else
                oWorkSheet.PageSetup.PrintTitleRows = "$2:$6"
                If oWorkSheet.PageSetup.PrintArea <> "" Then
                    oWorkSheet.PageSetup.PrintArea = sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                End If
                oExcel.Visible = True
                oWorkSheet.PrintPreview
                'oWorkSheet.PrintOut
                'oWorkBook.Close SaveChanges:=False
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
            'liberar memoria
            Set oExcel = Nothing
            Set oWorkBookPlantilla = Nothing
            Set oWorkBook = Nothing
            Set oWorkSheet = Nothing
        End If
    Set oConexion = Nothing
    Set rsreporte = Nothing
    Set lcBuscaParametro = Nothing
    Set mo_ReglasReportes = Nothing
Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    'Resume
    Exit Sub


End Sub

'***************daniel barrantes**************
'***************reporte: CONSOLIDADO DE FARMACIA
'***************
Sub CrearReporteConsolidadoFarmacia()
Dim lcLlave As String
Dim lnImpSubTot As Double: Dim lntImpSubTot As Double
Dim lnImpAnul As Double: Dim lntImpAnul As Double
Dim lnImpExo As Double: Dim lntImpExo As Double
Dim lnImpDevol As Double: Dim lntImpDevol As Double
Dim lnImpPagCta As Double: Dim lntImpPagCta As Double
Dim lnImpTot As Double: Dim lntImpTot As Double
Dim impigv As Double: Dim IGV As Double
Dim impbruto As Double: Dim inbruto As Double
Dim lnDctos As Double, lnImpRedondeo As Double, lntImpRedondeo As Double
Dim iFila As Long
Dim lRecordCount As Long, lcCadenaConexion As String
Dim rsreporte As New Recordset
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasReportes As New ReglasReportes
Dim CantidadSOAT As Long: Dim PrecioSOAT As Double
Dim lbEsOpenOffice As Boolean
Dim lcSql As String
Dim oConexion As New Connection, lbTienePagoAcuenta As Boolean, lnRedondeo As Double

oConexion.CommandTimeout = 300
oConexion.CursorLocation = adUseClient
oConexion.Open sighEntidades.CadenaConexion

lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)

On Error GoTo ManejadorError
    
    Const COL_FECHA = 2
    Const COL_USUARIO = 3
    Const COL_BOLETA = 4
    Const COL_NRO_HISTORIA = 5
    Const COL_RAZON_SOCIAL = 6
    Const COL_SUBTOTAL = 7
    Const COL_REDONDEO = 8
    Const COL_ANULADO = 9
    Const COL_EXONERADO = 10
    Const COL_DEVOLUCION = 11
    Const COL_PAGOCTA = 12
    Const COL_TOTAL_BRUTO = 13
    Const COL_IGV = 14
    Const COL_TOTAL_NETO = 15
    
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
        Dim lnHwnd As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
        Dim oRange As range
        Dim range As Excel.range
        Dim borders As Excel.borders
    End If
    
         lcCadenaConexion = sighEntidades.CadenaConexion
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\ECajaConsolidadoFarm.ods"
    '        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
    '        Chemin = "file:///" & App.Path & "\Plantillas\"
    '        Chemin = Replace(Chemin, "\", "/")
    '        Fichier = Chemin & "/OpenOffice.ods"
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
            Call Feuille.getcellbyposition(4, 1).setFormula("CONSOLIDADO DE FARMACIA")
           ' Call Feuille.getcellbyposition(13, 1).setFormula(lcBuscaParametro.RetornaFechaHoraServidorSQL)
            Call Feuille.getcellbyposition(2, 3).setFormula(ml_TextoDelFiltro)
        Else
            'Crea nueva hoja
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\ECajaConsolidadoFarm.xls")
            oWorkBookPlantilla.Worksheets("CajaConsolidadoFarm").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            oWorkSheet.Cells(2, 5).Value = "        CONSOLIDADO DE FARMACIA"
            'oWorkSheet.Cells(2, 14).Value = lcBuscaParametro.RetornaFechaHoraServidorSQL
            oWorkSheet.Cells(4, 3).Value = ml_TextoDelFiltro
        End If
        'Set rsReporte = ReporteConsolidadoFarmaciaDEBB(ml_IdGestionCaja, mda_FechaInicio, mda_FechaFin, ml_IdTurno, ml_IdCajero)
        Set rsreporte = mo_ReglasFacturacion.FarmaciaConsolidado(ml_IdGestionCaja, CDate(mda_FechaInicio), CDate(mda_FechaFin), ml_IdTurno, ml_IdCajero)
        If rsreporte.RecordCount = 0 Then
           MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
        Else
        iFila = 7
        lRecordCount = 0: lntImpSubTot = 0: lntImpAnul = 0: lntImpExo = 0: lntImpDevol = 0
        lntImpPagCta = 0: lntImpTot = 0: IGV = 0: impigv = 0: impbruto = 0: inbruto = 0: lntImpRedondeo = 0
        rsreporte.MoveFirst
        Do While Not rsreporte.EOF
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(COL_FECHA - 1, iFila - 1).setFormula("'" & Format(rsreporte.Fields!FechaCobranza, sighEntidades.DevuelveFechaSoloFormato_DMY))
                Call Feuille.getcellbyposition(COL_USUARIO - 1, iFila - 1).setFormula(mo_ReglasCaja.SeleccionaDatosCajero(rsreporte.Fields!IdCajero, sghIniciales))
                Call Feuille.getcellbyposition(COL_BOLETA - 1, iFila - 1).setFormula(rsreporte!nroSerie + " - " + rsreporte!nrodocumento)
                Call Feuille.getcellbyposition(COL_NRO_HISTORIA - 1, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsreporte!NroHistoriaClinica))
                Call Feuille.getcellbyposition(COL_RAZON_SOCIAL - 1, iFila - 1).setFormula(mo_ReporteUtil.NullToVacio(rsreporte!razonSocial))
            Else
                oWorkSheet.Cells(iFila, COL_FECHA).Value = "'" & Format(rsreporte.Fields!FechaCobranza, sighEntidades.DevuelveFechaSoloFormato_DMY)
                oWorkSheet.Cells(iFila, COL_USUARIO).Value = mo_ReglasCaja.SeleccionaDatosCajero(rsreporte.Fields!IdCajero, sghIniciales)
                oWorkSheet.Cells(iFila, COL_BOLETA).Value = rsreporte!nroSerie + " - " + rsreporte!nrodocumento
                oWorkSheet.Cells(iFila, COL_NRO_HISTORIA).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(rsreporte!NroHistoriaClinica, True)
                oWorkSheet.Cells(iFila, COL_RAZON_SOCIAL).Value = mo_ReporteUtil.NullToVacio(rsreporte!razonSocial)
            End If
            lnImpSubTot = rsreporte!Subtotal
            If rsreporte!idEstadoComprobante = 9 Then
               lnImpExo = 0
            Else
               lnImpExo = rsreporte!exoneraciones
            End If
            lnImpAnul = IIf(rsreporte!idEstadoComprobante = 9, rsreporte!TotalPagado, 0)
            lnImpPagCta = rsreporte!Adelantos
            lnImpTot = rsreporte!TotalPagado
            'If Trim(rsReporte!nroDocumento) = "958224" Then
            'lnImpRedondeo = 0
            'End If
            'lnImpRedondeo = mo_ReglasFacturacion.DevuelveRedondeoEnCadaBoletaFarmacia(rsReporte!totalPagado, rsReporte!exoneraciones + rsReporte!Dctos, rsReporte!IdComprobantePago, rsReporte!idEstadoComprobante, lcCadenaConexion)
            IGV = IIf(rsreporte!idEstadoComprobante = 9, 0, rsreporte!IGV)
            inbruto = IIf(rsreporte!idEstadoComprobante = 9, 0, rsreporte!Subtotal)
            lbTienePagoAcuenta = mo_ReglasFacturacion.ChequeaSiEsPagosAcuenta(rsreporte!IdComprobantePago, _
                                                        oConexion, rsreporte!Adelantos, lnImpRedondeo, _
                                                        rsreporte!IdTipoOrden, rsreporte!exoneraciones, _
                                                        rsreporte!Adelantos, rsreporte!idEstadoComprobante, _
                                                        rsreporte!TotalPagado)
            
            lRecordCount = lRecordCount + 1
            lcLlave = (rsreporte!nroSerie + rsreporte!nrodocumento)
            Do While Not rsreporte.EOF And lcLlave = (rsreporte!nroSerie + rsreporte!nrodocumento)
                    Select Case rsreporte!idEstadoComprobante
                    Case 6                                               'Devolucion
                    Case 9                                               'Anulado
                    Case 4
                       If rsreporte!idProducto = lnIdPagosACuenta Then                 'Pago a cuenta
                       Else
                       End If
                    End Select
               rsreporte.MoveNext
               If rsreporte.EOF Then
                  Exit Do
               End If
            Loop
            'inbruto = lnImpTot / 1.19
            'IGV = lnImpTot - inbruto
            If lnImpAnul > 0 Then
             '  IGV = 0
             '  inbruto = 0
               lnImpTot = 0
            ElseIf lnImpDevol > 0 Then
              ' IGV = 0
             '  inbruto = 0
               lnImpTot = 0
            End If
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(COL_SUBTOTAL - 1, iFila - 1).setFormula(lnImpSubTot)
                Call Feuille.getcellbyposition(COL_REDONDEO - 1, iFila - 1).setFormula(lnImpRedondeo)
                Call Feuille.getcellbyposition(COL_EXONERADO - 1, iFila - 1).setFormula(lnImpExo)
                Call Feuille.getcellbyposition(COL_ANULADO - 1, iFila - 1).setFormula(lnImpAnul)
                Call Feuille.getcellbyposition(COL_DEVOLUCION - 1, iFila - 1).setFormula(lnImpDevol)
                Call Feuille.getcellbyposition(COL_PAGOCTA - 1, iFila - 1).setFormula(lnImpPagCta)
                Call Feuille.getcellbyposition(COL_TOTAL_BRUTO - 1, iFila - 1).setFormula(inbruto)
                Call Feuille.getcellbyposition(COL_IGV - 1, iFila - 1).setFormula(IGV)
                Call Feuille.getcellbyposition(COL_TOTAL_NETO - 1, iFila - 1).setFormula(lnImpTot)
            Else
                oWorkSheet.Cells(iFila, COL_SUBTOTAL).Value = lnImpSubTot
                oWorkSheet.Cells(iFila, COL_REDONDEO).Value = lnImpRedondeo
                oWorkSheet.Cells(iFila, COL_EXONERADO).Value = lnImpExo
                oWorkSheet.Cells(iFila, COL_ANULADO).Value = lnImpAnul
                oWorkSheet.Cells(iFila, COL_DEVOLUCION).Value = lnImpDevol
                oWorkSheet.Cells(iFila, COL_PAGOCTA).Value = lnImpPagCta
                oWorkSheet.Cells(iFila, COL_TOTAL_BRUTO).Value = inbruto
                oWorkSheet.Cells(iFila, COL_IGV).Value = IGV
                oWorkSheet.Cells(iFila, COL_TOTAL_NETO).Value = lnImpTot
            End If
            lntImpSubTot = lntImpSubTot + lnImpSubTot
            lntImpExo = lntImpExo + lnImpExo
            lntImpAnul = lntImpAnul + lnImpAnul
            lntImpDevol = lntImpDevol + lnImpDevol
            lntImpTot = lntImpTot + lnImpTot
            lntImpPagCta = lntImpPagCta + lnImpPagCta
            impbruto = impbruto + inbruto ' lntImpTot / 1.19
            'impigv = lntImpTot - impbruto
            impigv = impigv + IGV
            lntImpRedondeo = lntImpRedondeo + lnImpRedondeo
            
            iFila = iFila + 1
        Loop
        iFila = iFila + 1
        
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(2) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(COL_TOTAL_NETO) & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Cantidad de Documentos: " + Trim(Str(lRecordCount)))
            Call Feuille.getcellbyposition(COL_SUBTOTAL - 1, iFila - 1).setFormula(lntImpSubTot)
            Call Feuille.getcellbyposition(COL_REDONDEO - 1, iFila - 1).setFormula(lntImpRedondeo)
            Call Feuille.getcellbyposition(COL_EXONERADO - 1, iFila - 1).setFormula(lntImpExo)
            Call Feuille.getcellbyposition(COL_ANULADO - 1, iFila - 1).setFormula(lntImpAnul)
            Call Feuille.getcellbyposition(COL_DEVOLUCION - 1, iFila - 1).setFormula(lntImpDevol)
            Call Feuille.getcellbyposition(COL_PAGOCTA - 1, iFila - 1).setFormula(lntImpPagCta)
            Call Feuille.getcellbyposition(COL_TOTAL_BRUTO - 1, iFila - 1).setFormula(impbruto)
            Call Feuille.getcellbyposition(COL_IGV - 1, iFila - 1).setFormula(impigv)
            Call Feuille.getcellbyposition(COL_TOTAL_NETO - 1, iFila - 1).setFormula(lntImpTot)
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, COL_TOTAL_NETO
            oWorkSheet.Cells(iFila, 2).Value = "Cantidad de Documentos: " + Trim(Str(lRecordCount))
            oWorkSheet.Cells(iFila, COL_SUBTOTAL).Value = lntImpSubTot
            oWorkSheet.Cells(iFila, COL_REDONDEO).Value = lntImpRedondeo
            oWorkSheet.Cells(iFila, COL_EXONERADO).Value = lntImpExo
            oWorkSheet.Cells(iFila, COL_ANULADO).Value = lntImpAnul
            oWorkSheet.Cells(iFila, COL_DEVOLUCION).Value = lntImpDevol
            oWorkSheet.Cells(iFila, COL_PAGOCTA).Value = lntImpPagCta
            oWorkSheet.Cells(iFila, COL_TOTAL_BRUTO).Value = impbruto
            oWorkSheet.Cells(iFila, COL_IGV).Value = impigv
            oWorkSheet.Cells(iFila, COL_TOTAL_NETO).Value = lntImpTot
        End If
        '
            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 0
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 16
                PrintArea(0).EndRow = iFila
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.GetFrame.getContainerWindow.SetVisible(True)
                MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
            Else
                oWorkSheet.PageSetup.PrintTitleRows = "$2:$6"
                If oWorkSheet.PageSetup.PrintArea <> "" Then
                   oWorkSheet.PageSetup.PrintArea = sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                End If
                oExcel.Visible = True
                oWorkSheet.PrintPreview
                'oWorkSheet.PrintOut
                'oWorkBook.Close SaveChanges:=False
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
        'liberar memoria
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
    Set oConexion = Nothing
    Set rsreporte = Nothing
    Set lcBuscaParametro = Nothing
    Set mo_ReglasReportes = Nothing
    Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub
End Sub



'***************daniel barrantes**************
'***************Busca datos para PARTE DIARIO
'***************
Function ReportDiarioDeCajaDEBB(lIdGestionCaja As Long, sEstados As String, lcFechaIni As String, lcFechaFin As String, lnIdTurno As Long, lnCajero As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String
Dim oConexion As New ADODB.Connection

        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        
        Set ReportDiarioDeCajaDEBB = Nothing
        
        sSQL = "select"
        sSQL = sSQL + " FactCatalogoServicios.IdProducto, FactPartidasPresupuestales.codigo , FactCatalogoServicios.Nombre, Sum(FacturacionServicios.TotalPorPagar) as TotalPorPagar"
        sSQL = sSQL + " from (((CajaGestion"
        sSQL = sSQL + " left join  CajaComprobantesPago on CajaGestion.IdGestionCaja = CajaComprobantesPago.IdGestionCaja)"
        sSQL = sSQL + " left join FacturacionServicios on CajaComprobantesPago.IdComprobantePago = FacturacionServicios.IdComprobantePago)"
        sSQL = sSQL + " left join FactCatalogoServicios on FacturacionServicios.IdProducto = FactCatalogoServicios.IdProducto)"
        sSQL = sSQL + " left join FactPartidasPresupuestales on FactCatalogoServicios.IdPartida = FactPartidasPresupuestales.IdPartidaPresupuestal"
        sSQL = sSQL + " Where FactCatalogoServicios.IdProducto Is Not Null"
        sSQL = sSQL + " and FacturacionServicios.IdEstadoFacturacion in (" + sEstados + ")"
        sSQL = sSQL + " and  CajaComprobantesPago.FechaCobranza Between (CONVERT(DATETIME,'" & lcFechaIni & " 00:00:00',103)) and (CONVERT(DATETIME,'" & lcFechaFin & " 23:59:59',103))"
        sSQL = sSQL + " and CajaGestion.idCaja=" & lIdGestionCaja
        sSQL = sSQL + " and CajaGestion.idTurno=" & lnIdTurno
        sSQL = sSQL + " and CajaGestion.idCajero=" & lnCajero
        sSQL = sSQL + " Group By"
        sSQL = sSQL + " FactCatalogoServicios.IdProducto , FactPartidasPresupuestales.codigo, FactCatalogoServicios.Nombre"
        
        With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = oConexion
        .CommandText = sSQL
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    
    Set ReportDiarioDeCajaDEBB = oRecordset
    oConexion.Close
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
    
End Function




Sub GenerarRecordsetTemporal()
    With rsRsTmp
          .Fields.Append "idPartida", adUnsignedBigInt
          .Fields.Append "Partida", adVarChar, 30, adFldIsNullable
          .Fields.Append "dPartida", adVarChar, 150, adFldIsNullable
          .Fields.Append "SubTotal", adDouble
          .Fields.Append "Exonerado", adDouble
          .Fields.Append "Anulado", adDouble
          .Fields.Append "Devolucion", adDouble
          .Fields.Append "PagoCta", adDouble
          .Fields.Append "Total", adDouble
          .LockType = adLockOptimistic
          .Open
    End With
End Sub


Sub GenerarRecordsetDctos()
    With rsRsTmp1
          .Fields.Append "Docume", adBSTR, 150
          .Fields.Append "Dctos", adDouble
          .LockType = adLockOptimistic
          .Open
    End With
End Sub

'***************daniel barrantes**************
'***************Imprime una BOLETA en EXCEL
'***************
Sub ImpresionBoletaEnExcel(nroSerie As String, nroDcto As String, lcFarmaciaServicio As Long)
    Dim mo_ReporteUtil As New ReporteUtil
    Dim oReglasCaja As New SIGHNegocios.ReglasCaja
    Dim rsreporte As New Recordset
    Dim sSQL As String: Dim lcAdelantos As String: Dim lcCuenta As String: Dim lcCajero As String
    Dim lnFila As Long:   Dim lnCant As Long: Dim lnPrimeraFila As Long
    Dim lnImporteEXO As Double: Dim lnSubTotal As Double: Dim lnTotal As Double
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim lbEsOpenOffice As Boolean
    Dim lcNombre As String, lcSql As String
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
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
        
        If lcFarmaciaServicio = sighEntidades.sghbien Then
           Set rsreporte = oReglasCaja.CajaComprobantePagoProductosPorNroSerieNroDocumento(nroSerie, nroDcto)
        Else
           Set rsreporte = oReglasCaja.CajaComprobantePagoServiciosPorNroSerieNroDocumento(nroSerie, nroDcto)
        End If
        
        If rsreporte.RecordCount > 0 Then
            'Crea nueva hoja
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\GalenHos.xls")
            oWorkBookPlantilla.Worksheets("Boleta").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            rsreporte.MoveFirst
            lnTotal = IIf(IsNull(rsreporte.Fields!TotalBoleta), "0", rsreporte.Fields!TotalBoleta)
            lcAdelantos = "Adelantos: " & Format(IIf(IsNull(rsreporte.Fields!Adelantos), 0, rsreporte.Fields!Adelantos), "######.#0")
            lcCuenta = "Cta: " & Trim(Str(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion))) & "    Ord.Pag: " & Trim(Str(rsreporte.Fields!IdOrden))
            lcCajero = oReglasCaja.SeleccionaDatosCajero(rsreporte.Fields!IdCajero, sghIniciales)
            lnFila = 1
            oWorkSheet.Cells(lnFila, 15).Value = "'" & IIf(rsreporte.Fields!IdTipoPago = 2, "DEVOLUCION", IIf(rsreporte.Fields!TotalBoleta = 0, "EXONERADO", ""))
            oWorkSheet.Cells(lnFila, 45).Value = "'" & nroSerie & " - " & nroDcto
            lnFila = lnFila + 2
            oWorkSheet.Cells(lnFila, 15).Value = "'" & IIf(rsreporte.Fields!idEstadoComprobante = 9, "A N U L A D O ", "")
            lnFila = lnFila + 2
            oWorkSheet.Cells(lnFila, 15).Value = "'" & Trim(IIf(IsNull(rsreporte.Fields!razonSocial), "", rsreporte.Fields!razonSocial)) + IIf(IsNull(rsreporte.Fields!NroHistoriaClinica), "(Particular)", "(HC: " & rsreporte.Fields!NroHistoriaClinica & ")")
            lnFila = lnFila + 1
            oWorkSheet.Cells(lnFila, 15).Value = "'" & Left(oReglasCaja.NombreServicioPorCuentaAtencion(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion)), 20)
            oWorkSheet.Cells(lnFila, 37).Value = "'" & Format(rsreporte.Fields!FechaCobranza, sighEntidades.DevuelveFechaSoloFormato_DMY)
            lnFila = lnFila + 2
            lnPrimeraFila = lnFila
            lnCant = 0
            Do While Not rsreporte.EOF
              If rsreporte.Fields!Cantidad > 0 And rsreporte.Fields!PrecioUnitario > 0 Then
                lnImporteEXO = lnImporteEXO + IIf(IsNull(rsreporte.Fields!exoneraciones), 0, rsreporte.Fields!exoneraciones)
                lnSubTotal = rsreporte.Fields!TotalPorPagar + IIf(IsNull(rsreporte.Fields!exoneraciones), 0, rsreporte.Fields!exoneraciones)
                oWorkSheet.Cells(lnFila, 10).Value = "'" & rsreporte.Fields!Codigo
                oWorkSheet.Cells(lnFila, 15).Value = "'" & Left(rsreporte.Fields!NombreProducto, 20)
                oWorkSheet.Cells(lnFila, 43).Value = "'" & rsreporte.Fields!Cantidad
                oWorkSheet.Cells(lnFila, 49).Value = "'" & Format(rsreporte.Fields!PrecioUnitario, "#####.#0")
                oWorkSheet.Cells(lnFila, 55).Value = "'" & Format(lnSubTotal, "######.#0")
                lnFila = lnFila + 1
                lnCant = lnCant + 1
              End If
              rsreporte.MoveNext
           Loop
           lnFila = lnFila + 4
           oWorkSheet.Cells(lnFila, 10).Value = "'" & lcCuenta
           lnFila = lnFila + 1
           oWorkSheet.Cells(lnFila, 10).Value = "'" & lcCajero
           oWorkSheet.Cells(lnFila, 15).Value = "'" & lcAdelantos
           oWorkSheet.Cells(lnFila, 35).Value = "'" & Format(lnTotal + lnImporteEXO, "######.#0")
           oWorkSheet.Cells(lnFila, 45).Value = "'" & Format(lnImporteEXO, "######.#0")
           oWorkSheet.Cells(lnFila, 55).Value = "'" & Format(lnTotal, "#######.#0")
           
           If oWorkSheet.PageSetup.PrintArea <> "" Then
              oWorkSheet.PageSetup.PrintArea = sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, lnFila)
           End If
           oExcel.Visible = True
           oWorkSheet.PrintPreview
           'oWorkSheet.PrintOut
           'oWorkBook.Close SaveChanges:=False
        End If
    Set oWorkSheet = Nothing
    Set oExcel = Nothing
Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Resume
End Sub

'***************daniel barrantes**************
'***************Imprime solo una BOLETA (FORMATO) y salta una PAGINA
'***************
Sub ImpresionBoletaEnDOS(nroSerie As String, nroDcto As String, lcFarmaciaServicio As Long)
Dim lnLineaSiguienteIMPR As Long: Dim lnNumLetrasIMPR As Long
Dim rsreporte As New ADODB.Recordset
Dim lnTotal As Double: Dim lnImporteEXO As Double: Dim lnSubTotal As Double
Dim lcAdelantos As String: Dim lcCuenta As String: Dim lcCajero As String
Dim lnFila As Long: Dim lnCant As Integer: Dim lnPrimeraFila As Long
Dim mo_ReporteUtil As New ReporteUtil
Dim oReglasCaja As New SIGHNegocios.ReglasCaja
Dim lcBuscaParametro As New SIGHDatos.Parametros

        On Error GoTo ManejadorError
        If lcFarmaciaServicio = sighEntidades.sghbien Then
           Set rsreporte = oReglasCaja.CajaComprobantePagoProductosPorNroSerieNroDocumento(nroSerie, nroDcto)
        Else
           Set rsreporte = oReglasCaja.CajaComprobantePagoServiciosPorNroSerieNroDocumento(nroSerie, nroDcto)
        End If
        If rsreporte.RecordCount > 0 Then
            '*****************Cabecera Inicio****************
            Printer.ScaleMode = 1
            'Printer.PaperSize = 256
            Printer.FontTransparent = True ' Esto es porque sino el texto te lo imprime siempre en fondo blanco.
            Printer.FontName = "Arial Narrow": Printer.FontSize = 8
'            Printer.FontName = "Sans Serif 12cpi"                     '<= tipo de letra proporcional
            Printer.Height = (305 * 56.7) + (Printer.TextHeight(" "))  '305mm<= Alto
            Printer.Width = (202 * 56.7)                               '202mm<= Ancho
            'Printer.PaperSize = 20
            Printer.Print ""                                           '<= Iniciar; impresora
            
            '....CurrentX <---Columna.............CurrentY <----Fila
            '....la "Fila" aumenta de 350 en 350 para pasar a otra Fila
            lnLineaSiguienteIMPR = 250
            '....la "Columna" si es 100 es una LETRA, si es 400 son 4 LETRAS
            lnNumLetrasIMPR = 100
            
            rsreporte.MoveFirst
            lnTotal = IIf(IsNull(rsreporte.Fields!TotalBoleta), "0", rsreporte.Fields!TotalBoleta)
            lcAdelantos = "Adelantos: " & Format(IIf(IsNull(rsreporte.Fields!Adelantos), 0, rsreporte.Fields!Adelantos), "######.#0")
            lcCuenta = "Cta: " & Trim(Str(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion))) & "    Ord.Pag: " & Trim(Str(rsreporte.Fields!IdOrden))
            lcCajero = oReglasCaja.SeleccionaDatosCajero(rsreporte.Fields!IdCajero, sghIniciales)
            
            Printer.CurrentX = 15 * lnNumLetrasIMPR: Printer.CurrentY = 0 * lnLineaSiguienteIMPR
            Printer.Print IIf(rsreporte.Fields!IdTipoPago = 2, "DEVOLUCION", IIf(rsreporte.Fields!TotalBoleta = 0, "EXONERADO", ".."))
            
            Printer.CurrentX = 15 * lnNumLetrasIMPR: Printer.CurrentY = 1 * lnLineaSiguienteIMPR
            Printer.Print IIf(rsreporte.Fields!idEstadoComprobante = 9, "A N U L A D O ", "..")
            
            Printer.CurrentX = 8 * lnNumLetrasIMPR: Printer.CurrentY = 3 * lnLineaSiguienteIMPR
            Printer.Print Trim(IIf(IsNull(rsreporte.Fields!razonSocial), "..", rsreporte.Fields!razonSocial)) + IIf(IsNull(rsreporte.Fields!NroHistoriaClinica), "(Particular)", "(HC: " & rsreporte.Fields!NroHistoriaClinica & ")")
            
            Printer.CurrentX = 49 * lnNumLetrasIMPR: Printer.CurrentY = 3 * lnLineaSiguienteIMPR
            Printer.Print nroSerie & " - " & nroDcto
            
            Printer.CurrentX = 8 * lnNumLetrasIMPR: Printer.CurrentY = 4 * lnLineaSiguienteIMPR
            Printer.Print Left(oReglasCaja.NombreServicioPorCuentaAtencion(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion)), 20)
            
            Printer.CurrentX = 40 * lnNumLetrasIMPR: Printer.CurrentY = 4 * lnLineaSiguienteIMPR
            Printer.Print Format(rsreporte.Fields!FechaCobranza, sighEntidades.DevuelveFechaSoloFormato_DMY)
            
            
            lnFila = (6 * lnLineaSiguienteIMPR)
            
            lnPrimeraFila = lnFila
            lnCant = 0
            Do While Not rsreporte.EOF
              If rsreporte.Fields!Cantidad > 0 And rsreporte.Fields!PrecioUnitario > 0 Then
                lnImporteEXO = lnImporteEXO + IIf(IsNull(rsreporte.Fields!importeEXO), 0, rsreporte.Fields!importeEXO)
                lnSubTotal = rsreporte.Fields!TotalPorPagar + IIf(IsNull(rsreporte.Fields!importeEXO), 0, rsreporte.Fields!importeEXO)
                
                Printer.CurrentX = 2 * lnNumLetrasIMPR: Printer.CurrentY = lnFila
                Printer.Print rsreporte.Fields!Codigo
                
                Printer.CurrentX = 8 * lnNumLetrasIMPR: Printer.CurrentY = lnFila
                Printer.Print Left(rsreporte.Fields!NombreProducto, 32)
                
                Printer.CurrentX = 40 * lnNumLetrasIMPR: Printer.CurrentY = lnFila
                Printer.Print rsreporte.Fields!Cantidad
                
                Printer.CurrentX = 47 * lnNumLetrasIMPR: Printer.CurrentY = lnFila
                Printer.Print Format(rsreporte.Fields!PrecioUnitario, "#####.#0")
                
                Printer.CurrentX = 53 * lnNumLetrasIMPR: Printer.CurrentY = lnFila
                Printer.Print Format(lnSubTotal, "######.#0")
                
                lnFila = lnFila + (1 * lnLineaSiguienteIMPR)
                lnCant = lnCant + 1
              End If
              rsreporte.MoveNext
           Loop
           lnFila = lnPrimeraFila + (10 * lnLineaSiguienteIMPR)
           Printer.CurrentX = 10 * lnNumLetrasIMPR: Printer.CurrentY = lnFila
           Printer.Print lcCuenta
           
           lnFila = lnFila + (1 * lnLineaSiguienteIMPR)
           
           Printer.CurrentX = 10 * lnNumLetrasIMPR: Printer.CurrentY = lnFila
           Printer.Print lcCajero
           
           Printer.CurrentX = 15 * lnNumLetrasIMPR: Printer.CurrentY = lnFila
           Printer.Print lcAdelantos
           
           Printer.CurrentX = 35 * lnNumLetrasIMPR: Printer.CurrentY = lnFila
           Printer.Print Format(lnTotal + lnImporteEXO, "######.#0")
           
           Printer.CurrentX = 45 * lnNumLetrasIMPR: Printer.CurrentY = lnFila
           Printer.Print Format(lnImporteEXO, "######.#0")
           
           Printer.CurrentX = 53 * lnNumLetrasIMPR: Printer.CurrentY = lnFila
           Printer.Print Format(lnTotal, "#######.#0")
           
           'Cierra Impresion
           Printer.EndDoc
        
        End If

        Exit Sub
ManejadorError:
        MsgBox Err.Description
        Resume
End Sub

Function Reembolso_ChequeaSiImprimeDetalleItems(rsreporte As Recordset) As Boolean
    If Not IsNull(rsreporte!Observaciones) Then
       If Left(rsreporte!Observaciones, 9) = "CUENTA_R:" Then
          If lcBuscaParametro.SeleccionaFilaParametro(574) = "S" Then
             Reembolso_ChequeaSiImprimeDetalleItems = True
          End If
       End If
    End If
    rsreporte.MoveFirst

End Function

Sub Reembolso_CargaItemsDeCuenta(rsRsBoletaDetallada As Recordset, ByRef rsreporte As Recordset, oConexion As Connection)
        Dim lnIdReembolsoServicios As Long, lnIdReembolsoFarmacia As Long, lnIdCuenta As Long, lnImporteItem As Double
        Dim oRsCabecera As New Recordset
        rsRsBoletaDetallada.MoveFirst
        lnIdReembolsoServicios = Val(lcBuscaParametro.SeleccionaFilaParametro(251))
        lnIdReembolsoFarmacia = Val(lcBuscaParametro.SeleccionaFilaParametro(252))
        lnIdCuenta = Val(Mid(rsRsBoletaDetallada!Observaciones, 10, 100))
        '
        oRsCabecera.Open "select * from reporte_cabecera where nroCuenta=" & Trim(Str(lnIdCuenta)) & " order by ServicioHosp", oConexion, adOpenKeyset, adLockOptimistic
        '
        'rsRsBoletaDetallada.MoveFirst
        Do While Not rsRsBoletaDetallada.EOF
            Select Case rsRsBoletaDetallada!idProducto
            Case lnIdReembolsoFarmacia
               oRsCabecera.Filter = "Motivo='F'"
            Case lnIdReembolsoServicios
               oRsCabecera.Filter = "Motivo='S'"
            Case Else
               oRsCabecera.Filter = "Motivo='SF'"
            End Select
            If oRsCabecera.RecordCount > 0 Then
                oRsCabecera.MoveFirst
                Do While Not oRsCabecera.EOF
'                  oRsCabecera!Motivo = IIf(rsItems!PuntoCarga = lcXfarmacia, "F", "S")
    '                  oRsCabecera!Paciente = rsItems!Codigo
    '                  oRsCabecera!ServicioHosp = Left(rsItems!Item, 100)
    '                  oRsCabecera!idUsuario = rsItems!Cantidad
    '                  oRsCabecera!Estancia = CCur(rsItems!Precio)
    '                  oRsCabecera!NroCuenta = rsItems!ml_idCuentaAtencion
    '                  oRsCabecera!recibe = Trim(Str(rsItems!idProducto))
                         lnImporteItem = Round(oRsCabecera!idUsuario * Val(oRsCabecera!Estancia), 2)
                         rsreporte.AddNew
                         rsreporte.Fields!IdComprobantePago = rsRsBoletaDetallada.Fields!IdComprobantePago
                         rsreporte.Fields!idUsuarioExonera = IIf(IsNull(rsRsBoletaDetallada.Fields!idUsuarioExonera), 0, rsRsBoletaDetallada.Fields!idUsuarioExonera)
                         rsreporte.Fields!Codigo = Trim(oRsCabecera!Paciente)
                         rsreporte.Fields!NombreProducto = Left(oRsCabecera!ServicioHosp, 250)
                         rsreporte.Fields!TotalBoleta = rsRsBoletaDetallada.Fields!TotalBoleta
                         rsreporte.Fields!idCuentaAtencion = IIf(IsNull(rsRsBoletaDetallada.Fields!idCuentaAtencion), 0, rsRsBoletaDetallada.Fields!idCuentaAtencion)
                         rsreporte.Fields!IdOrden = rsRsBoletaDetallada.Fields!IdOrden
                         rsreporte.Fields!idPreVenta = 0
                         rsreporte.Fields!IdOrdenPago = rsRsBoletaDetallada.Fields!IdOrdenPago
                         rsreporte.Fields!IdCajero = rsRsBoletaDetallada.Fields!IdCajero
                         rsreporte.Fields!IdTipoPago = rsRsBoletaDetallada.Fields!IdTipoPago
                         rsreporte.Fields!idEstadoComprobante = rsRsBoletaDetallada.Fields!idEstadoComprobante
                         rsreporte.Fields!razonSocial = rsRsBoletaDetallada.Fields!razonSocial
                         rsreporte.Fields!NroHistoriaClinica = IIf(IsNull(rsRsBoletaDetallada!NroHistoriaClinica), 0, rsRsBoletaDetallada!NroHistoriaClinica)
                         rsreporte.Fields!FechaCobranza = rsRsBoletaDetallada.Fields!FechaCobranza
                         rsreporte.Fields!Adelantos = rsRsBoletaDetallada.Fields!Adelantos
                         rsreporte.Fields!TotalPorPagar = lnImporteItem ' rsRsBoletaDetallada.Fields!TotalPorPagar
                         rsreporte.Fields!Cantidad = oRsCabecera!idUsuario
                         rsreporte.Fields!PrecioUnitario = Val(oRsCabecera!Estancia)
                         rsreporte.Fields!Observaciones = rsRsBoletaDetallada.Fields!Observaciones
                         rsreporte.Fields!exoneraciones = rsRsBoletaDetallada.Fields!exoneraciones
                         rsreporte.Fields!nombreCaja = rsRsBoletaDetallada.Fields!nombreCaja
                         rsreporte.Fields!IdCaja = rsRsBoletaDetallada.Fields!IdCaja
                         rsreporte.Fields!SerieImpresoraDefault = IIf(IsNull(rsRsBoletaDetallada!SerieImpresoraDefault), "", rsRsBoletaDetallada!SerieImpresoraDefault)
                         rsreporte.Fields!SerieImpresora2 = IIf(IsNull(rsRsBoletaDetallada!SerieImpresora2), "", rsRsBoletaDetallada!SerieImpresora2)
                         rsreporte.Fields!IdTipoOrden = rsRsBoletaDetallada.Fields!IdTipoOrden
                         rsreporte.Fields!DNI = IIf(IsNull(rsRsBoletaDetallada.Fields!DNI), "", rsRsBoletaDetallada.Fields!DNI)
                         rsreporte!IdTipoComprobante = rsRsBoletaDetallada!IdTipoComprobante
                         rsreporte!idProducto = Val(oRsCabecera!recibe)
                         rsreporte!Subtotal = rsRsBoletaDetallada!Subtotal
                         rsreporte!IGV = rsRsBoletaDetallada!IGV
                         rsreporte.Update
                         oRsCabecera.MoveNext
                Loop
            End If
            rsRsBoletaDetallada.MoveNext
        Loop
        Set oRsCabecera = Nothing
End Sub


'***************daniel barrantes**************
'***************Imprime una BOLETA (FORMATO continuo) y el CABEZAL DE IMPRESORA pasa a la
'***************Siguiente BOLETA
Sub ImpresionBoletaEnDosTYPE(lbMuestraVistaPrevia As Boolean, lcRucDireccionProveedor As String, lcVuelto As String, _
                             nroSerie As String, nroDcto As String, lbEsTicketSullana As Boolean, lcFarmaciaServicio As Long, _
                             lbEsBoletaPorPagoDeCuentaHospEmergSoloDeServicio As Boolean, _
                             Optional lbEsCopiaDeBoleta As Boolean, Optional lbGeneraArchivoTexto As Boolean)
Dim lbEsTicketFACTURA As Boolean
Dim lnLineaSiguienteIMPR As Long: Dim lnNumLetrasIMPR As Long
Dim rsreporte As New ADODB.Recordset: Dim rsReporte1 As New ADODB.Recordset
Dim rsRsBoletaDetallada As New ADODB.Recordset
Dim lnTotal As Double: Dim lnImporteEXO As Double: Dim lnSubTotal As Double
Dim lcAdelantos As String: Dim lcCuenta As String: Dim lcCajero As String
Dim lnFila As Long: Dim lnCant As Integer: Dim lnPrimeraFila As Long
Dim mo_ReporteUtil As New ReporteUtil
Dim oReglasCaja As New SIGHNegocios.ReglasCaja
Dim NúmeroRegistro As Integer    ' Declara variables.
Dim MiRegistroImpresion As ImpresionDOS
Dim lcCantidad As String: Dim lcPrecioUnitario As String
Dim lcSubTotal As String: Dim lcCajero1 As String
Dim lcProducto As String: Dim lcArchivo As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lbSigue As Boolean: Dim lcPuntoCarga As String
Dim lcNombrePuntoCarga As String: Dim lnAdelantos As Double
Dim lnNumeroMaxItemBoleta As Long, ldFechaBoleta As Date
Dim lcNombreCaja As String, lcNombreServicio As String
Dim lnIdCaja As Long, lcIdUsuarioExonera As String
Dim lcSerieImpresoraDefault As String, lcSerieImpresora2 As String
Dim lcSql As String, oRsTmp1 As New Recordset
Dim lbElDocumentoEsTicket As Boolean
Dim lbElDocumentoEsFactura As Boolean, lbReembolso_ImprimeDetalleItems As Boolean
Dim lnVigv As Double, lnVsubTotal As Double, lcCodigoPr As String, lnIGVfactura As Double
Dim oConexion As New Connection
        On Error GoTo ManejadorError
        oConexion.CommandTimeout = 300
        oConexion.Open sighEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        
        If lcFarmaciaServicio = sighEntidades.sghbien Then
           Set rsRsBoletaDetallada = oReglasCaja.CajaComprobantePagoProductosPorNroSerieNroDocumentoConexion(nroSerie, nroDcto, oConexion)
        Else
           Set rsRsBoletaDetallada = oReglasCaja.CajaComprobantePagoServiciosPorNroSerieNroDocumentoConexion(nroSerie, nroDcto, oConexion)
        End If
        '
        lnIGVfactura = -1
        lnNumeroMaxItemBoleta = Val(wxParametro102)
        lbElDocumentoEsTicket = IIf(rsRsBoletaDetallada.Fields!IdTipoComprobante = 4, True, False)
        lbElDocumentoEsFactura = IIf(rsRsBoletaDetallada.Fields!IdTipoComprobante = 2, True, False)
        'debb-2017
        If rsRsBoletaDetallada.Fields!IdTipoComprobante = 3 And wxParametro527 = "S" Then   'Es una BOLETA, pero se Imprime como TICKET
           lbElDocumentoEsTicket = True
        ElseIf rsRsBoletaDetallada.Fields!IdTipoComprobante = 2 And wxParametro543 = "S" Then   'Es una FACTURA, pero se Imprime como TICKET
           'lbEsTicketFACTURA = True
           lbElDocumentoEsTicket = True   ' False
        End If
        lnVigv = IIf(IsNull(rsRsBoletaDetallada!IGV), 0, rsRsBoletaDetallada!IGV)
        lnVsubTotal = rsRsBoletaDetallada!Subtotal
        
        If rsRsBoletaDetallada.Fields!IdTipoComprobante = 2 And lcFarmaciaServicio = sighEntidades.sghbien And (wxParametro208 = "03543" Or wxParametro208 = "03655") Then
          lnIGVfactura = lnVigv
        End If
        
        
        '**************************SOLO H.R.CAJAMARCA-inicio*******************************
        If wxParametro208 = "07686" Then
            If lcFarmaciaServicio = sighEntidades.sghbien Then
                'HRC boleta FARMACIA (Preventas/PagarCuentaTotal(solo FArmacia))*******************
                lnNumeroMaxItemBoleta = Val(wxParametro285)  ' Val(lcBuscaParametro.SeleccionaFilaParametro(285))
            ElseIf lbEsBoletaPorPagoDeCuentaHospEmergSoloDeServicio = True Then
                'HRC boleta OPCION DE CAJA: "PAGAR CUENTA TOTAL" (SOLO SERVICIOS)*******************
                lnNumeroMaxItemBoleta = Val(wxParametro286)  'Val(lcBuscaParametro.SeleccionaFilaParametro(286))
            End If
        End If
        '
        Dim lcCodigoPrd As String, lcNombrePrd As String
        lcCodigoPrd = ""
        lcNombrePrd = ""
        If rsRsBoletaDetallada.RecordCount = 1 Then
           rsRsBoletaDetallada.MoveFirst
           If rsRsBoletaDetallada!idProducto = Val(wxParametro549) Then
                mo_AdminServComunes.DevuelveCodigoYdescripcionSegunId Val(wxParametro549), lcCodigoPrd, lcNombrePrd
           End If
        End If
        '
        lbReembolso_ImprimeDetalleItems = Reembolso_ChequeaSiImprimeDetalleItems(rsRsBoletaDetallada)
        '
        If rsRsBoletaDetallada.RecordCount > lnNumeroMaxItemBoleta Or lcCodigoPrd <> "" Or lbReembolso_ImprimeDetalleItems = True Then
            With rsreporte
                .Fields.Append "TotalBoleta", adCurrency
                .Fields.Append "dctos", adCurrency
                .Fields.Append "idCuentaAtencion", adInteger
                .Fields.Append "IdOrden", adInteger
                .Fields.Append "IdOrdenPago", adInteger
                .Fields.Append "IdPreVenta", adInteger
                .Fields.Append "IdCajero", adInteger
                .Fields.Append "idTipoPago", adInteger
                .Fields.Append "idEstadoComprobante", adInteger
                .Fields.Append "razonSocial", adVarChar, 500, adFldIsNullable
                .Fields.Append "NroHistoriaClinica", adInteger
                .Fields.Append "FechaCobranza", adDate
                .Fields.Append "totalPorPagar", adCurrency
                .Fields.Append "cantidad", adInteger
                .Fields.Append "precioUnitario", adCurrency
                .Fields.Append "NombreProducto", adVarChar, 500, adFldIsNullable
                .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
                .Fields.Append "exoneraciones", adCurrency
                .Fields.Append "adelantos", adCurrency
                .Fields.Append "Observaciones", adVarChar, 250, adFldIsNullable
                .Fields.Append "nombreCaja", adVarChar, 100, adFldIsNullable
                .Fields.Append "idCaja", adInteger
                .Fields.Append "idComprobantePago", adInteger
                .Fields.Append "idUsuarioExonera", adInteger
                .Fields.Append "SerieImpresoraDefault", adVarChar, 20, adFldIsNullable
                .Fields.Append "SerieImpresora2", adVarChar, 20, adFldIsNullable
                .Fields.Append "idTipoOrden", adInteger
                .Fields.Append "dni", adVarChar, 11, adFldIsNullable
                .Fields.Append "IdTipoComprobante", adInteger
                .Fields.Append "idProducto", adInteger
                .Fields.Append "subTotal", adDouble
                .Fields.Append "igv", adDouble
                .LockType = adLockOptimistic
                .Open
            End With
            If lbReembolso_ImprimeDetalleItems = True Then
               Reembolso_CargaItemsDeCuenta rsRsBoletaDetallada, rsreporte, oConexion
            ElseIf rsRsBoletaDetallada.RecordCount > 0 Then
               rsRsBoletaDetallada.MoveFirst
               Do While Not rsRsBoletaDetallada.EOF
                  If lcFarmaciaServicio = sighEntidades.sghbien Then
                     lcPuntoCarga = "1"
                     lcNombrePuntoCarga = "FARMACIA"
                  Else
                     If lcCodigoPrd = "" Then
                        Set rsReporte1 = oReglasCaja.FactOrdenServicioSeleccionarPuntoCargaPorIdOrdenConexion(rsRsBoletaDetallada.Fields!IdOrden, oConexion)
                        lcPuntoCarga = "0"
                        lcNombrePuntoCarga = ".."
                        If rsReporte1.RecordCount > 0 Then
                           lcPuntoCarga = Trim(Str(rsReporte1.Fields!idPuntoCarga))
                           lcNombrePuntoCarga = Trim(rsReporte1.Fields!dPuntoCArga)
                        End If
                     Else
                        lcPuntoCarga = lcCodigoPrd
                        lcNombrePuntoCarga = rsRsBoletaDetallada!Observaciones
                     End If
                  End If
                  lbSigue = True
                  If rsreporte.RecordCount > 0 Then
                     rsreporte.MoveFirst
                     rsreporte.Find "codigo='" & lcPuntoCarga & "'"
                     If Not rsreporte.EOF Then
                        lbSigue = False
                     End If
                  End If
                  If lbSigue = True Then
                        rsreporte.AddNew
                        rsreporte.Fields!IdComprobantePago = rsRsBoletaDetallada.Fields!IdComprobantePago
                        rsreporte.Fields!idUsuarioExonera = IIf(IsNull(rsRsBoletaDetallada.Fields!idUsuarioExonera), 0, rsRsBoletaDetallada.Fields!idUsuarioExonera)
                        rsreporte.Fields!Codigo = lcPuntoCarga
                        rsreporte.Fields!NombreProducto = Left(lcNombrePuntoCarga, 250)
                        rsreporte.Fields!TotalBoleta = rsRsBoletaDetallada.Fields!TotalBoleta
                        rsreporte.Fields!idCuentaAtencion = IIf(IsNull(rsRsBoletaDetallada.Fields!idCuentaAtencion), 0, rsRsBoletaDetallada.Fields!idCuentaAtencion)
                        If lcFarmaciaServicio = sighEntidades.sghbien Then
                            If IsNull(rsRsBoletaDetallada.Fields!idPreVenta) Then
                                rsreporte.Fields!IdOrden = rsRsBoletaDetallada.Fields!IdOrden
                            Else
                                rsreporte.Fields!idPreVenta = rsRsBoletaDetallada.Fields!idPreVenta
                            End If
                        Else
                            rsreporte.Fields!IdOrdenPago = rsRsBoletaDetallada.Fields!IdOrdenPago
                        End If
                        rsreporte.Fields!IdCajero = rsRsBoletaDetallada.Fields!IdCajero
                        rsreporte.Fields!IdTipoPago = rsRsBoletaDetallada.Fields!IdTipoPago
                        rsreporte.Fields!idEstadoComprobante = rsRsBoletaDetallada.Fields!idEstadoComprobante
                        rsreporte.Fields!razonSocial = rsRsBoletaDetallada.Fields!razonSocial
                        rsreporte.Fields!NroHistoriaClinica = rsRsBoletaDetallada.Fields!NroHistoriaClinica
                        rsreporte.Fields!FechaCobranza = rsRsBoletaDetallada.Fields!FechaCobranza
                        rsreporte.Fields!Adelantos = rsRsBoletaDetallada.Fields!Adelantos
                        rsreporte.Fields!TotalPorPagar = rsRsBoletaDetallada.Fields!TotalPorPagar
                        
                        If lcCodigoPrd = "" Then
                            rsreporte.Fields!Cantidad = 0
                            rsreporte.Fields!PrecioUnitario = 0
                            rsreporte.Fields!Observaciones = rsRsBoletaDetallada.Fields!Observaciones
                        Else
                            rsreporte.Fields!Cantidad = rsRsBoletaDetallada!Cantidad
                            rsreporte.Fields!PrecioUnitario = rsRsBoletaDetallada!PrecioUnitario
                            rsreporte.Fields!Observaciones = " "
                            lnIGVfactura = rsRsBoletaDetallada!IGV
                        End If
                        rsreporte.Fields!exoneraciones = rsRsBoletaDetallada.Fields!exoneraciones
                        rsreporte.Fields!nombreCaja = rsRsBoletaDetallada.Fields!nombreCaja
                        rsreporte.Fields!IdCaja = rsRsBoletaDetallada.Fields!IdCaja
                        rsreporte.Fields!SerieImpresoraDefault = rsRsBoletaDetallada.Fields!SerieImpresoraDefault
                        rsreporte.Fields!SerieImpresora2 = rsRsBoletaDetallada.Fields!SerieImpresora2
                        rsreporte.Fields!IdTipoOrden = rsRsBoletaDetallada.Fields!IdTipoOrden
                        rsreporte.Fields!DNI = IIf(IsNull(rsRsBoletaDetallada.Fields!DNI), "", rsRsBoletaDetallada.Fields!DNI)
                        rsreporte!IdTipoComprobante = rsRsBoletaDetallada!IdTipoComprobante
                        rsreporte!idProducto = rsRsBoletaDetallada!idProducto
                        rsreporte!Subtotal = rsRsBoletaDetallada!Subtotal
                        rsreporte!IGV = rsRsBoletaDetallada!IGV
                  Else
                        rsreporte.Fields!TotalPorPagar = rsreporte.Fields!TotalPorPagar + rsRsBoletaDetallada.Fields!TotalPorPagar
                        rsreporte.Fields!Cantidad = 0
                        rsreporte.Fields!PrecioUnitario = 0
                  End If
                  rsreporte.Update
                  rsRsBoletaDetallada.MoveNext
               Loop
            End If
        Else
            Set rsreporte = rsRsBoletaDetallada.Clone
        End If
        If rsreporte.RecordCount > 0 Then
            lnIdCaja = rsreporte.Fields!IdCaja
            
            Dim retval
            Dim PosIniX, PosIniY, TempY, lnCmY
            If lcFarmaciaServicio = sighEntidades.sghbien Then
               If IsNull(rsreporte.Fields!idPreVenta) Then
                   lcCuenta = "Ord.Pag: " & Trim(Str(rsreporte.Fields!IdOrden)) & "      N°Cta: " & Trim(Str(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion)))
               Else
                   lcCuenta = "PreVenta: " & Trim(Str(rsreporte.Fields!idPreVenta)) & "      N°Cta: " & Trim(Str(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion)))
               End If
            Else
                lcCuenta = "Ord.Pag: " & Trim(Str(rsreporte.Fields!IdOrdenPago)) & "      N°Cta: " & Trim(Str(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion)))
            End If
            '************** contenido de boleta.bat
            '**************      NET USE LPT1:\\OLIDATA\EPSONFX880
            '**************          OLIDATA     <-NOMBRE DE  LA PC
            '**************          EPSONFX880  <-NOMBRE DE LA  IMPRESORA, QUE DEBERA ESTAR COMPARTIDA
            '**************
            '**************       type c:\boleta.txt >prn
            '**************
            '************** graba archivo boleta.txt
            lcArchivo = "c:\boleta.txt"
            Kill lcArchivo
            Open lcArchivo For Binary Access Write As #1
            rsreporte.MoveFirst
            lnTotal = IIf(IsNull(rsreporte.Fields!TotalBoleta), "0", rsreporte.Fields!TotalBoleta)
            lcAdelantos = "Adelantos    : " & Format(IIf(IsNull(rsreporte.Fields!Adelantos), 0, rsreporte.Fields!Adelantos), "######.#0")
            lnAdelantos = rsreporte.Fields!Adelantos
            'ypalomino 01102014
            lcSerieImpresoraDefault = IIf(IsNull(rsreporte.Fields!SerieImpresoraDefault), "", rsreporte.Fields!SerieImpresoraDefault)
            'lcSerieImpresora2 = rsReporte.Fields!SerieImpresora2
            'lcCuenta = "Ord.Pag: " & Trim(Str(rsReporte.Fields!IdOrden)) & "      N°Cta: " & Trim(Str(IIf(IsNull(rsReporte.Fields!idCuentaAtencion), 0, rsReporte.Fields!idCuentaAtencion)))
            '
            'If Val(wxParametro208) = 1910 Then
            If lbEsTicketSullana = True Then
               'SULLANA
               lcCajero1 = oReglasCaja.SeleccionaDatosCajeroConexion(rsreporte.Fields!IdCajero, sghApellidosYnombres, oConexion)
            Else
               lcCajero1 = oReglasCaja.SeleccionaDatosCajeroConexion(rsreporte.Fields!IdCajero, sghIniciales, oConexion)
            End If
            
            If lcFarmaciaServicio = sighEntidades.sghbien Then
               If Not IsNull(rsreporte.Fields!IdComprobantePago) Then
                    Set oRsTmp1 = mo_ReglasFarmacia.farmPreVentaXidComprobantePago(rsreporte.Fields!IdComprobantePago, oConexion)
                    If oRsTmp1.RecordCount > 0 Then
                       If Not IsNull(oRsTmp1.Fields!idVendedor) Then
                          lcCajero1 = lcCajero1 & "/Far:" & oReglasCaja.SeleccionaDatosCajero(oRsTmp1.Fields!idVendedor, sghIniciales, oConexion)
                       End If
                    End If
                    oRsTmp1.Close
               End If
            End If

            '
            lnImporteEXO = rsreporte.Fields!exoneraciones
            ldFechaBoleta = rsreporte.Fields!FechaCobranza
            lcNombreCaja = rsreporte.Fields!nombreCaja
            lcNombreServicio = Left(oReglasCaja.NombreServicioPorCuentaAtencionConexion(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion), oConexion) + String(33, " "), 33)
            lcIdUsuarioExonera = ""
            If lnImporteEXO > 0 And rsreporte.Fields!idUsuarioExonera > 0 Then
               lcIdUsuarioExonera = oReglasCaja.SeleccionaDatosCajeroConexion(rsreporte.Fields!idUsuarioExonera, sghIniciales, oConexion)
               If rsreporte.Fields!idCuentaAtencion > 0 Then
                  lcCajero1 = lcCajero1 & "/Ex:" & lcIdUsuarioExonera & "/N°Exo " & mo_ReglasFacturacion.NroExoneracionXcuenta(rsreporte.Fields!idCuentaAtencion, oConexion) & "/"
               Else
                  lcCajero1 = lcCajero1 & "/Ex:" & lcIdUsuarioExonera
               End If
            End If
            'genera TXT para aquellos que usan FORMATO AUTOMATICO DE BOLETAS/FACTURAS
            If lbGeneraArchivoTexto = True And wxParametro288 = "S" Then
                lcCajero$ = Chr(15) + String(15, " ") & IIf(rsreporte.Fields!IdTipoPago = 2, "DEVOLUCION", IIf(rsreporte.Fields!TotalBoleta = 0, "EXONERADO", "..")) & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                lcCajero$ = String(15, " ") & IIf(rsreporte.Fields!idEstadoComprobante = 9, "A N U L A D O ", "..") & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                lcCajero$ = String(8, " ") & "RUC: " & Left(Trim(IIf(IsNull(rsreporte.Fields!ruc), "..", rsreporte.Fields!ruc)) + String(33, " "), 33)
                Put #1, , lcCajero$
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                lcCajero$ = String(8, " ") & "Razón Social: " & Left(Trim(IIf(IsNull(rsreporte.Fields!razonSocial), "..", rsreporte.Fields!razonSocial)) + IIf(IsNull(rsreporte.Fields!NroHistoriaClinica), "(Particular)", "(HC: " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsreporte.Fields!NroHistoriaClinica)), False) & ")") + String(33, " "), 33)
                Put #1, , lcCajero$
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                lcCajero$ = String(5, " ") & "Factura N° " & nroSerie & " - " & nroDcto & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                lcCajero$ = String(5, " ") & "Fecha: " & Format(rsreporte.Fields!FechaCobranza, sighEntidades.DevuelveFechaSoloFormato_DMY) & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                NúmeroRegistro = 1
                Do While Not rsreporte.EOF
                   lnSubTotal = rsreporte.Fields!TotalPorPagar
                   lcCantidad = Format(rsreporte.Fields!Cantidad, "#####")
                   lcPrecioUnitario = Format(rsreporte.Fields!PrecioUnitario, "#####.#0")
                   lcSubTotal = Format(lnSubTotal, "######.#0")
                   lcProducto = Left(rsreporte.Fields!NombreProducto, 35)
                   lcCodigoPr = Left(Trim(rsreporte!Codigo), 5)
                   lcCajero$ = lcCodigoPr + String(6 - Len(lcCodigoPr), " ") & _
                               "  " + Left(lcProducto, 35) + String(35 - Len(lcProducto), " ")
                   Put #1, , lcCajero$
                   lcCajero$ = String$(5, " ") + lcCantidad + String(5 - Len(lcCantidad), " ") & _
                               lcPrecioUnitario + String$(8 - Len(lcPrecioUnitario), " ") & _
                               lcSubTotal + String$(10 - Len(lcSubTotal), " ") & _
                               Chr(13) + Chr(10)
                   Put #1, , lcCajero$
                   NúmeroRegistro = NúmeroRegistro + 1
                   rsreporte.MoveNext
                Loop
                
                If NúmeroRegistro < 12 Then
                   Do While NúmeroRegistro < 12
                      lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                      Put #1, , lcCajero$
                      NúmeroRegistro = NúmeroRegistro + 1
                   Loop
                End If
                lcCajero$ = String(3, " ") & "Cuenta " & lcCuenta & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                lcCajero$ = String(7, " ") & "SubTotal: " & Format(lnVsubTotal, "######.#0") & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                lcCajero$ = String(7, " ") & "Igv: " & Format(lnVigv, "######.#0") & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                lcCajero$ = String(7, " ") & "Total: " & Format(lnTotal, "######.#0") & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                lcCajero$ = String(3, " ") & "Son: " & sighEntidades.Numlet(sighEntidades.DevuelveNumeroSinDecimales(lnTotal)) & " con " & sighEntidades.DevuelveSoloDecimales(lnTotal) & "/100   Soles" & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                
                Close #1
            End If
            '
            '********FORMATO AUTOMATICO DE BOLETA (inicio)
            'Previamente se debe haber definido para la PC donde se emitirá la Boleta:
            '  **Formato de Papel
            '  **Numero de Lineas (parametro=102)
            If wxParametro288 = "S" Then
               Dim oRptBoleta As New RptBoleta
               oRptBoleta.ImprimeBoletaEnFormatoUsuario lbEsTicketFACTURA, lcRucDireccionProveedor, lcVuelto, rsreporte, lcCajero1, _
                                                        lcNombreServicio, lnTotal, lnImporteEXO, lcCuenta, lcAdelantos, lcSerieImpresoraDefault, lcSerieImpresora2, lbMuestraVistaPrevia, _
                                                        nroSerie & " - " & nroDcto, _
                                                        IIf(lcFarmaciaServicio = sighEntidades.sghbien, True, False), _
                                                        lbElDocumentoEsTicket, lnVsubTotal, lnVigv, _
                                                        lbElDocumentoEsFactura, oConexion, lbEsTicketSullana, _
                                                        lbEsCopiaDeBoleta
               Set oRptBoleta = Nothing
               Exit Sub
            End If
            '********FORMATO AUTOMATICO DE BOLETA (fin)
            '
            Select Case wxParametro208
            Case "05195"                  'Hospital Belen Trujillo
                'columna 15,fila 0
                lcCajero$ = Chr(15) + String(15, " ") & IIf(rsreporte.Fields!IdTipoPago = 2, "DEVOLUCION", IIf(rsreporte.Fields!TotalBoleta = 0, "EXONERADO", "..")) & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                'columna 15,fila 1
                lcCajero$ = String(15, " ") & IIf(rsreporte.Fields!idEstadoComprobante = 9, "A N U L A D O ", "..") & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                'columna 1,fila 2  (vacia)
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                'columna 8,fila 3
                lcCajero$ = String(8, " ") & Left(Trim(IIf(IsNull(rsreporte.Fields!razonSocial), "..", rsreporte.Fields!razonSocial)) + IIf(IsNull(rsreporte.Fields!NroHistoriaClinica), "(Particular)", "(HC: " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsreporte.Fields!NroHistoriaClinica)), False) & ")") + String(33, " "), 33)
                Put #1, , lcCajero$
                'columna 49,fila 3
                lcCajero$ = String(5, " ") & nroSerie & " - " & nroDcto & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                'columna 8,fila 4
                'lcCajero$ = String(8, " ") & Left(oReglasCaja.NombreServicioPorCuentaAtencion(IIf(IsNull(rsReporte.Fields!idCuentaAtencion), 0, rsReporte.Fields!idCuentaAtencion)) + String(33, " "), 33)
                lcCajero$ = String(8, " ") & Left(oReglasCaja.NombreServicioPorCuentaAtencionConexion(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion), oConexion) + String(33, " "), 33)
                Put #1, , lcCajero$
                'columna 40,fila 4
                lcCajero$ = String(5, " ") & Format(rsreporte.Fields!FechaCobranza, sighEntidades.DevuelveFechaSoloFormato_DMY) & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                'columna 1,fila 5  (vacia)
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                'columna 1,fila 6  (vacia)
                'lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                'Put #1, , lcCajero$
                
                
                
                
                NúmeroRegistro = 1
                Do While Not rsreporte.EOF
                   lnSubTotal = rsreporte.Fields!TotalPorPagar
                   lcCantidad = Format(rsreporte.Fields!Cantidad, "#####")
                   lcPrecioUnitario = Format(rsreporte.Fields!PrecioUnitario, "#####.#0")
                   lcSubTotal = Format(lnSubTotal, "######.#0")
                   lcProducto = Left(rsreporte.Fields!NombreProducto, 35)
                   
                   
                   
                   lcCajero$ = rsreporte.Fields!Codigo + String(6 - Len(rsreporte.Fields!Codigo), " ") & _
                               "  " + Left(lcProducto, 35) + String(35 - Len(lcProducto), " ")
                   Put #1, , lcCajero$
                   lcCajero$ = String$(5, " ") + lcCantidad + String(5 - Len(lcCantidad), " ") & _
                               lcPrecioUnitario + String$(8 - Len(lcPrecioUnitario), " ") & _
                               lcSubTotal + String$(10 - Len(lcSubTotal), " ") & _
                               Chr(13) + Chr(10)
                   Put #1, , lcCajero$
                   NúmeroRegistro = NúmeroRegistro + 1
                   rsreporte.MoveNext
                Loop
                
                If NúmeroRegistro < 12 Then
                   Do While NúmeroRegistro < 12
                      'columna 1,fila X  (vacia)
                      lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                      Put #1, , lcCajero$
                      NúmeroRegistro = NúmeroRegistro + 1
                   Loop
                End If
                'columna 10,fila X
                lcCajero$ = String(3, " ") & lcCajero1 & lcCuenta & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                lcCajero$ = String(10, " ") & Left(lcAdelantos & String(22, " "), 22)
                Put #1, , lcCajero$
                'columna 35,fila X+2
                lcCajero$ = String(7, " ") & Format(lnTotal + lnImporteEXO, "######.#0")
                Put #1, , lcCajero$
                'columna 45,fila X+2
                lcCajero$ = String(7, " ") & Format(lnImporteEXO, "######.#0")
                Put #1, , lcCajero$
                'columna 53,fila X+2
                lcCajero$ = String(7, " ") & Format(lnTotal, "######.#0") & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                'columna 1,fila X  (vacia)
                lcCajero$ = String(3, " ") & "Son: " & sighEntidades.Numlet(sighEntidades.DevuelveNumeroSinDecimales(lnTotal)) & " con " & sighEntidades.DevuelveSoloDecimales(lnTotal) & "/100   Soles" & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                'columna 1,fila X  (vacia)
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                'columna 1,fila X  (vacia)
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                Put #1, , lcCajero$
                
                Close #1
                rsreporte.Close
                '************** imprime boleta.txt
                retval = Shell(App.Path & "\plantillas\boleta.bat", 1)
            Case "03543", "03655"                           'HospitalRegionalAyacucho/Hospital HUANTA
                lcCodigoPr = rsreporte.Fields!NombreProducto
                PosIniX = 1.2
                PosIniY = 0.6
                Printer.Font = "Verdana"
                Printer.FontSize = 8
                Printer.ScaleMode = vbCentimeters '7
                'columna 15,fila 0
                lcCajero$ = IIf(rsreporte.Fields!IdTipoPago = 2, "DEVOLUCION", IIf(rsreporte.Fields!TotalBoleta = 0, "EXONERADO", ".."))  'Chr(15) + String(35, " ") & IIf(rsReporte.Fields!IdTipoPago = 2, "DEVOLUCION", IIf(rsReporte.Fields!TotalBoleta = 0, "EXONERADO", "..")) & Chr(13) & Chr(10)
                Printer.CurrentY = PosIniY
                Printer.CurrentX = PosIniX + 5.5
                Printer.Print lcCajero$ ' Put #1, , lcCajero$
                'columna 15,fila 1
                lcCajero$ = IIf(rsreporte.Fields!idEstadoComprobante = 9, "A N U L A D O ", "..") 'String(35, " ") & IIf(rsReporte.Fields!IdEstadoComprobante = 9, "A N U L A D O ", "..") & Chr(13) & Chr(10)
                Printer.CurrentX = PosIniX + 5.5
                Printer.Print lcCajero$  'Put #1, , lcCajero$
                'columna 1,fila 2  (vacia)
                lcCajero$ = String(45, " ") 'String(45, " ") & Chr(13) & Chr(10)
'                Printer.Print 'Put #1, , lcCajero$
                'columna 1,fila 2  (vacia)
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
'                Printer.Print 'Put #1, , lcCajero$
                'columna 1,fila 2  (vacia)
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
'                Printer.Print 'Put #1, , lcCajero$
                'columna 1,fila 2  (vacia)
                'lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                lcCajero$ = lcCuenta 'String(13, " ") & lcCuenta & Chr(13) & Chr(10)
                Printer.CurrentX = PosIniX + 2.4
                Printer.CurrentY = PosIniY + 2.4
                Printer.Print lcCajero$ 'Put #1, , lcCajero$
                'columna 8,fila 3
                lcCajero$ = Trim(IIf(IsNull(rsreporte.Fields!razonSocial), "..", rsreporte.Fields!razonSocial)) + IIf(IsNull(rsreporte.Fields!NroHistoriaClinica), "(Particular)", "(HC: " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsreporte.Fields!NroHistoriaClinica)), False) & ")") 'String(13, " ") & Left(Trim(IIf(IsNull(rsReporte.Fields!RazonSocial), "..", rsReporte.Fields!RazonSocial)) + IIf(IsNull(rsReporte.Fields!NroHistoriaClinica), "(Particular)", "(HC: " & rsReporte.Fields!NroHistoriaClinica & ")") + String(33, " "), 33)
                TempY = Printer.CurrentY
                Printer.CurrentX = PosIniX + 2.4
                Printer.Print lcCajero$ 'Put #1, , lcCajero$
                'columna 49,fila 3
                lcCajero$ = nroSerie & " - " & nroDcto  'String(20, " ") & nroSerie & " - " & nroDcto & Chr(13) & Chr(10)
                Printer.CurrentY = TempY
                Printer.CurrentX = PosIniX + 10.2
                Printer.Print lcCajero$ 'Put #1, , lcCajero$
                'columna 8,fila 4
                lcCajero$ = oReglasCaja.NombreServicioPorCuentaAtencionConexion(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion), oConexion)  'String(13, " ") & Left(oReglasCaja.NombreServicioPorCuentaAtencion(IIf(IsNull(rsReporte.Fields!idCuentaAtencion), 0, rsReporte.Fields!idCuentaAtencion)) + String(33, " "), 33)
                Printer.CurrentX = PosIniX + 2.4
                TempY = Printer.CurrentY
                Printer.Print lcCajero$ 'Put #1, , lcCajero$
                'columna 40,fila 4
                lcCajero$ = Format(rsreporte.Fields!FechaCobranza, "dd/MM/yyyy hh:mm") 'String(20, " ") & Format(rsReporte.Fields!FechaCobranza, "dd/MM/yyyy hh:mm") & Chr(13) & Chr(10)
                Printer.CurrentY = TempY
                Printer.CurrentX = PosIniX + 10.2
                Printer.Print lcCajero$  'Put #1, , lcCajero$
                'columna 1,fila 5  (vacia)
                lcCajero$ = IIf(IsNull(rsreporte.Fields!Observaciones), "", rsreporte.Fields!Observaciones) 'String(13, " ") & rsReporte.Fields!Observaciones & Chr(13) & Chr(10)
                Printer.CurrentX = PosIniX + 2.4
                Printer.Print lcCajero$ 'Put #1, , lcCajero$
                'columna 1,fila 5  (vacia)
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
'                Printer.Print 'Put #1, , lcCajero$
                'columna 1,fila 5  (vacia)
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
'                Printer.Print 'Put #1, , lcCajero$
                
                NúmeroRegistro = 1
                TempY = PosIniY + 4.8
                Printer.CurrentY = TempY
                Do While Not rsreporte.EOF
                   lnSubTotal = rsreporte.Fields!TotalPorPagar
                   lcCantidad = Format(rsreporte.Fields!Cantidad, "#####")
                   lcPrecioUnitario = Format(rsreporte.Fields!PrecioUnitario, "####0.0000")
                   lcSubTotal = Format(lnSubTotal, "#####0.00")
                   lcProducto = Left(rsreporte.Fields!NombreProducto, 48)
                   
                   
                   'lcCajero$ = Left(rsReporte.Fields!codigo, 5) + String(6 - Len(Trim(Left(rsReporte.Fields!codigo, 5))), " ") & "  " + Left(lcProducto, 48) + String(48 - Len(lcProducto), " ")                                'Left(rsReporte.Fields!codigo, 5) + String(6 - Len(Trim(Left(rsReporte.Fields!codigo, 5))), " ") & "  " + Left(lcProducto, 48) + String(48 - Len(lcProducto), " ")
                   lcCajero$ = rsreporte.Fields!Codigo
                   Printer.CurrentX = PosIniX + 0.2 '0.7
                   Printer.CurrentY = TempY
                   Printer.Print lcCajero$
                   lcCajero$ = Left(lcProducto, 52) 'Left(rsReporte.Fields!codigo, 5) + String(6 - Len(Trim(Left(rsReporte.Fields!codigo, 5))), " ") & "  " + Left(lcProducto, 48) + String(48 - Len(lcProducto), " ")
                   Printer.CurrentX = PosIniX + 1.4 '2.1
                   Printer.CurrentY = TempY
                   Printer.Print lcCajero$ 'Put #1, , lcCajero$
                   'lcCajero$ = String$(10, " ") + lcCantidad + String(5 - Len(lcCantidad), " ") & lcPrecioUnitario + String$(13 - Len(lcPrecioUnitario), " ") & lcSubTotal + String$(10 - Len(lcSubTotal), " ") 'String$(10, " ") + lcCantidad + String(5 - Len(lcCantidad), " ") & lcPrecioUnitario + String$(13 - Len(lcPrecioUnitario), " ") & lcSubTotal + String$(10 - Len(lcSubTotal), " ") & Chr(13) + Chr(10)
                   lcCajero$ = lcCantidad
                   Printer.CurrentX = PosIniX + 10 '9.6
                   Printer.CurrentY = TempY
                   Printer.Print lcCajero$ 'Put #1, , lcCajero$
                   lcCajero$ = lcPrecioUnitario
                   Printer.CurrentX = PosIniX + 11.9 - Printer.TextWidth(lcCajero$)
                   Printer.CurrentY = TempY
                   Printer.Print lcCajero$
                   lcCajero$ = lcSubTotal
                   Printer.CurrentX = PosIniX + 14.6 - Printer.TextWidth(lcCajero$)
                   Printer.CurrentY = TempY
                   Printer.Print lcCajero$
                   NúmeroRegistro = NúmeroRegistro + 1
                   rsreporte.MoveNext
                   TempY = Printer.CurrentY
                Loop
                
                If NúmeroRegistro < 14 Then
                   Do While NúmeroRegistro < 14
                      If lnIGVfactura <> -1 Then
                            lcProducto = Mid(lcCodigoPr, (48 * (NúmeroRegistro - 1)) + 1, 48)
                            If Len(lcProducto) > 0 Then
                               lcCajero$ = "     "
                               Printer.CurrentX = PosIniX + 0.2 '0.7
                               Printer.CurrentY = TempY
                               Printer.Print lcCajero$
                               lcCajero$ = lcProducto
                               Printer.CurrentX = PosIniX + 1.4 '2.1
                               Printer.CurrentY = TempY
                               Printer.Print lcCajero$
                               TempY = Printer.CurrentY
                            End If
                      End If
                      'columna 1,fila X  (vacia)
                      lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
'                      Printer.Print 'Put #1, , lcCajero$
                      NúmeroRegistro = NúmeroRegistro + 1
                   Loop
                End If
                'columna 10,fila X
                lcCajero$ = lcCajero1
                Printer.CurrentY = PosIniY + 10
                Printer.CurrentX = PosIniX + 0.6
                Printer.Print lcCajero$ 'Put #1, , lcCajero$
                lcCajero$ = "SUBTOTAL:   " & Format(lnTotal + lnImporteEXO + lnAdelantos, "#####0.00")
                Printer.CurrentY = PosIniY + 10
                Printer.CurrentX = PosIniX + 14.6 - Printer.TextWidth(lcCajero$)
                Printer.Print lcCajero$
                
                lcCajero$ = lcAdelantos  'String(3, " ") & Left(lcAdelantos & String(22, " "), 22)
                TempY = Printer.CurrentY
                Printer.CurrentX = PosIniX + 0.6
                Printer.Print lcCajero$ 'Put #1, , lcCajero$
                'columna 45,fila X+2
                If lnIGVfactura <> -1 Then
                    lcCajero$ = "IGV:   " & Format(lnIGVfactura, "#####0.00")
                Else
                    lcCajero$ = "DCTO:   " & Format(lnImporteEXO, "#####0.00")
                End If
                Printer.CurrentX = PosIniX + 14.6 - Printer.TextWidth(lcCajero$)
                Printer.CurrentY = TempY
                Printer.Print lcCajero$ 'Put #1, , lcCajero$
                'columna 1,fila X  (vacia)
                lcProducto = Trim("Son: " & sighEntidades.Numlet(sighEntidades.DevuelveNumeroSinDecimales(lnTotal)) & " con " & sighEntidades.DevuelveSoloDecimales(lnTotal) & "/100   Soles") 'Left("Son: " & sighEntidades.Numlet(sighEntidades.DevuelveNumeroSinDecimales(lnTotal)) & " con " & sighEntidades.DevuelveSoloDecimales(lnTotal) & "/100   Nuevos Soles", 60)
                lcCajero$ = lcProducto 'String(3, " ") & Left(lcProducto, 60) + String(60 - Len(lcProducto), " ")
                Printer.CurrentX = PosIniX + 0.6
                TempY = Printer.CurrentY
                Printer.Print lcCajero$ 'Put #1, , lcCajero$
                'columna 53,fila X+2
                lcCajero$ = "TOTAL:   " & Format(lnTotal, "#####0.00") 'String(11, " ") & "TOTAL:   " & Format(lnTotal, "######.#0") & Chr(13) & Chr(10)
                Printer.CurrentX = PosIniX + 14.6 - Printer.TextWidth(lcCajero$)
                Printer.CurrentY = TempY
                Printer.Print lcCajero$ 'Put #1, , lcCajero$
                'columna 1,fila X  (vacia)
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                'columna 1,fila X  (vacia)
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                'columna 1,fila X  (vacia)
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                'columna 1,fila X  (vacia)
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                'columna 1,fila X  (vacia)
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                'columna 1,fila X  (vacia)
                lcCajero$ = String(45, " ") & Chr(13) & Chr(10)
                Close #1
                rsreporte.Close
                Printer.EndDoc
            Case "01908", "17077"              'Hospital JAMO-Tumbes 'Actualizado 06102014
                If lcFarmaciaServicio = sighEntidades.sghbien Then  'Solo para Farmacia
                   SeteaOtraImpresoraDiferenteAlDefault mo_ReglasCaja.CajaDevuelveNombreImpresoraPorId(lnIdCaja, False)
                Else
                   SeteaOtraImpresoraDiferenteAlDefault mo_ReglasCaja.CajaDevuelveNombreImpresoraPorId(lnIdCaja, True)
                End If
                '
                PosIniX = 0.5
                PosIniY = 0.6
                lnCmY = 0.4
                Printer.Font = "Verdana"
                Printer.FontSize = 8
                Printer.ScaleMode = vbCentimeters
'                Printer.PaperSize = 256
'                Printer.Width = 13
'                Printer.Height = 9.2
                
                'columna 15,fila 0
                lcCajero$ = Chr(15) + String(35, " ") & IIf(rsreporte.Fields!IdTipoPago = 2, "DEVOLUCION", IIf(rsreporte.Fields!TotalBoleta = 0, "EXONERADO", "..")) '& Chr(13) & Chr(10)
                Printer.CurrentY = PosIniY
                Printer.CurrentX = PosIniX
                Printer.Print lcCajero$
                PosIniY = PosIniY + lnCmY
                'columna 15,fila 1
                lcCajero$ = String(35, " ") & IIf(rsreporte.Fields!idEstadoComprobante = 9, "A N U L A D O ", "..") '& Chr(13) & Chr(10)
                Printer.CurrentY = PosIniY
                Printer.CurrentX = PosIniX
                Printer.Print lcCajero$
                PosIniY = PosIniY + lnCmY
                'columna 1,fila 2  (vacia)
                lcCajero$ = String(45, " ") '& Chr(13) & Chr(10)
                Printer.CurrentY = PosIniY
                Printer.CurrentX = PosIniX
                Printer.Print lcCajero$
                PosIniY = PosIniY + lnCmY
                'columna 1,fila 2  (vacia)
                lcCajero$ = String(45, " ") '& Chr(13) & Chr(10)
                Printer.CurrentY = PosIniY
                Printer.CurrentX = PosIniX
                Printer.Print lcCajero$
                PosIniY = PosIniY + lnCmY
                'columna 8,fila 3
                lcCajero$ = String(13, " ") & Left(Trim(IIf(IsNull(rsreporte.Fields!razonSocial), "..", rsreporte.Fields!razonSocial)) + IIf(IsNull(rsreporte.Fields!NroHistoriaClinica), "(Particular)", "(HC: " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsreporte.Fields!NroHistoriaClinica)), False) & ")") + String(43, " "), 43) & String(25, " ") & nroSerie & " - " & nroDcto '& Chr(13) & Chr(10)
                Printer.CurrentY = PosIniY
                Printer.CurrentX = PosIniX
                Printer.Print lcCajero$
                PosIniY = PosIniY + lnCmY
                'columna 8,fila 4
                If IsNull(rsreporte.Fields!idCuentaAtencion) Then
                   lcCajero$ = ""
                Else
                   lcCajero$ = " (Cta: " & Trim(Str(rsreporte.Fields!idCuentaAtencion)) & ")"
                End If
                'lcCajero$ = String(13, " ") & Left(oReglasCaja.NombreServicioPorCuentaAtencion(IIf(IsNull(rsReporte.Fields!idCuentaAtencion), 0, rsReporte.Fields!idCuentaAtencion)) + String(40, " "), 40) & lcCajero$
                lcCajero$ = String(13, " ") & Left(oReglasCaja.NombreServicioPorCuentaAtencionConexion(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion), oConexion) + String(40, " "), 40) & lcCajero$
                Printer.CurrentY = PosIniY
                Printer.CurrentX = PosIniX
                Printer.Print lcCajero$
                PosIniY = PosIniY + lnCmY
                'columna 1,fila 5  (vacia)
                lcCajero$ = String(13, " ") & Left(rsreporte.Fields!Observaciones + String(40, " "), 40) '& String(28, " ") & Format(rsReporte.Fields!FechaCobranza, sighEntidades.DevuelveFechaSoloFormato_DMY_HM) '& Chr(13) & Chr(10)
                Printer.CurrentY = PosIniY
                Printer.CurrentX = PosIniX
                Printer.Print lcCajero$
                lcCajero$ = Format(rsreporte.Fields!FechaCobranza, sighEntidades.DevuelveFechaSoloFormato_DMY_HM)  '& Chr(13) & Chr(10)
                Printer.CurrentY = PosIniY
                Printer.CurrentX = PosIniX + 12 - Printer.TextWidth(lcCajero$)
                Printer.Print lcCajero$
                PosIniY = PosIniY + lnCmY
                'columna 1,fila 5  (vacia)
                lcCajero$ = String(45, " ") '& Chr(13) & Chr(10)
                Printer.CurrentY = PosIniY
                Printer.CurrentX = PosIniX
                Printer.Print lcCajero$
                PosIniY = PosIniY + lnCmY
                
                NúmeroRegistro = 1
                Do While Not rsreporte.EOF
                   lnSubTotal = rsreporte.Fields!TotalPorPagar
                   lcCantidad = Format(rsreporte.Fields!Cantidad, "#####")
                   lcPrecioUnitario = Format(rsreporte.Fields!PrecioUnitario, "#####.###0")
                   lcSubTotal = Format(lnSubTotal, "######.#0")
                   lcProducto = Left(rsreporte.Fields!NombreProducto, 40)
                   lcCajero$ = String$(3, " ") + lcCantidad + String(5 - Len(lcCantidad), " ") & _
                               Left(rsreporte.Fields!Codigo, 5) + String(6 - Len(Trim(Left(rsreporte.Fields!Codigo, 5))), " ") & _
                               "  " + Left(lcProducto, 40) + String(40 - Len(lcProducto), " ")
                   Printer.CurrentY = PosIniY
                   Printer.CurrentX = PosIniX
                   Printer.Print lcCajero$
                   'precio
                   lcCajero$ = lcPrecioUnitario
                   Printer.CurrentY = PosIniY
                   Printer.CurrentX = PosIniX + 10.5 - Printer.TextWidth(lcCajero$)
                   Printer.Print lcCajero$
                   'Total
                   lcCajero$ = lcSubTotal
                   Printer.CurrentY = PosIniY
                   Printer.CurrentX = PosIniX + 12.3 - Printer.TextWidth(lcCajero$)
                   Printer.Print lcCajero$
                   '
                   PosIniY = PosIniY + lnCmY
                   NúmeroRegistro = NúmeroRegistro + 1
                   rsreporte.MoveNext
                Loop
                
                If NúmeroRegistro < (lnNumeroMaxItemBoleta + 1) Then
                   Do While NúmeroRegistro < (lnNumeroMaxItemBoleta + 1)
                      'columna 1,fila X  (vacia)
                      lcCajero$ = String(45, " ") '& Chr(13) & Chr(10)
                      Printer.CurrentY = PosIniY
                      Printer.CurrentX = PosIniX
                      Printer.Print lcCajero$
                      PosIniY = PosIniY + lnCmY
                      NúmeroRegistro = NúmeroRegistro + 1
                   Loop
                End If
                'columna 10,fila X
                lcCajero$ = String(3, " ") & lcCajero1 & " (" & Left(lcAdelantos & String(22, " "), 22) & ") (Desc: " & Format(lnImporteEXO, "######.#0") & ")" & String(31, " ")
                Printer.CurrentY = PosIniY
                Printer.CurrentX = PosIniX
                Printer.Print lcCajero$
                lcCajero$ = "SUBTOTAL: " & Format(lnTotal + lnImporteEXO + lnAdelantos, "######.#0") '& Chr(13) & Chr(10)
                Printer.CurrentY = PosIniY
                Printer.CurrentX = PosIniX + 12.3 - Printer.TextWidth(lcCajero$)
                Printer.Print lcCajero$
                PosIniY = PosIniY + lnCmY
                'columna 1,fila X  (vacia)
                If NúmeroRegistro < 9 Then
                   Do While NúmeroRegistro < 9
                      'columna 1,fila X  (vacia)
                      lcCajero$ = String(45, " ") '& Chr(13) & Chr(10)
                      Printer.CurrentY = PosIniY
                      Printer.CurrentX = PosIniX
                      Printer.Print lcCajero$
                      PosIniY = PosIniY + lnCmY
                      NúmeroRegistro = NúmeroRegistro + 1
                   Loop
                End If
                lcProducto = Left("Son: " & sighEntidades.Numlet(sighEntidades.DevuelveNumeroSinDecimales(lnTotal)) & " con " & sighEntidades.DevuelveSoloDecimales(lnTotal) & "/100   Soles", 60)
                lcCajero$ = String(3, " ") & Left(lcProducto, 60) + String(60 - Len(lcProducto), " ")
                Printer.CurrentY = PosIniY
                Printer.CurrentX = PosIniX
                Printer.Print lcCajero$
                
                      lcCajero$ = String(45, " ") '& Chr(13) & Chr(10)
                      Printer.CurrentY = PosIniY
                      Printer.CurrentX = PosIniX
                      Printer.Print lcCajero$
                      PosIniY = PosIniY + lnCmY
                
                
                
                lcCajero$ = Format(lnTotal, "######.#0")  '& Chr(13) & Chr(10)
                Printer.CurrentY = PosIniY
                Printer.CurrentX = PosIniX + 12.3 - Printer.TextWidth(lcCajero$)
                Printer.Print lcCajero$
                '
                Close #1
                rsreporte.Close
                Printer.EndDoc
            Case "07686"  'Hospital Regional Cajamarca
                If lcFarmaciaServicio = sighEntidades.sghbien Then
                    '*****************HRC boleta FARMACIA (Preventas/AltaHospitalaria/AltaEmergencia)*******************
                    SeteaOtraImpresoraDiferenteAlDefault mo_ReglasCaja.CajaDevuelveNombreImpresoraPorId(lnIdCaja, False)
                    PosIniX = 0.5
                    PosIniY = 0.6
                    lnCmY = 0.4
                    Printer.Font = "Arial"
                    Printer.FontSize = 8
                    Printer.ScaleMode = vbCentimeters
                    lcCajero$ = Chr(15) + String(35, " ") & IIf(rsreporte.Fields!IdTipoPago = 2, "DEVOLUCION", IIf(rsreporte.Fields!TotalBoleta = 0, "EXONERADO", ".."))
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(35, " ") & IIf(rsreporte.Fields!idEstadoComprobante = 9, "A N U L A D O ", "..")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(45, " ")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(45, " ")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(45, " ")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(45, " ")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(15, " ") & Left(Trim(IIf(IsNull(rsreporte.Fields!razonSocial), "..", rsreporte.Fields!razonSocial)) + IIf(IsNull(rsreporte.Fields!NroHistoriaClinica), "(Particular)", "(HC: " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsreporte.Fields!NroHistoriaClinica)), False) & ")") + String(43, " "), 43) & String(5, " ") & nroSerie & " - " & nroDcto '& Chr(13) & Chr(10)
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    If IsNull(rsreporte.Fields!idCuentaAtencion) Then
                       lcCajero$ = ""
                    Else
                       lcCajero$ = " (Cta: " & Trim(Str(rsreporte.Fields!idCuentaAtencion)) & ")"
                    End If
                    'lcCajero$ = String(55, " ") & Left(oReglasCaja.NombreServicioPorCuentaAtencion(IIf(IsNull(rsReporte.Fields!idCuentaAtencion), 0, rsReporte.Fields!idCuentaAtencion)) + String(25, " "), 25) & lcCajero$
                    lcCajero$ = String(55, " ") & Left(oReglasCaja.NombreServicioPorCuentaAtencionConexion(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion), oConexion) + String(25, " "), 25) & lcCajero$
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(13, " ") & Left(rsreporte.Fields!Observaciones + String(40, " "), 40)
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(45, " ")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    
                    NúmeroRegistro = 1
                    Do While Not rsreporte.EOF
                       lnSubTotal = rsreporte.Fields!TotalPorPagar
                       lcCantidad = Format(rsreporte.Fields!Cantidad, "#####")
                       lcPrecioUnitario = Format(rsreporte.Fields!PrecioUnitario, "#####.###0")
                       lcSubTotal = Format(lnSubTotal, "######.#0")
                       lcProducto = Left(rsreporte.Fields!NombreProducto, 30)
                       lcCajero$ = String$(3, " ") + lcCantidad + String(5 - Len(lcCantidad), " ") & _
                                   Left(rsreporte.Fields!Codigo, 5) + String(6 - Len(Trim(Left(rsreporte.Fields!Codigo, 5))), " ") & _
                                   "  " + Left(lcProducto, 30) + String(30 - Len(lcProducto), " ")
                       Printer.CurrentY = PosIniY
                       Printer.CurrentX = PosIniX
                       Printer.Print lcCajero$
                       'precio
                       lcCajero$ = lcPrecioUnitario
                       Printer.CurrentY = PosIniY
                       Printer.CurrentX = PosIniX + 7.5 - Printer.TextWidth(lcCajero$)
                       Printer.Print lcCajero$
                       'Total
                       lcCajero$ = lcSubTotal
                       Printer.CurrentY = PosIniY
                       Printer.CurrentX = PosIniX + 9.3 - Printer.TextWidth(lcCajero$)
                       Printer.Print lcCajero$
                       '
                       PosIniY = PosIniY + lnCmY
                       NúmeroRegistro = NúmeroRegistro + 1
                       rsreporte.MoveNext
                    Loop
                    
                    If NúmeroRegistro < (lnNumeroMaxItemBoleta + 1) Then
                       Do While NúmeroRegistro < (lnNumeroMaxItemBoleta + 1)
                          'columna 1,fila X  (vacia)
                          lcCajero$ = String(45, " ") '& Chr(13) & Chr(10)
                          Printer.CurrentY = PosIniY
                          Printer.CurrentX = PosIniX
                          Printer.Print lcCajero$
                          PosIniY = PosIniY + lnCmY
                          NúmeroRegistro = NúmeroRegistro + 1
                       Loop
                    End If
                    '
                    lcProducto = Left("Son: " & sighEntidades.Numlet(sighEntidades.DevuelveNumeroSinDecimales(lnTotal)) & " con " & sighEntidades.DevuelveSoloDecimales(lnTotal) & "/100   Soles", 60)
                    lcCajero$ = String(3, " ") & Left(lcProducto, 60) + String(60 - Len(lcProducto), " ")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(18, " ") & Format(ldFechaBoleta, sighEntidades.DevuelveHoraSoloFormato_HM) & String(5, " ") & "SUBTOTAL:   " & Format(lnTotal + lnImporteEXO + lnAdelantos, "#####0.00") & String(5, " ") & "DCTO:   " & Format(lnImporteEXO, "#####0.00")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(18, " ") & Format(ldFechaBoleta, sighEntidades.DevuelveFechaSoloFormato_DMY) & String(70, " ") & Format(lnTotal, "######.#0")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(18, " ") & lcNombreCaja
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(18, " ") & lcCajero1
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    '
                    Close #1
                    rsreporte.Close
                ElseIf lbEsBoletaPorPagoDeCuentaHospEmergSoloDeServicio = True Then
                    '*****************HRC boleta OPCION DE CAJA: "PAGAR CUENTA TOTAL" (SOLO SERVICIOS)*******************
                    Dim lcMedicoEgreso As String, lcServicio As String, ldFIngreso As Date, ldFAlta As Date, lnEdad As Long, lcCama As String
                    lcMedicoEgreso = "": lcServicio = "": ldFIngreso = 0: ldFAlta = 0: lnEdad = 0: lcCama = ""
                    If rsreporte.Fields!idCuentaAtencion > 0 Then
                       'Set rsReporte1 = oReglasCaja.AtencionesSeleccionarDatosCabeceraBoletaHRCxCuenta(rsReporte.Fields!idCuentaAtencion)
                       Set rsReporte1 = oReglasCaja.AtencionesSeleccionarDatosCabeceraBoletaHRCxCuentaConexion(rsreporte.Fields!idCuentaAtencion, oConexion)
                       If rsReporte1.RecordCount > 0 Then
                          lcMedicoEgreso = Trim(rsReporte1.Fields!ApellidoPaterno) & " " & Trim(rsReporte1.Fields!ApellidoMaterno) & " " & rsReporte1.Fields!Nombres
                          lcServicio = IIf(IsNull(rsReporte1.Fields!nombre), "", rsReporte1.Fields!nombre)
                          ldFIngreso = rsReporte1.Fields!FechaIngreso
                          ldFAlta = rsReporte1.Fields!fechaEgreso
                          lnEdad = rsReporte1.Fields!Edad
                          lcCama = IIf(IsNull(rsReporte1.Fields!Codigo), "", rsReporte1.Fields!Codigo)
                       End If
                       rsReporte1.Close
                    End If
                    PosIniX = 0.5
                    PosIniY = 0.6
                    lnCmY = 0.4
                    'Printer.PaperSize = "soloCtaServicios"
                    'Printer.PaperSize = vbPRPSUser    '256
                    Printer.Font = "Tahoma"
                    Printer.FontSize = 8
                    Printer.ScaleMode = vbCentimeters
                    'Printer.Height = 13.9
                    lcCajero$ = Chr(15) + String(35, " ") & IIf(rsreporte.Fields!IdTipoPago = 2, "DEVOLUCION", IIf(rsreporte.Fields!TotalBoleta = 0, "EXONERADO", ".."))
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(35, " ") & IIf(rsreporte.Fields!idEstadoComprobante = 9, "A N U L A D O ", "..")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(15, " ") & nroSerie & " - " & nroDcto
                    If IsNull(rsreporte.Fields!idCuentaAtencion) Then
                       lcCajero$ = lcCajero$ & ""
                    Else
                       lcCajero$ = lcCajero$ & String(5, " ") & " (Cta: " & Trim(Str(rsreporte.Fields!idCuentaAtencion)) & ")"
                    End If
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(5, " ") & Left(rsreporte.Fields!Observaciones + String(40, " "), 40)
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(15, " ") & Left(Trim(IIf(IsNull(rsreporte.Fields!razonSocial), "..", rsreporte.Fields!razonSocial)), 43) & String(51, " ") & lnEdad
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(15, " ") & IIf(IsNull(rsreporte.Fields!NroHistoriaClinica), "(Particular)", "(HC: " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsreporte.Fields!NroHistoriaClinica)), False) & ")") & String(73, " ") & lcCama
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    'Medico y Servicio Alta
                    lcCajero$ = String(24, " ") & Left(lcMedicoEgreso & String(30, " "), 30) & String(40, " ") & Left(lcServicio & String(31, " "), 30)
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    'Fecha Ingreso y Alta
                    lcCajero$ = String(28, " ") & ldFIngreso & String(64, " ") & ldFAlta
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(45, " ")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(45, " ")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    
                    NúmeroRegistro = 1
                    Do While Not rsreporte.EOF
                       lnSubTotal = rsreporte.Fields!TotalPorPagar
                       lcCantidad = Format(rsreporte.Fields!Cantidad, "#####")
                       lcPrecioUnitario = Format(rsreporte.Fields!PrecioUnitario, "#####.###0")
                       lcSubTotal = Format(lnSubTotal, "######.#0")
                       lcProducto = Left(rsreporte.Fields!NombreProducto, 50)
                       lcCajero$ = String$(1, " ") & _
                                   Left(rsreporte.Fields!Codigo, 5) + String(6 - Len(Trim(Left(rsreporte.Fields!Codigo, 5))), " ") & _
                                   "  " + Left(lcProducto, 50) + String(50 - Len(lcProducto), " ")
                       Printer.CurrentY = PosIniY
                       Printer.CurrentX = PosIniX
                       Printer.Print lcCajero$
                       'cantidad
                       lcCajero$ = lcCantidad
                       Printer.CurrentY = PosIniY
                       Printer.CurrentX = PosIniX + 7.5 - Printer.TextWidth(lcCajero$)
                       Printer.Print lcCajero$
                       'precio
                       lcCajero$ = lcPrecioUnitario
                       Printer.CurrentY = PosIniY
                       Printer.CurrentX = PosIniX + 9.5 - Printer.TextWidth(lcCajero$)
                       Printer.Print lcCajero$
                       'Total
                       lcCajero$ = lcSubTotal
                       Printer.CurrentY = PosIniY
                       Printer.CurrentX = PosIniX + 11.5 - Printer.TextWidth(lcCajero$)
                       Printer.Print lcCajero$
                       '
                       PosIniY = PosIniY + lnCmY
                       NúmeroRegistro = NúmeroRegistro + 1
                       rsreporte.MoveNext
                    Loop
                    
                    If NúmeroRegistro < (lnNumeroMaxItemBoleta + 1) Then
                       Do While NúmeroRegistro < (lnNumeroMaxItemBoleta + 1)
                          'columna 1,fila X  (vacia)
                          lcCajero$ = String(45, " ") '& Chr(13) & Chr(10)
                          Printer.CurrentY = PosIniY
                          Printer.CurrentX = PosIniX
                          Printer.Print lcCajero$
                          PosIniY = PosIniY + lnCmY
                          NúmeroRegistro = NúmeroRegistro + 1
                       Loop
                    End If
                    '
                    lcCajero$ = String(18, " ") & Format(ldFechaBoleta, sighEntidades.DevuelveFechaSoloFormato_DMY) & String(95, " ") & Format(lnTotal + lnImporteEXO + lnAdelantos, "######.#0")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(18, " ") & Format(ldFechaBoleta, sighEntidades.DevuelveHoraSoloFormato_HM) & String(107, " ") & Format(lnImporteEXO, "######.#0")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(18, " ") & lcCajero1 & String(105, " ") & Format(lnTotal, "######.#0")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcProducto = Left("Son: " & sighEntidades.Numlet(sighEntidades.DevuelveNumeroSinDecimales(lnTotal)) & " con " & sighEntidades.DevuelveSoloDecimales(lnTotal) & "/100  Soles", 60)
                    lcCajero$ = String(3, " ") & Left(lcProducto, 60) + String(60 - Len(lcProducto), " ")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    '
                    Close #1
                    rsreporte.Close
                Else
                    '*****************HRC boleta RESTO DE OPCIONES DE CAJA*******************
                    PosIniX = 0.5
                    PosIniY = 0.6
                    lnCmY = 0.4
                    'Printer.PaperSize = "soloServicios"
                    Printer.Font = "Arial"
                    Printer.FontSize = 8
                    Printer.ScaleMode = vbCentimeters
                    lcCajero$ = Chr(15) + String(5, " ") & IIf(rsreporte.Fields!IdTipoPago = 2, "DEVOLUCION", IIf(rsreporte.Fields!TotalBoleta = 0, "EXONERADO", ".."))
                    lcCajero$ = lcCajero$ & String(5, " ") & IIf(rsreporte.Fields!idEstadoComprobante = 9, "A N U L A D O ", "..")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    If IsNull(rsreporte.Fields!idCuentaAtencion) Then
                       lcCajero$ = ""
                    Else
                       lcCajero$ = " (Cta: " & Trim(Str(rsreporte.Fields!idCuentaAtencion)) & ")"
                    End If
                    lcCajero$ = lcCajero$ & String(5, " ") & Left(rsreporte.Fields!Observaciones + String(40, " "), 40)
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(15, " ") & Left(Trim(IIf(IsNull(rsreporte.Fields!razonSocial), "..", rsreporte.Fields!razonSocial)) + IIf(IsNull(rsreporte.Fields!NroHistoriaClinica), "(Particular)", "(HC: " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsreporte.Fields!NroHistoriaClinica)), False) & ")") + String(43, " "), 43) & String(5, " ") & nroSerie & " - " & nroDcto
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '(vacia)
                    lcCajero$ = String(45, " ")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '(vacia)
                    lcCajero$ = String(45, " ")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    NúmeroRegistro = 1
                    Do While Not rsreporte.EOF
                       lnSubTotal = rsreporte.Fields!TotalPorPagar
                       lcCantidad = Format(rsreporte.Fields!Cantidad, "#####")
                       lcPrecioUnitario = Format(rsreporte.Fields!PrecioUnitario, "#####.###0")
                       lcSubTotal = Format(lnSubTotal, "######.#0")
                       lcProducto = Left(rsreporte.Fields!NombreProducto, 40)
                       lcCajero$ = String$(1, " ") & _
                                   Left(rsreporte.Fields!Codigo, 5) + String(5 - Len(Trim(Left(rsreporte.Fields!Codigo, 5))), " ") & _
                                   "  " + Left(lcProducto, 40) + String(40 - Len(lcProducto), " ")
                       Printer.CurrentY = PosIniY
                       Printer.CurrentX = PosIniX
                       Printer.Print lcCajero$
                       'Cantidad
                       lcCajero$ = lcCantidad
                       Printer.CurrentY = PosIniY
                       Printer.CurrentX = PosIniX + 7.5 - Printer.TextWidth(lcCajero$)
                       Printer.Print lcCajero$
'                       'Precio
'                       lcCajero$ = lcPrecioUnitario
'                       Printer.CurrentY = PosIniY
'                       Printer.CurrentX = PosIniX + 7.5 - Printer.TextWidth(lcCajero$)
'                       Printer.Print lcCajero$
                       'Total
                       lcCajero$ = lcSubTotal
                       Printer.CurrentY = PosIniY
                       Printer.CurrentX = PosIniX + 9.3 - Printer.TextWidth(lcCajero$)
                       Printer.Print lcCajero$
                       '
                       PosIniY = PosIniY + lnCmY
                       NúmeroRegistro = NúmeroRegistro + 1
                       rsreporte.MoveNext
                    Loop
                    '
                    If NúmeroRegistro < (lnNumeroMaxItemBoleta + 1) Then
                       Do While NúmeroRegistro < (lnNumeroMaxItemBoleta + 1)
                          '(vacia)
                          lcCajero$ = String(45, " ")
                          Printer.CurrentY = PosIniY
                          Printer.CurrentX = PosIniX
                          Printer.Print lcCajero$
                          PosIniY = PosIniY + lnCmY
                          NúmeroRegistro = NúmeroRegistro + 1
                       Loop
                    End If
                    '
                    lcCajero$ = String(18, " ") & Format(ldFechaBoleta, sighEntidades.DevuelveFechaSoloFormato_DMY) & String(67, " ") & Format(lnTotal + lnImporteEXO + lnAdelantos, "######.#0")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(18, " ") & Format(ldFechaBoleta, sighEntidades.DevuelveHoraSoloFormato_HM) & String(77, " ") & Format(lnImporteEXO, "######.#0")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcProducto = Left("Son: " & sighEntidades.Numlet(sighEntidades.DevuelveNumeroSinDecimales(lnTotal)) & " con " & sighEntidades.DevuelveSoloDecimales(lnTotal) & "/100   Soles", 60)
                    lcCajero$ = String(3, " ") & Left(lcProducto, 60) + String(60 - Len(lcProducto), " ") & String(17, " ") & Format(lnTotal, "######.#0")
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    PosIniY = PosIniY + lnCmY
                    '
                    lcCajero$ = String(18, " ") & lcCajero1
                    Printer.CurrentY = PosIniY
                    Printer.CurrentX = PosIniX
                    Printer.Print lcCajero$
                    '
                    Close #1
                    rsreporte.Close
                End If
                Printer.EndDoc
            End Select
            
        End If
        oConexion.Close
        Set oConexion = Nothing
        Set rsreporte = Nothing
        Set rsReporte1 = Nothing
        Set rsRsBoletaDetallada = Nothing
        Set mo_ReporteUtil = Nothing
        Set oReglasCaja = Nothing
        Exit Sub
ManejadorError:
        Resume Next
        If Err.Number = 55 Then
           Close #1
           Resume
        Else
           If Err.Number = 53 Then
              Resume Next
           ElseIf Err.Number = 380 Then
              Resume Next
           ElseIf Err.Number = 70 Then
              Resume Next
           Else
              MsgBox Err.Description
              Resume
           End If
        End If
End Sub



Sub GenerarRecordsetTemporal1()
    With rsRsTmp
          .Fields.Append "FechaCobranza", adDouble
          .Fields.Append "NroSerie", adVarChar, 5, adFldIsNullable
          .Fields.Append "NroDocumento", adVarChar, 20, adFldIsNullable
          .Fields.Append "NroHistoriaClinica", adVarChar, 50, adFldIsNullable
          .Fields.Append "RazonSocial", adVarChar, 100, adFldIsNullable
          .Fields.Append "adelantos", adDouble
          .Fields.Append "ImporteEXO", adDouble
          .Fields.Append "idEstadoComprobante", adUnsignedBigInt
          .Fields.Append "totalPorPagar", adDouble
          .Fields.Append "IdProducto", adUnsignedBigInt
          .Fields.Append "totalPagado", adDouble
          .Fields.Append "redondeo", adDouble
          .Fields.Append "cajero", adVarChar, 20, adFldIsNullable
          .Fields.Append "Devoluciones", adDouble
          .LockType = adLockOptimistic
          .Open
    End With
End Sub

Sub ImpresionFUAformato(FechaIngreso As String, HoraIngreso As String, Paciente As String, NroHistoriaClinica As Long, _
                       Servicio As String, Medico As String, lcTipoServicio As String, lnIdAtencion As Long, _
                       ml_idOrden As String, ml_idCuentaAtencion As String, lcFormaPago As String, lcNroCola As String, _
                       lnIdUsuario As Long, lcServicioDelTarifario As String, lcFichaFamiliar As String, _
                       lnIdTipoNumeracion As Long, lcParametro216 As String, lcParametro306 As String, _
                       lcPaterno As String, lcMaterno As String, lcPrimerNombre As String, _
                       lcSegundoNombre As String, lnIdTipoSexo As Long, ldFechaNacimiento As Date, _
                       ldDNI As String, Optional lnIdMedico As Long)
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Dim rsreporte As New Recordset
        Dim mrs_Tmp As New Recordset
        Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
        Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
        Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
        Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
        Dim lcNombreHospital As String, lnIdprograma As Long, lnIdTurno As Long, lcTurno As String
        Dim lcProfesionMedico As String
        Dim lcRuc As String, lcDireccion As String, lcTlefono As String
        Dim lcDevuelveFechaServidor As New SIGHDatos.Parametros
        Dim lcImpresora As String, lcImpresoraActual As String
        
 
            lcNombreHospital = lcBuscaParametro.SeleccionaFilaParametro(205)
            lcRuc = lcBuscaParametro.SeleccionaFilaParametro(339)
            lcDireccion = lcBuscaParametro.SeleccionaFilaParametro(206)
            lcTlefono = lcBuscaParametro.SeleccionaFilaParametro(207)
            lcImpresora = Trim(lcBuscaParametro.SeleccionaFilaParametro(336))
            '
            With mrs_Tmp
                .Fields.Append "FechaIngreso", adVarChar, 100, adFldIsNullable
                .Fields.Append "Turno", adVarChar, 100, adFldIsNullable
                .Fields.Append "Paciente", adVarChar, 100, adFldIsNullable
                .Fields.Append "Usuario", adVarChar, 100, adFldIsNullable
                .Fields.Append "NroHistoriaClinica", adVarChar, 100, adFldIsNullable
                .Fields.Append "Servicio", adVarChar, 100, adFldIsNullable
                .Fields.Append "Medico", adVarChar, 100, adFldIsNullable
                .Fields.Append "Especialidad", adVarChar, 100, adFldIsNullable
                .Fields.Append "Interconsulta", adVarChar, 100, adFldIsNullable
                .Fields.Append "ColaTipoS", adVarChar, 100, adFldIsNullable
                .Fields.Append "FechaHoraImp", adVarChar, 100, adFldIsNullable
                .LockType = adLockOptimistic
                .Open
            End With
            '
            lcProfesionMedico = ""
            If lnIdMedico > 0 Then
               lcProfesionMedico = mo_ReglasComunes.TiposEmpleadosSeleccionarIdMedico(lnIdMedico)
               If lcProfesionMedico <> "" Then
                  lcProfesionMedico = lcProfesionMedico
               End If
            End If
           
            '
            Dim lcMedicoDNI As String, lcMedicoColeg As String, lcMedicoRNE As String
            lcMedicoDNI = ""
            lcMedicoColeg = ""
            lcMedicoRNE = ""
            Set rsreporte = mo_ReglasDeProgMedica.MedicosSeleccionarXIdMedico(lnIdMedico)
            If rsreporte.RecordCount > 0 Then
               lcMedicoDNI = IIf(IsNull(rsreporte!DNI), "", Trim(rsreporte!DNI))
               lcMedicoColeg = IIf(IsNull(rsreporte!Colegiatura), "", Trim(rsreporte!Colegiatura))
               lcMedicoRNE = IIf(IsNull(rsreporte!rne), "", rsreporte!rne)
            End If
            rsreporte.Close
            '
            Dim lnIdServicio As Long
            lcTurno = ""
            lnIdServicio = 0
            Set rsreporte = mo_ReglasDeProgMedica.CitasSeleccionarXfechaMedico(FechaIngreso, lnIdMedico)
            rsreporte.Filter = "horaInicio='" & HoraIngreso & "'"
            If rsreporte.RecordCount > 0 Then
               lnIdServicio = rsreporte!IdServicio
               lnIdprograma = rsreporte!IdProgramacion
               rsreporte.Close
               Set rsreporte = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarXidentificador(lnIdprograma)
               If rsreporte.RecordCount > 0 Then
                  lnIdTurno = rsreporte.Fields!IdTurno
                  rsreporte.Close
                  Set rsreporte = mo_ReglasDeProgMedica.TurnosSeleccionarPorIdentificador(lnIdTurno)
                  If rsreporte.RecordCount > 0 Then
                     lcTurno = " (" & Trim(rsreporte!descripcion) & ")"
                  End If
               End If
            End If
            rsreporte.Close
            '
            Dim lcUPS As String
            lcUPS = ""
            Set rsreporte = mo_ReglasComunes.ServiciosSeleccionarXidentificador(lnIdServicio)
            If rsreporte.RecordCount > 0 Then
               lcUPS = rsreporte!codigoServicioFUA
            End If
            rsreporte.Close
            
            mrs_Tmp.AddNew
            If lcBuscaParametro.SeleccionaFilaParametro(281) = "x" Then  'solo el HRA
            '**** Programa: se agrego campo Turno que ya no esta concatenado con la fecha  (sale en la siguiente linea)
            '**** Programado por:Eder Yamill Palomino Espinoza
            '**** Fecha: 06102014
               
               mrs_Tmp.Fields!FechaIngreso = "Fecha: " & FechaIngreso & IIf(lcParametro306 = "S", "", lcTurno)
              
            Else
               mrs_Tmp.Fields!FechaIngreso = "Fecha: " & FechaIngreso & " Hr: " & HoraIngreso '& lcTurno
            End If
            mrs_Tmp.Fields!Paciente = "Paciente: " & Paciente
            mrs_Tmp.Fields!Usuario = "Usuario: " & mo_ReglasCaja.SeleccionaDatosCajero(lnIdUsuario, sghUsuario)
            'debb-09/09/2015 (inicio)
            If lcParametro306 = "S" Then
               mrs_Tmp.Fields!FechaHoraImp = "Fecha y Hora de Impresión: " & _
                                             lcDevuelveFechaServidor.RetornaFechaServidorSQL & " " & _
                                             lcDevuelveFechaServidor.RetornaHoraServidorSQLserverFormatoGalenhos
            Else
               mrs_Tmp.Fields!FechaHoraImp = "Fecha Impres: " & lcDevuelveFechaServidor.RetornaFechaServidorSQL & " " & _
                                             lcDevuelveFechaServidor.RetornaHoraServidorSQLserverFormatoGalenhos
            End If
            'debb-09/09/2015 (fin)
'            mrs_Tmp.Fields!FechaHoraImp = "Fecha y Hora de Impresión: " & lcDevuelveFechaServidor.RetornaFechaServidorSQL & " " & lcDevuelveFechaServidor.RetornaHoraServidorSQLserverFormatoGalenhos
            If lcBuscaParametro.SeleccionaFilaParametro(277) = "S" And Trim(lcFichaFamiliar) <> "" Then
               mrs_Tmp.Fields!NroHistoriaClinica = "Ficha Familiar: " & Trim(lcFichaFamiliar)
            Else
                If lcParametro306 = "S" Then
                    mrs_Tmp.Fields!NroHistoriaClinica = "N°Historia: " & _
                                   HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(NroHistoriaClinica)), False) & _
                                   " " & mo_ReglasArchivoClinico.HistoriaClinicaEsNueva(NroHistoriaClinica, lnIdTipoNumeracion)
                Else
                    mrs_Tmp.Fields!NroHistoriaClinica = "N°Historia: " & _
                                   HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(NroHistoriaClinica)), False) & _
                                   " " & mo_ReglasArchivoClinico.HistoriaClinicaEsNueva(NroHistoriaClinica, lnIdTipoNumeracion) & "     No Cuenta: " & ml_idCuentaAtencion
                End If
            End If
            mrs_Tmp.Fields!Servicio = "Serv: " & Servicio
            '**** Programa: se agrego campo Especialidad que ya no esta concatenado con medico (sale en la siguiente linea)
            '**** Programado por:Eder Yamill Palomino Espinoza
            '**** Fecha: 06102014
            'mrs_Tmp.Fields!Medico = Left("Médico: " & Trim(Medico) & lcProfesionMedico, 100)
            mrs_Tmp.Fields!Medico = lcProfesionMedico & " - " & Trim(Medico) '& lcProfesionMedico
            If lcNroCola = "" Then
                If IsNumeric(ml_idOrden) Then
                    mrs_Tmp.Fields!Interconsulta = "Ord. Pago: " & CStr(ml_idOrden)
                Else
                    mrs_Tmp.Fields!Interconsulta = ""
                End If
                mrs_Tmp.Fields!ColaTipoS = lcTipoServicio
            Else
                If lcParametro306 = "S" Then
                    mrs_Tmp.Fields!Interconsulta = "Interconsulta: Si (   )             No (  )"
                Else
                    mrs_Tmp.Fields!Interconsulta = "Interconsulta:           Si (   )             No (  )"
                End If
                mrs_Tmp.Fields!ColaTipoS = "Cupo:     " & lcNroCola
            End If
            mrs_Tmp.Update
            'Reporte
            On Error Resume Next
            '
            If Not (lcImpresora = "" Or lcImpresora = "0") Then
               lcImpresoraActual = Printer.DeviceName
               sighEntidades.ImpresoraPredeterminada lcImpresora
            End If
            '
            '***************** Ticket a 2 copias en 1 pagina (Impresora: EPSON)**********************
            Set FUA.DataSource = mrs_Tmp
            FUA.Sections("Cabecera").Controls("etDni").Caption = ldDNI
            FUA.Sections("Cabecera").Controls("etPaterno").Caption = UCase(lcPaterno)
            FUA.Sections("Cabecera").Controls("etMaterno").Caption = UCase(lcMaterno)
            FUA.Sections("Cabecera").Controls("etPrimerNombre").Caption = UCase(lcPrimerNombre)
            FUA.Sections("Cabecera").Controls("etSegundoNombre").Caption = UCase(lcSegundoNombre)
            FUA.Sections("Cabecera").Controls("etSexoM").Caption = IIf(lnIdTipoSexo = 1, "X", "")
            FUA.Sections("Cabecera").Controls("etSexoF").Caption = IIf(lnIdTipoSexo = 1, "", "X")
            FUA.Sections("Cabecera").Controls("edNacDia").Caption = Trim(Str(Day(ldFechaNacimiento)))
            FUA.Sections("Cabecera").Controls("edNacMes").Caption = Trim(Str(Month(ldFechaNacimiento)))
            FUA.Sections("Cabecera").Controls("etNacAnio").Caption = Trim(Str(Year(ldFechaNacimiento)))
            FUA.Sections("Cabecera").Controls("etNroHistoria").Caption = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(NroHistoriaClinica)), False)
            FUA.Sections("Cabecera").Controls("etAtenDia").Caption = Left(FechaIngreso, 2)
            FUA.Sections("Cabecera").Controls("etAtenMes").Caption = Mid(FechaIngreso, 4, 2)
            FUA.Sections("Cabecera").Controls("etAtenAnio").Caption = Right(FechaIngreso, 4)
            FUA.Sections("Cabecera").Controls("etUPS").Caption = lcUPS
            FUA.Sections("Cabecera").Controls("etMedicoDNI").Caption = lcMedicoDNI
            FUA.Sections("Cabecera").Controls("etMedico").Caption = Medico
            FUA.Sections("Cabecera").Controls("etMedicoColegiatura").Caption = lcMedicoColeg
            FUA.Sections("Cabecera").Controls("etRNE").Caption = lcMedicoRNE
            If sighEntidades.Parametro301valorInt <> "1" Then
               Set FUA.Sections("Cabecera").Controls("Image1").Picture = LoadPicture("")
            End If
            FUA.RightMargin = 0  'getReportRightMargin
            FUA.TopMargin = 0 'getReportTopMargin
            FUA.LeftMargin = 0 'getReportLeftMargin
            FUA.BottomMargin = 0 ' getReportBottomMargin
            FUA.Orientation = rptOrientPortrait
            'If lcParametro216 <> "1" Then
               FUA.Show 1
           ' Else
           '    FUA.PrintReport
           ' End If
            '
            If Not (lcImpresora = "" Or lcImpresora = "0") Then
               sighEntidades.ImpresoraPredeterminada lcImpresoraActual
            End If
            '
        'End If
        'Set rsReporte = Nothing
        Set mrs_Tmp = Nothing
        Set mo_ReglasCaja = Nothing
        Set mo_ReglasComunes = Nothing
        Set mo_ReglasArchivoClinico = Nothing
        Set mo_ReglasDeProgMedica = Nothing
End Sub

'debb-13/05/2016
'***************daniel barrantes**************
'***************Imprime la PreCUENTA
'***************
Sub ImpresionPreCuenta(FechaIngreso As String, HoraIngreso As String, Paciente As String, NroHistoriaClinica As Long, _
                       Servicio As String, Medico As String, lcTipoServicio As String, lnIdAtencion As Long, _
                       ml_idOrden As String, ml_idCuentaAtencion As String, lcFormaPago As String, lcNroCola As String, _
                       lnIdUsuario As Long, lcServicioDelTarifario As String, lcFichaFamiliar As String, _
                       lnIdTipoNumeracion As Long, lcParametro216 As String, lcParametro306 As String, _
                       lbImprimePDF As Boolean, Optional lnIdMedico As Long)
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Dim rsreporte As New Recordset
        Dim mrs_Tmp As New Recordset
        Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
        Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
        Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
        Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
        Dim lcNombreHospital As String, lnIdprograma As Long, lnIdTurno As Long, lcTurno As String
        Dim lcProfesionMedico As String
        Dim lcRuc As String, lcDireccion As String, lcTlefono As String
        Dim lcDevuelveFechaServidor As New SIGHDatos.Parametros
        Dim lcImpresora As String, lcImpresoraActual As String
        
        '******** Crystal Report
'        Dim oRptClaseCry As New rCrystal
'        If lcBuscaParametro.SeleccionaFilaParametro(216) <> "1" Then
'           oRptClaseCry.DestinoReporte = sghPantalla
'        Else
'           oRptClaseCry.DestinoReporte = sghImpresora
'        End If
'        oRptClaseCry.TipoReporte = "ImpresionPreCuenta"
'        oRptClaseCry.lcTipoServicio = lcTipoServicio
'        oRptClaseCry.LcIdCuentaAtencion = ml_idCuentaAtencion
'        oRptClaseCry.lcFormaPago = "IAFA: " & lcFormaPago
'        oRptClaseCry.lcNroCola = lcNroCola
'        oRptClaseCry.idUsuario = lnIdUsuario
'        oRptClaseCry.lcServicioDelTarifario = Mid(lcServicioDelTarifario, InStr(lcServicioDelTarifario, "=") + 1, 100)
'        oRptClaseCry.idAtencion = lnIdAtencion
'        oRptClaseCry.Show vbModal
'        Set oRptClaseCry = Nothing
        '******** dataReport
        'Set rsReporte = mo_AdminReportes.ReporteAtencionesParaHistoriaClinica(lnIdAtencion)
        'If rsReporte.RecordCount = 0 Then
        '    MsgBox "No existen datos", vbInformation, "Reporte"
        'Else
            lcNombreHospital = lcBuscaParametro.SeleccionaFilaParametro(205)
            lcRuc = lcBuscaParametro.SeleccionaFilaParametro(339)
            lcDireccion = lcBuscaParametro.SeleccionaFilaParametro(206)
            lcTlefono = lcBuscaParametro.SeleccionaFilaParametro(207)
            lcImpresora = Trim(lcBuscaParametro.SeleccionaFilaParametro(336))
            '
            With mrs_Tmp
                .Fields.Append "FechaIngreso", adVarChar, 100, adFldIsNullable
                .Fields.Append "Turno", adVarChar, 100, adFldIsNullable
                .Fields.Append "Paciente", adVarChar, 100, adFldIsNullable
                .Fields.Append "Usuario", adVarChar, 100, adFldIsNullable
                .Fields.Append "NroHistoriaClinica", adVarChar, 100, adFldIsNullable
                .Fields.Append "Servicio", adVarChar, 100, adFldIsNullable
                .Fields.Append "Medico", adVarChar, 100, adFldIsNullable
                .Fields.Append "Especialidad", adVarChar, 100, adFldIsNullable
                .Fields.Append "Interconsulta", adVarChar, 100, adFldIsNullable
                .Fields.Append "ColaTipoS", adVarChar, 100, adFldIsNullable
                .Fields.Append "FechaHoraImp", adVarChar, 100, adFldIsNullable
                .LockType = adLockOptimistic
                .Open
            End With
            '
            lcProfesionMedico = ""
            If lnIdMedico > 0 Then
               lcProfesionMedico = mo_ReglasComunes.TiposEmpleadosSeleccionarIdMedico(lnIdMedico)
               If lcProfesionMedico <> "" Then
                  lcProfesionMedico = lcProfesionMedico
               End If
            End If
            '
            lcTurno = ""
            Set rsreporte = mo_ReglasDeProgMedica.CitasSeleccionarXfechaMedico(FechaIngreso, lnIdMedico)
            rsreporte.Filter = "horaInicio='" & HoraIngreso & "'"
            If rsreporte.RecordCount > 0 Then
               lnIdprograma = rsreporte!IdProgramacion
               rsreporte.Close
               Set rsreporte = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarXidentificador(lnIdprograma)
               If rsreporte.RecordCount > 0 Then
                  lnIdTurno = rsreporte.Fields!IdTurno
                  rsreporte.Close
                  Set rsreporte = mo_ReglasDeProgMedica.TurnosSeleccionarPorIdentificador(lnIdTurno)
                  If rsreporte.RecordCount > 0 Then
                     lcTurno = " (" & Trim(rsreporte!descripcion) & ")"
                  End If
               End If
            End If
            rsreporte.Close
            '
            mrs_Tmp.AddNew
            If lcBuscaParametro.SeleccionaFilaParametro(281) = "x" Then  'solo el HRA
            '**** Programa: se agrego campo Turno que ya no esta concatenado con la fecha  (sale en la siguiente linea)
            '**** Programado por:Eder Yamill Palomino Espinoza
            '**** Fecha: 06102014
               
               mrs_Tmp.Fields!FechaIngreso = "Fecha: " & FechaIngreso & IIf(lcParametro306 = "S", "", lcTurno)
              
            Else
               mrs_Tmp.Fields!FechaIngreso = "Fecha: " & FechaIngreso & " Hr: " & HoraIngreso '& lcTurno
            End If
            mrs_Tmp.Fields!Paciente = "Paciente: " & Paciente
            mrs_Tmp.Fields!Usuario = "Usuario: " & mo_ReglasCaja.SeleccionaDatosCajero(lnIdUsuario, sghUsuario)
            'debb-09/09/2015 (inicio)
            If lcParametro306 = "S" Then
               mrs_Tmp.Fields!FechaHoraImp = "Fecha y Hora de Impresión: " & _
                                             lcDevuelveFechaServidor.RetornaFechaServidorSQL & " " & _
                                             lcDevuelveFechaServidor.RetornaHoraServidorSQLserverFormatoGalenhos
            Else
               mrs_Tmp.Fields!FechaHoraImp = "Fecha Impres: " & lcDevuelveFechaServidor.RetornaFechaServidorSQL & " " & _
                                             lcDevuelveFechaServidor.RetornaHoraServidorSQLserverFormatoGalenhos
            End If
            'debb-09/09/2015 (fin)
'            mrs_Tmp.Fields!FechaHoraImp = "Fecha y Hora de Impresión: " & lcDevuelveFechaServidor.RetornaFechaServidorSQL & " " & lcDevuelveFechaServidor.RetornaHoraServidorSQLserverFormatoGalenhos
            If lcBuscaParametro.SeleccionaFilaParametro(277) = "S" And Trim(lcFichaFamiliar) <> "" Then
               mrs_Tmp.Fields!NroHistoriaClinica = "Ficha Familiar: " & Trim(lcFichaFamiliar)
            Else
                If lcParametro306 = "S" Then
                    mrs_Tmp.Fields!NroHistoriaClinica = "N°Historia: " & _
                                   HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(NroHistoriaClinica)), False) & _
                                   " " & mo_ReglasArchivoClinico.HistoriaClinicaEsNueva(NroHistoriaClinica, lnIdTipoNumeracion)
                Else
                    mrs_Tmp.Fields!NroHistoriaClinica = "N°Historia: " & _
                                   HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(NroHistoriaClinica)), False) & _
                                   " " & mo_ReglasArchivoClinico.HistoriaClinicaEsNueva(NroHistoriaClinica, lnIdTipoNumeracion) & "     No Cuenta: " & ml_idCuentaAtencion
                End If
            End If
            mrs_Tmp.Fields!Servicio = "Serv: " & Servicio
            '**** Programa: se agrego campo Especialidad que ya no esta concatenado con medico (sale en la siguiente linea)
            '**** Programado por:Eder Yamill Palomino Espinoza
            '**** Fecha: 06102014
            'mrs_Tmp.Fields!Medico = Left("Médico: " & Trim(Medico) & lcProfesionMedico, 100)
            mrs_Tmp.Fields!Medico = lcProfesionMedico & " - " & Trim(Medico) '& lcProfesionMedico
            If lcNroCola = "" Then
                If IsNumeric(ml_idOrden) Then
                    mrs_Tmp.Fields!Interconsulta = "Ord. Pago: " & CStr(ml_idOrden)
                Else
                    mrs_Tmp.Fields!Interconsulta = ""
                End If
                mrs_Tmp.Fields!ColaTipoS = lcTipoServicio
            Else
                If lcParametro306 = "S" Then
                    mrs_Tmp.Fields!Interconsulta = "Interconsulta: Si (   )             No (  )"
                Else
                    mrs_Tmp.Fields!Interconsulta = "Interconsulta:           Si (   )             No (  )"
                End If
                mrs_Tmp.Fields!ColaTipoS = "Cupo:     " & lcNroCola
            End If
            mrs_Tmp.Update
            'Reporte
            On Error Resume Next
            '
            If Not (lcImpresora = "" Or lcImpresora = "0") Then
               lcImpresoraActual = Printer.DeviceName
               sighEntidades.ImpresoraPredeterminada lcImpresora
            End If
            '
            Dim lbSePuedeImprimirPDF As Boolean, lcArchivoPDF As String
            If lbImprimePDF = True Then
               lcArchivoPDF = App.Path & "\" & Trim(Str(ml_idCuentaAtencion)) & ".pdf"
               If SePuedeImprimirPDF(lcArchivoPDF, False) = True Then
                  lbSePuedeImprimirPDF = True
               End If
            End If
            '
            If lcParametro306 = "S" Then
               '***************** Ticket a 1 copias en 1 pagina (Impresora: ticketera)**********************
                '**** Programa: se agrego Turno y Especialidad
                '**** Programado por:Eder Yamill Palomino Espinoza
                '**** Fecha: 06102014
                Set CePreCta.DataSource = mrs_Tmp
                CePreCta.Sections("Cabecera").Controls("lcHospital1").Caption = UCase(lcNombreHospital)
                CePreCta.Sections("Cabecera").Controls("DatosEESS").Caption = UCase("RUC: " & lcRuc) & vbCrLf & UCase(lcDireccion) & vbCrLf & "Telef: " & UCase(lcTlefono)
                CePreCta.Sections("Cabecera").Controls("NombreTicket").Caption = UCase("TICKET DE CITA")
'                CePreCta.Sections("Detalle").Controls("Turno").Caption = "Turno: " & lcTurno
                If lcTurno <> "" Then lcTurno = "Turno: " & lcTurno
                CePreCta.Sections("Detalle").Controls("Turno").Caption = lcTurno
                CePreCta.Sections("Detalle").Controls("lcNroHNroC").Caption = mrs_Tmp.Fields!NroHistoriaClinica
                CePreCta.Sections("Detalle").Controls("lcNroCuenta").Caption = "N°Cuenta: " & ml_idCuentaAtencion
                CePreCta.Sections("Detalle").Controls("Especialidad").Caption = "Especialidad: " & lcProfesionMedico

                CePreCta.Sections("Detalle").Controls("lcFormaPago").Caption = lcFormaPago
                CePreCta.Sections("Detalle").Controls("lcServicioDelTarifario").Caption = lcServicioDelTarifario

                CePreCta.Sections("PieInforme").Controls("lcUsuario").Caption = "Fec: " & Format(lcDevuelveFechaServidor.RetornaFechaHoraServidorSQL, sighEntidades.DevuelveFechaSoloFormato_DMY_HM) & "  " & "Us: " & mo_ReglasCaja.SeleccionaDatosCajero(lnIdUsuario, sghUsuario)
                CePreCta.Sections("PieInforme").Controls("Terminal").Caption = "Terminal:  " & lcBuscaParametro.RetornaNombreDeServidor
                CePreCta.Sections("PieInforme").Controls("TextoVacio").Caption = lcBuscaParametro.SeleccionaFilaParametro(346)
                CePreCta.Sections("PieInforme").Controls("FechahoraImp").Caption = "Fecha y Hora de Impresión: " & lcDevuelveFechaServidor.RetornaFechaServidorSQL & " " & lcDevuelveFechaServidor.RetornaHoraServidorSQLserverFormatoGalenhos
                
                CePreCta.RightMargin = 100
                CePreCta.TopMargin = 100
                CePreCta.LeftMargin = 100
                CePreCta.BottomMargin = 100
                
                CePreCta.Orientation = rptOrientPortrait
                If lbSePuedeImprimirPDF = True Then
                   CePreCta.PrintReport False
                   
                Else
                    If lcParametro216 <> "1" Then
                       CePreCta.Show 1
                    Else
                       CePreCta.PrintReport
                    End If
                End If
                SeteaOtraImpresoraDefault sighEntidades.ImpresoraDefaultDeEstaPC
            Else
                '***************** Ticket a 2 copias en 1 pagina (Impresora: EPSON)**********************
                Set CePreCuenta.DataSource = mrs_Tmp
                CePreCuenta.Sections("Detalle").Controls("lcHospital1").Caption = UCase(lcNombreHospital)
                CePreCuenta.Sections("Detalle").Controls("lcHospital2").Caption = UCase(lcNombreHospital)
                CePreCuenta.Sections("Detalle").Controls("lcFormaPago").Caption = lcFormaPago
                CePreCuenta.Sections("Detalle").Controls("lcFormaPago1").Caption = lcFormaPago
                CePreCuenta.Sections("Detalle").Controls("lcServicioDelTarifario").Caption = lcServicioDelTarifario
                CePreCuenta.Sections("Detalle").Controls("lcServicioDelTarifario1").Caption = lcServicioDelTarifario
                'debb-09/09/2015 (inicio)
                CePreCuenta.Sections("Detalle").Controls("mensaje").Caption = lcBuscaParametro.SeleccionaFilaParametro(346)
                CePreCuenta.Sections("Detalle").Controls("mensaje1").Caption = lcBuscaParametro.SeleccionaFilaParametro(346)
                'debb-09/09/2015 (fin)
                CePreCuenta.Orientation = rptOrientPortrait
                If lbSePuedeImprimirPDF = True Then
                   CePreCuenta.PrintReport
                   
                Else
                    If lcParametro216 <> "1" Then
                       CePreCuenta.Show 1
                    Else
                       CePreCuenta.PrintReport
                    End If
                End If
                SeteaOtraImpresoraDefault sighEntidades.ImpresoraDefaultDeEstaPC
            End If
            '
            If Not (lcImpresora = "" Or lcImpresora = "0") Then
               sighEntidades.ImpresoraPredeterminada lcImpresoraActual
            End If
            '
        'End If
        'Set rsReporte = Nothing
        Set mrs_Tmp = Nothing
        Set mo_ReglasCaja = Nothing
        Set mo_ReglasComunes = Nothing
        Set mo_ReglasArchivoClinico = Nothing
        Set mo_ReglasDeProgMedica = Nothing
End Sub




Sub ImpresionFacturaEnDosTYPE(nroSerie As String, nroDcto As String, lcFarmaciaServicio As Long)
  Dim lnLineaSiguienteIMPR As Long, lnNumLetrasIMPR As Long
  Dim rsreporte As New ADODB.Recordset, rsReporte1 As New ADODB.Recordset
  Dim rsFacturaDetallada As New ADODB.Recordset
  Dim lnTotal As Double, lnImporteEXO As Double, lnSubTotal As Double
  Dim lcAdelantos As String, lcCuenta As String, lcCajero As String
  Dim lnFila As Long, lnCant As Integer, lnPrimeraFila As Long
  Dim mo_ReporteUtil As New ReporteUtil
  Dim oReglasCaja As New SIGHNegocios.ReglasCaja
  Dim NumeroRegistro As Integer 'Declara variables.
  Dim MiRegistroImpresion As ImpresionDOS
  Dim lcCantidad As String, lcPrecioUnitario As String
  Dim lcSubTotal As String, lcCajero1 As String
  Dim lcProducto As String, lcArchivo As String
  Dim lcBuscaParametro As New SIGHDatos.Parametros
  Dim lbSigue As Boolean, lcPuntoCarga As String
  Dim lcNombrePuntoCarga As String, lnAdelantos As Double
  Dim IGV As Double
  Dim lnVigv As Double, lnVsubTotal As Double
  IGV = Val(lcBuscaParametro.SeleccionaFilaParametro(221)) / 100
  On Error GoTo ManejadorError
  
  If lcFarmaciaServicio = sighEntidades.sghbien Then
    Set rsFacturaDetallada = oReglasCaja.CajaComprobantePagoProductosPorNroSerieNroDocumento(nroSerie, nroDcto)
  Else
    Set rsFacturaDetallada = oReglasCaja.CajaComprobantePagoServiciosPorNroSerieNroDocumento(nroSerie, nroDcto)
  End If
  
  If rsFacturaDetallada.RecordCount > Val(lcBuscaParametro.SeleccionaFilaParametro(103)) Then
    lnVigv = IIf(IsNull(rsreporte.Fields!IGV), 0, rsreporte.Fields!IGV)
    lnVsubTotal = rsFacturaDetallada.Fields!Subtotal
    With rsreporte
      .Fields.Append "TotalBoleta", adCurrency
      .Fields.Append "dctos", adCurrency
      .Fields.Append "idCuentaAtencion", adInteger
      .Fields.Append "IdOrden", adInteger
      .Fields.Append "IdOrdenPago", adInteger
      .Fields.Append "IdPreVenta", adInteger
      .Fields.Append "IdCajero", adInteger
      .Fields.Append "idTipoPago", adInteger
      .Fields.Append "idEstadoComprobante", adInteger
      .Fields.Append "razonSocial", adVarChar, 50, adFldIsNullable
      .Fields.Append "NroHistoriaClinica", adInteger
      .Fields.Append "FechaCobranza", adDate
      .Fields.Append "totalPorPagar", adCurrency
      .Fields.Append "cantidad", adInteger
      .Fields.Append "precioUnitario", adCurrency
      .Fields.Append "NombreProducto", adVarChar, 150, adFldIsNullable
      .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
      .Fields.Append "exoneraciones", adCurrency
      .Fields.Append "adelantos", adCurrency
      .Fields.Append "Observaciones", adVarChar, 100, adFldIsNullable
      .LockType = adLockOptimistic
      .Open
    End With
    If rsFacturaDetallada.RecordCount > 0 Then
      rsFacturaDetallada.MoveFirst
      Do While Not rsFacturaDetallada.EOF
        If lcFarmaciaServicio = sighEntidades.sghbien Then
          lcPuntoCarga = "1"
          lcNombrePuntoCarga = "FARMACIA"
        Else
          Set rsReporte1 = oReglasCaja.FactOrdenServicioSeleccionarPuntoCargaPorIdOrden(rsFacturaDetallada.Fields!IdOrden)
          lcPuntoCarga = "0"
          lcNombrePuntoCarga = ".."
          If rsReporte1.RecordCount > 0 Then
            lcPuntoCarga = Trim(Str(rsReporte1.Fields!idPuntoCarga))
            lcNombrePuntoCarga = Trim(rsReporte1.Fields!dPuntoCArga)
          End If
        End If
        lbSigue = True
        If rsreporte.RecordCount > 0 Then
          rsreporte.MoveFirst
          rsreporte.Find "codigo='" & lcPuntoCarga & "'"
          If Not rsreporte.EOF Then lbSigue = False
        End If
        If lbSigue = True Then
          rsreporte.AddNew
          rsreporte.Fields!Codigo = lcPuntoCarga
          rsreporte.Fields!NombreProducto = lcNombrePuntoCarga
          rsreporte.Fields!TotalBoleta = rsFacturaDetallada.Fields!TotalBoleta
          rsreporte.Fields!idCuentaAtencion = rsFacturaDetallada.Fields!idCuentaAtencion
          If lcFarmaciaServicio = sighEntidades.sghbien Then
             If IsNull(rsFacturaDetallada.Fields!idPreVenta) Then
                rsreporte.Fields!IdOrden = rsFacturaDetallada.Fields!IdOrden
             Else
                rsreporte.Fields!idPreVenta = rsFacturaDetallada.Fields!idPreVenta
             End If
          Else
             rsreporte.Fields!IdOrdenPago = rsFacturaDetallada.Fields!IdOrdenPago
          End If
          rsreporte.Fields!IdCajero = rsFacturaDetallada.Fields!IdCajero
          rsreporte.Fields!IdTipoPago = rsFacturaDetallada.Fields!IdTipoPago
          rsreporte.Fields!idEstadoComprobante = rsFacturaDetallada.Fields!idEstadoComprobante
          rsreporte.Fields!razonSocial = rsFacturaDetallada.Fields!razonSocial
          rsreporte.Fields!NroHistoriaClinica = rsFacturaDetallada.Fields!NroHistoriaClinica
          rsreporte.Fields!FechaCobranza = rsFacturaDetallada.Fields!FechaCobranza
          rsreporte.Fields!Dctos = rsFacturaDetallada.Fields!Dctos
          rsreporte.Fields!TotalPorPagar = rsFacturaDetallada.Fields!TotalPorPagar
          rsreporte.Fields!Cantidad = 0
          rsreporte.Fields!PrecioUnitario = 0
          rsreporte.Fields!exoneraciones = rsFacturaDetallada.Fields!exoneraciones
          rsreporte.Fields!Observaciones = rsFacturaDetallada.Fields!Observaciones
        Else
          rsreporte.Fields!TotalPorPagar = rsreporte.Fields!TotalPorPagar + rsFacturaDetallada.Fields!TotalPorPagar
          rsreporte.Fields!Cantidad = 0
          rsreporte.Fields!PrecioUnitario = 0
        End If
        rsreporte.Update
        rsFacturaDetallada.MoveNext
      Loop
    End If
  Else
    If lcFarmaciaServicio = sighEntidades.sghbien Then
      Set rsreporte = oReglasCaja.CajaComprobantePagoProductosPorNroSerieNroDocumento(nroSerie, nroDcto)
    Else
      Set rsreporte = oReglasCaja.CajaComprobantePagoServiciosPorNroSerieNroDocumento(nroSerie, nroDcto)
    End If
  End If
  If rsreporte.RecordCount > 0 Then
    lnVigv = IIf(IsNull(rsreporte.Fields!IGV), 0, rsreporte.Fields!IGV)
    lnVsubTotal = rsFacturaDetallada.Fields!Subtotal
    If lcFarmaciaServicio = sighEntidades.sghbien Then
      If IsNull(rsreporte.Fields!idPreVenta) Then
        lcCuenta = "    Ord.Pag: " & Trim(Str(rsreporte.Fields!IdOrden)) & "      N°Cta: " & Trim(Str(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion)))
      Else
        lcCuenta = "PreVenta: " & Trim(Str(rsreporte.Fields!idPreVenta)) & "      N°Cta: " & Trim(Str(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion)))
      End If
    Else
      lcCuenta = "Ord.Pag: " & Trim(Str(rsreporte.Fields!IdOrdenPago)) & "      N°Cta: " & Trim(Str(IIf(IsNull(rsreporte.Fields!idCuentaAtencion), 0, rsreporte.Fields!idCuentaAtencion)))
    End If
    
    '================================ Inicia Impresión
InicioImp:
    rsreporte.MoveFirst
    lnTotal = IIf(IsNull(rsreporte.Fields!TotalBoleta), "0", rsreporte.Fields!TotalBoleta)
    lcAdelantos = "Pago a Cuenta: " & Format(IIf(IsNull(rsreporte.Fields!Dctos), 0, rsreporte.Fields!Dctos), "0.00")
    lnAdelantos = rsreporte.Fields!Dctos
    'lcCuenta = "Ord.Pag: " & Trim(Str(rsReporte.Fields!IdOrden)) & "      N°Cta: " & Trim(Str(IIf(IsNull(rsReporte.Fields!idCuentaAtencion), 0, rsReporte.Fields!idCuentaAtencion)))
    lcCajero1 = oReglasCaja.SeleccionaDatosCajero(rsreporte.Fields!IdCajero, sghIniciales)
    lnImporteEXO = rsreporte.Fields!exoneraciones
    
    Dim Texto As String
    Printer.Font = "Verdana"
    Printer.FontSize = 10
    'Printer.PrintQuality = -1
    Printer.ScaleMode = 7
  
Cabecera:
    Printer.CurrentY = 0.9
    Printer.Print Tab(20); IIf(rsreporte!idEstadoComprobante = 9, "A N U L A D O ", "")
    Printer.Print Tab(93); nroSerie & " - " & nroDcto
    Printer.Print Tab(35); ""
    Printer.CurrentY = Printer.CurrentY + 0.15
    Printer.Print Tab(20); IIf(rsreporte!IdTipoPago = 2, "DEVOLUCION", IIf(rsreporte!TotalBoleta = 0, "EXONERADO", ""))
    Printer.Print Tab(35); ""
    Printer.Print Tab(35); ""
    Printer.CurrentY = Printer.CurrentY + 0.07
    Printer.Print Tab(20); lcCuenta
    Printer.Print Tab(18); oReglasCaja.NombreServicioPorCuentaAtencion(IIf(IsNull(rsreporte!idCuentaAtencion), 0, rsreporte!idCuentaAtencion)) & "   " & Trim(rsreporte.Fields!Observaciones)
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Tab(18); Left(Trim(IIf(IsNull(rsreporte.Fields!razonSocial), "", UCase(rsreporte.Fields!razonSocial))) & IIf(IsNull(rsreporte!NroHistoriaClinica), " (Particular)", " (HC: " & rsreporte!NroHistoriaClinica & ")"), 70)
    Dim D As Date
    D = CDate(rsreporte!FechaCobranza)
    Printer.CurrentY = Printer.CurrentY + 0.38
    Printer.Print Tab(18); IIf(IsNull(rsreporte.Fields!ruc), "", UCase(rsreporte.Fields!ruc)); Tab(95); Format(D, "dd"); Tab(105); Format(D, "mm"); Tab(115); Format(D, "yyyy")
    Printer.CurrentY = Printer.CurrentY + 0.38
    Printer.Print Tab(18); "" '--> aqui completar Dirección del usuario
    Printer.Print Tab(35); ""
    Printer.Print Tab(35); ""
    Printer.CurrentY = Printer.CurrentY + 0.2

Detalle:
    NumeroRegistro = 1
    Do While Not rsreporte.EOF
      Printer.Print Tab(13); Format(rsreporte!Cantidad, "0"); Tab(19); rsreporte!Codigo; Tab(27); Left(rsreporte!NombreProducto, 40); Tab(95); Format(rsreporte!PrecioUnitario, "0.00"); Tab(112); Format(rsreporte!TotalPorPagar, "0.00")
      Printer.CurrentY = Printer.CurrentY + 0.21
      NumeroRegistro = NumeroRegistro + 1
      rsreporte.MoveNext
    Loop
  
    If NumeroRegistro < 10 Then
      Do While NumeroRegistro < 10
        Printer.Print
        Printer.CurrentY = Printer.CurrentY + 0.21
        NumeroRegistro = NumeroRegistro + 1
      Loop
    End If
  
Fin:
    Printer.CurrentY = Printer.CurrentY - 0.7
    Printer.Print Tab(15); sighEntidades.Numlet(sighEntidades.DevuelveNumeroSinDecimales(lnTotal)) & " CON " & sighEntidades.DevuelveSoloDecimales(lnTotal) & "/100"
    Printer.Print Tab(8); lcCajero1, Tab(30); Format(lnVsubTotal, "0.00")   'Format(lnTotal + lnImporteEXO + lnAdelantos, "0.00")
    Printer.Print Tab(112); Format(lnVigv, "0.00")                          'Printer.Print Tab(112); Format(lnTotal / (IGV + 1), "0.00")
    Printer.Print Tab(33); lcAdelantos
    Printer.CurrentY = Printer.CurrentY - 0.1
    Printer.Print Tab(48); Format(D, "dd"); Tab(54); Format(D, "mmmm"); Tab(75); Format(D, "yy"); Tab(112); Format(lnTotal * IGV / (IGV + 1), "0.00")
    Printer.Print
    Printer.CurrentY = Printer.CurrentY - 0.1
    Printer.Print Tab(33); "DCTO:   " & Format(lnImporteEXO, "0.00"); Tab(112); Format(lnTotal, "0.00")
    Printer.EndDoc
    rsreporte.Close
    Exit Sub
    '================================
  End If
  Exit Sub
  
ManejadorError:
  If Err.Number = 55 Then
    Resume
  Else
    If Err.Number = 53 Then
      Resume Next
    Else
      MsgBox Err.Description
    End If
  End If
End Sub


Sub CentroCostoDetallado(lbEnExcel As Boolean, lcTitulo As String, lcSubTitulo As String, ml_IdCaja As Long, mda_FechaInicio As Date, mda_FechaFin As Date, ml_IdTurno As Long, ml_IdCajero As Long, ml_IdCentroCostos As Long, lnHwnd As Long)
        Dim mrs_Tmp As New Recordset
        Dim rsreporte As New Recordset
        Dim rsTmp As New Recordset
        Dim LcTexto2 As String
        Dim lnPagoCta As Double, lnExoneracion As Double, lnAnulado As Double
        Dim lnImptotal As Double
        Dim lbPrimeraVez As Boolean
        Dim lnCanExoneracion As Long, lnCanTotal As Long, lnCanAnulado As Long
        Dim lcPie As String
        With mrs_Tmp
            .Fields.Append "idProducto", adInteger
            .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
            .Fields.Append "Producto", adVarChar, 100, adFldIsNullable
            .Fields.Append "Precio", adDouble
            .Fields.Append "SubTotal", adDouble
            .Fields.Append "canSubTotal", adDouble
            .Fields.Append "ImpExonerado", adDouble
            .Fields.Append "CanExonerado", adInteger
            .Fields.Append "ImpAnulado", adDouble
            .Fields.Append "CanAnulado", adInteger
            .Fields.Append "PagoCta", adDouble
            .Fields.Append "ImpTotal", adDouble
            .Fields.Append "CanTotal", adInteger
            .LockType = adLockOptimistic
            .Open
        End With
        'SERVICIOS emitidos en CAJA SERVICIO
        Set rsreporte = mo_ReglasFacturacion.ServicioConsolidado(ml_IdCaja, mda_FechaInicio, mda_FechaFin, ml_IdTurno, ml_IdCajero)
        rsreporte.Filter = "idCentroCosto=" & ml_IdCentroCostos
        If rsreporte.RecordCount > 0 Then
           rsreporte.MoveFirst
           Do While Not rsreporte.EOF
              LcTexto2 = rsreporte.Fields!nroSerie + rsreporte.Fields!nrodocumento
              lnPagoCta = rsreporte.Fields!Adelantos
              Do While Not rsreporte.EOF And LcTexto2 = (rsreporte.Fields!nroSerie + rsreporte.Fields!nrodocumento)
                    lbPrimeraVez = True
                    If mrs_Tmp.RecordCount > 0 Then
                       mrs_Tmp.MoveFirst
                       mrs_Tmp.Find "idProducto=" & rsreporte.Fields!idProducto
                       If Not mrs_Tmp.EOF Then
                          lbPrimeraVez = False
                       End If
                    End If
                    lnExoneracion = 0: lnCanExoneracion = 0
                    If rsreporte.Fields!exoneraciones > 0 Then
                       Set rsTmp = mo_ReglasFacturacion.facturacionServicioFinanciamientosSoloExonerados(rsreporte.Fields!IdOrden, rsreporte.Fields!idProducto)
                       If rsTmp.RecordCount > 0 Then
                          lnExoneracion = rsTmp.Fields!TotalFinanciado
                          lnCanExoneracion = rsTmp.Fields!CantidadFinanciada
                       End If
                       rsTmp.Close
                    End If
                    lnAnulado = 0: lnCanAnulado = 0
                    If rsreporte.Fields!idEstadoComprobante = 9 Then
                       lnAnulado = rsreporte.Fields!Total
                       lnCanAnulado = rsreporte.Fields!Cantidad
                    End If
                    If lbPrimeraVez = True Then
                          mrs_Tmp.AddNew
                          mrs_Tmp.Fields!idProducto = rsreporte.Fields!idProducto
                          mrs_Tmp.Fields!Producto = Left(rsreporte.Fields!dProducto, 100)
                          mrs_Tmp.Fields!Codigo = rsreporte.Fields!Codigo
                          mrs_Tmp.Fields!precio = rsreporte.Fields!precio
                    End If
                    lnImptotal = 0: lnCanTotal = 0
                    If lnAnulado = 0 Then
                       lnImptotal = rsreporte.Fields!Total - lnExoneracion - lnPagoCta
                       lnCanTotal = rsreporte.Fields!Cantidad
                    End If
                    mrs_Tmp.Fields!Subtotal = mrs_Tmp.Fields!Subtotal + rsreporte.Fields!Total
                    mrs_Tmp.Fields!canSubtotal = mrs_Tmp.Fields!canSubtotal + rsreporte.Fields!Cantidad
                    mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracion
                    mrs_Tmp.Fields!CanExonerado = mrs_Tmp.Fields!CanExonerado + lnCanExoneracion
                    mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnulado
                    mrs_Tmp.Fields!CanAnulado = mrs_Tmp.Fields!CanAnulado + lnCanAnulado
                    mrs_Tmp.Fields!PagoCta = mrs_Tmp.Fields!PagoCta + lnPagoCta
                    mrs_Tmp.Fields!ImpTotal = mrs_Tmp.Fields!ImpTotal + lnImptotal
                    mrs_Tmp.Fields!CanTotal = mrs_Tmp.Fields!CanTotal + lnCanTotal
                    mrs_Tmp.Update
                    lnPagoCta = 0
                    rsreporte.MoveNext
                    If rsreporte.EOF Then
                       Exit Do
                    End If
              Loop
           Loop
        End If
        'MEDICAMENTOS emitidos en CAJA SERVICIO
        Set rsreporte = mo_ReglasFacturacion.FarmaciaConsolidado(ml_IdCaja, mda_FechaInicio, mda_FechaFin, ml_IdTurno, ml_IdCajero)
        rsreporte.Filter = "idCentroCosto=" & ml_IdCentroCostos
        If rsreporte.RecordCount > 0 Then
           rsreporte.MoveFirst
           Do While Not rsreporte.EOF
              LcTexto2 = rsreporte.Fields!nroSerie + rsreporte.Fields!nrodocumento
              lnPagoCta = rsreporte.Fields!Adelantos
              Do While Not rsreporte.EOF And LcTexto2 = (rsreporte.Fields!nroSerie + rsreporte.Fields!nrodocumento)
                    lbPrimeraVez = True
                    If mrs_Tmp.RecordCount > 0 Then
                       mrs_Tmp.MoveFirst
                       mrs_Tmp.Find "idProducto=" & rsreporte.Fields!idProducto
                       If Not mrs_Tmp.EOF Then
                          lbPrimeraVez = False
                       End If
                    End If
                    lnExoneracion = 0: lnCanExoneracion = 0
                    If rsreporte.Fields!exoneraciones > 0 Then
                       Set rsTmp = mo_ReglasFacturacion.facturacionBienesFinanciamientosSoloExonerados(rsreporte.Fields!movNumero, rsreporte.Fields!MovTipo, rsreporte.Fields!idProducto)
                       If rsTmp.RecordCount > 0 Then
                          lnExoneracion = rsTmp.Fields!TotalFinanciado
                          lnCanExoneracion = rsTmp.Fields!CantidadFinanciada
                       End If
                       rsTmp.Close
                    End If
                    lnAnulado = 0: lnCanAnulado = 0
                    If rsreporte.Fields!idEstadoComprobante = 9 Then
                       lnAnulado = rsreporte.Fields!TotalPagar
                       lnCanAnulado = rsreporte.Fields!CantidadPagar
                    End If
                    If lbPrimeraVez = True Then
                          mrs_Tmp.AddNew
                          mrs_Tmp.Fields!idProducto = rsreporte.Fields!idProducto
                          mrs_Tmp.Fields!Producto = rsreporte.Fields!Producto
                          mrs_Tmp.Fields!Codigo = rsreporte.Fields!Codigo
                          mrs_Tmp.Fields!precio = rsreporte.Fields!PrecioVenta
                    End If
                    lnImptotal = 0: lnCanTotal = 0
                    If lnAnulado = 0 Then
                       lnImptotal = rsreporte.Fields!TotalPagar - lnExoneracion - lnPagoCta
                       lnCanTotal = rsreporte.Fields!CantidadPagar
                    End If
                    mrs_Tmp.Fields!Subtotal = mrs_Tmp.Fields!Subtotal + rsreporte.Fields!TotalPagar
                    mrs_Tmp.Fields!canSubtotal = mrs_Tmp.Fields!canSubtotal + rsreporte.Fields!CantidadPagar
                    mrs_Tmp.Fields!ImpExonerado = mrs_Tmp.Fields!ImpExonerado + lnExoneracion
                    mrs_Tmp.Fields!CanExonerado = mrs_Tmp.Fields!CanExonerado + lnCanExoneracion
                    mrs_Tmp.Fields!ImpAnulado = mrs_Tmp.Fields!ImpAnulado + lnAnulado
                    mrs_Tmp.Fields!CanAnulado = mrs_Tmp.Fields!CanAnulado + lnCanAnulado
                    mrs_Tmp.Fields!PagoCta = mrs_Tmp.Fields!PagoCta + lnPagoCta
                    mrs_Tmp.Fields!ImpTotal = mrs_Tmp.Fields!ImpTotal + lnImptotal
                    mrs_Tmp.Fields!CanTotal = mrs_Tmp.Fields!CanTotal + lnCanTotal
                    mrs_Tmp.Update
                    lnPagoCta = 0
                    rsreporte.MoveNext
                    If rsreporte.EOF Then
                       Exit Do
                    End If
              Loop
           Loop
        End If
        lcPie = "Nro Tarifas: " & Trim(Str(mrs_Tmp.RecordCount))
        If lbEnExcel = True Then
           Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
           mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, lcSubTitulo, lcPie, lnHwnd
        Else
            Set ECentroCdetalle.DataSource = mrs_Tmp
            ECentroCdetalle.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
            ECentroCdetalle.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
            ECentroCdetalle.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
            ECentroCdetalle.Sections("cabecera").Controls("lblhora").Caption = lcBuscaParametro.RetornaHoraServidorSQL
            ECentroCdetalle.Sections("cabecera").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaServidorSQL
            ECentroCdetalle.Sections("cabecera").Controls("lblTitulo").Caption = lcTitulo
            ECentroCdetalle.Sections("cabecera").Controls("lblSubTitulo").Caption = lcSubTitulo
            Set ECentroCdetalle.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
            ECentroCdetalle.Sections("pie").Controls("lblPie").Caption = lcPie
            ECentroCdetalle.Orientation = rptOrientPortrait
            ECentroCdetalle.Show 1
            'debb-27/05/2015
            Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
            mo_ReglasComunes.grabaTablaAuditoria ("ECentroCdetalle: " & _
                                           Mid(lcSubTitulo, IIf(InStr(lcSubTitulo, "FILTROS: ") > 0, 10, 1)))
            Set mo_ReglasComunes = Nothing
            '
        End If
End Sub

'debb-01/03/2015
'Sub ImpresionOrdenMedica(lnIdReceta As Long, EsMedicamento As Boolean)
Sub ImpresionOrdenMedica(lnIdReceta As Long, lbImpresionIndirecta As Boolean)
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Dim rsreporte As New Recordset
        Dim rsReporte1 As New Recordset
        Dim mrs_Tmp As New Recordset
        Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
        Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
        Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
        Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
        Dim lcNombreHospital As String, lnIdprograma As Long, lnIdTurno As Long, lcTurno As String
        Dim lcRuc As String, lcDireccion As String, lcTlefono As String
        Dim lcDevuelveFechaServidor As New SIGHDatos.Parametros
        Dim FechaReceta As Date, numeroFiliacion As String, lcDx As String, lcFont_detalle As String, lcFont_detalle1 As String
        Dim oConexion As New Connection
        Dim lnPos1 As Integer, lnPos2 As Integer, lcSoloEstosDx As String, lcEPS As String
        
        
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        
        On Error GoTo ManejadorErrorImprime
        Set rsreporte = mo_ReglasComunes.RecetaCabeceraPoridReceta(lnIdReceta)
        If wxParametro351 = "S" Then
           Set rsreporte = HCigualDNI_DevuelveRsConHistoriaOCHOdigitos(rsreporte, "NroHistoriaClinica")
        End If
        
        Set rsReporte1 = mo_ReglasComunes.RecetaDetalleServicioPorIdReceta(lnIdReceta)
        
        lcSoloEstosDx = ""
        If wxParametro545 = "S" Then
           If rsReporte1.RecordCount > 0 Then
              lcSoloEstosDx = "/"
              rsReporte1.MoveFirst
              Do While Not rsReporte1.EOF
                 If InStr(lcSoloEstosDx, Trim(rsReporte1!Dx)) = 0 Then
                    lcSoloEstosDx = lcSoloEstosDx & Trim(rsReporte1!Dx) & "/"
                 End If
                 rsReporte1.MoveNext
              Loop
              rsReporte1.MoveFirst
           End If
        End If
        
        If UCase(Left(rsreporte!TipoServicio, 3)) = "HOS" Then
           lcDx = mo_ReglasFacturacion.DevuelveDxTodosIngresoSoloHospEmerg(3, rsreporte!idAtencion, oConexion, lcSoloEstosDx)
        ElseIf InStr(UCase(rsreporte!TipoServicio), "EMERG") > 0 Then
           lcDx = mo_ReglasFacturacion.DevuelveDxTodosIngresoSoloHospEmerg(2, rsreporte!idAtencion, oConexion, lcSoloEstosDx)
        Else
           lcDx = mo_ReglasFacturacion.DevuelveDxAltaMedicaTodosDx(rsreporte!idAtencion, 1, lcSoloEstosDx)
        End If
        
        lcNombreHospital = lcBuscaParametro.SeleccionaFilaParametro(205)
        lcRuc = lcBuscaParametro.SeleccionaFilaParametro(339)
        lcDireccion = lcBuscaParametro.SeleccionaFilaParametro(206)
        lcTlefono = lcBuscaParametro.SeleccionaFilaParametro(207)
        If rsreporte!IdFormaPago = 2 Then
           numeroFiliacion = BuscarNumeroAfiliacion(rsreporte.Fields!idAtencion, rsreporte.Fields!idCuentaAtencion)
        Else
           numeroFiliacion = ""
        End If
        '
        lcFont_detalle = ""
        lcFont_detalle1 = lcBuscaParametro.SeleccionaFilaParametro(507)
        If lcFont_detalle1 <> "" Then
            lnPos1 = InStr(lcFont_detalle1, "/")
            lnPos2 = InStr(lnPos1 + 1, lcFont_detalle1, "/")
            If lnPos1 > 0 And lnPos2 > lnPos1 Then
                If InStr(UCase(rsreporte!TipoServicio), "HOSP") > 0 Then
                      If (lnPos2 - lnPos1 - 1) > 2 Then
                         lcFont_detalle = Mid(lcFont_detalle1, lnPos1 + 1, lnPos2 - lnPos1 - 1)
                      End If
                ElseIf InStr(UCase(rsreporte!TipoServicio), "EMER") > 0 Then
                      If lnPos2 < Len(lcFont_detalle1) Then
                          lcFont_detalle = Mid(lcFont_detalle1, lnPos2 + 1)
                      End If
                Else
                      If lnPos1 > 2 Then
                         lcFont_detalle = Left(lcFont_detalle1, lnPos1 - 1)
                      End If
                End If
            End If
        End If
        '
'        lcEPS = ""
'        If Not IsNull(rsreporte!EpsPorcentaje) Then
'           If rsreporte!EpsPorcentaje > 0 Then
'              lcEPS = " (cubre: " & Trim(Str(rsreporte!EpsPorcentaje)) & " % )"
'           End If
'        End If
        Dim lnEpsPorcentaje As Double
        lnEpsPorcentaje = mo_ReporteUtil.DevuelveEpsPorcentaje(rsreporte!EpsPorcentaje)
        lcEPS = mo_ReporteUtil.DevuelveEPScubre(lnEpsPorcentaje)
        
        '
        Set OrdenMedica.DataSource = rsReporte1
        OrdenMedica.Sections("Cabecera").Controls("lcHospital1").Caption = UCase(lcNombreHospital)
        OrdenMedica.Sections("Cabecera").Controls("DatosEESS").Caption = UCase("RUC: " & lcRuc) & vbCrLf & UCase(lcDireccion) & vbCrLf & "Telef: " & UCase(lcTlefono)
        OrdenMedica.Sections("Cabecera").Controls("NombreTicket").Caption = UCase("ORDEN MEDICA")
        OrdenMedica.Sections("Cabecera").Controls("FechaReceta").Caption = "Fecha/Hora Atencion: " & rsreporte.Fields!FechaReceta
        OrdenMedica.Sections("Cabecera").Controls("PuntoCarga").Caption = IIf(rsreporte!idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion, _
                   "REALIZAR en: " & mo_CabeceraReportes.RetornaNombrePuntoCargaRecetaConsumoEnServicio(rsreporte!idReceta), _
                                                                          "Servicio: " & UCase(Trim(rsreporte.Fields!PuntoCarga)))
        OrdenMedica.Sections("Cabecera").Controls("Servicio").Caption = "Consultorio: " & UCase(rsreporte.Fields!Servicio)
        OrdenMedica.Sections("Cabecera").Controls("OrdenMedica").Caption = "NºOrd.Med: " & rsreporte.Fields!idReceta
        OrdenMedica.Sections("Cabecera").Controls("lcNroCuenta").Caption = "NºCta: " & rsreporte.Fields!idCuentaAtencion
        OrdenMedica.Sections("Cabecera").Controls("Medico").Caption = "Prof. de la Salud: " & UCase(rsreporte.Fields!NombreMedico)
        OrdenMedica.Sections("Cabecera").Controls("Paciente").Caption = "Paciente: " & rsreporte.Fields!ApellidoPaterno & " " & rsreporte.Fields!ApellidoMaterno & " " & rsreporte.Fields!PrimerNombre
        OrdenMedica.Sections("Cabecera").Controls("TipoPlan").Caption = "Tipo Plan: " & Trim(rsreporte.Fields!descripcion) & lcEPS & _
                                                                        IIf(numeroFiliacion = "", "", " (Nº Afil: " & numeroFiliacion & ")")
        OrdenMedica.Sections("Cabecera").Controls("dx").Caption = lcDx
        OrdenMedica.Sections("Cabecera").Controls("HCyEdad").Caption = "Hc: " & Trim(Str(rsreporte!NroHistoriaClinica)) & _
                                   "  (Edad: " & rsreporte!EdadPaciente & ")"

        OrdenMedica.Sections("PieInforme").Controls("Terminal").Caption = "Terminal:  " & lcBuscaParametro.RetornaNombreDeServidor
        OrdenMedica.Sections("PieInforme").Controls("lcUsuario").Caption = "Fecha: " & Format(lcDevuelveFechaServidor.RetornaFechaHoraServidorSQL, sighEntidades.DevuelveFechaSoloFormato_DMY_HM)
        OrdenMedica.Sections("PieInforme").Controls("TextoVacio").Caption = lcBuscaParametro.SeleccionaFilaParametro(346)
        
        If lcFont_detalle <> "" Then
            OrdenMedica.Sections("Detalle").Controls("txtNombreProducto").Font.Name = lcFont_detalle
            OrdenMedica.Sections("Detalle").Controls("txtTotalPorPagar").Font.Name = lcFont_detalle
        End If
        
        OrdenMedica.RightMargin = 100
        OrdenMedica.TopMargin = 100
        OrdenMedica.LeftMargin = 100
        OrdenMedica.BottomMargin = 100
        OrdenMedica.Orientation = rptOrientPortrait
        
        If lbImpresionIndirecta = False Then
            OrdenMedica.Show 1
        Else
            OrdenMedica.PrintReport
        End If
        Set rsreporte = Nothing
        Set mrs_Tmp = Nothing
        Set mo_ReglasCaja = Nothing
        Set mo_ReglasComunes = Nothing
        Set mo_ReglasArchivoClinico = Nothing
        Set mo_ReglasDeProgMedica = Nothing
        Set oConexion = Nothing
        
        Exit Sub
        
ManejadorErrorImprime:
        Select Case Err.Number
        Case 1004
            MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
        Case 8555
            MsgBox "No se puede obtener información de la impresora, seleccione Configuración en el menú Inicio de Windows, haga clic en Impresoras y después ponga una impresora como predeterminada. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
        Case Else
            MsgBox Err.Description
            Resume
        End Select
End Sub
'Actualizado 26092014
Public Sub imprimirReceta(idReceta As Long, numeroFiliacion As String, _
        TipoServicio As String, lbImpresionIndirecta As Boolean)
    Dim crApp As New CRAXDRT.Application
    Dim crReport As New CRAXDRT.Report
    Dim crParameters As CRAXDRT.ParameterFieldDefinitions
    Dim crParameter As CRAXDRT.ParameterFieldDefinition
    Dim rutaRaiz As String
    rutaRaiz = App.Path
    If InStr(rutaRaiz, "SIGH") > 0 Then
        rutaRaiz = rutaRaiz & "/Plantillas"
    Else
        rutaRaiz = rutaRaiz & "/Plantillas"
        If Dir$(rutaRaiz, vbDirectory) = "" Then
            rutaRaiz = App.Path & "/GalenHos/Plantillas"
        End If
    End If
    Set crReport = crApp.OpenReport(rutaRaiz & "/rptRecetaMaestro.rpt", 1)
    
'    crReport.DiscardSavedData
'    crReport.Database.SetDataSource crearRecorsetTemporal(rsReporte)
'    rsReporte.MoveFirst
'    FormBoleta.Caption = lcTipoComprobante
    
    Set crParameters = crReport.ParameterFields
    For Each crParameter In crParameters
        Select Case crParameter.ParameterFieldName
            Case "@lnidReceta":
                crParameter.AddCurrentValue (idReceta)
            Case "@NumeroFiliacion":
                crParameter.AddCurrentValue (numeroFiliacion)
            Case "@NombreServicio":
                crParameter.AddCurrentValue (TipoServicio)
        End Select
    Next
    
    crReport.PaperOrientation = crDefaultPaperOrientation
    
    If lbImpresionIndirecta = False Then
        Dim ViewReporte As New FormBoletaRpt
        FormBoletaRpt.Caption = "Titulo"
        FormBoletaRpt.CrvBoleta.ReportSource = crReport
        FormBoletaRpt.CrvBoleta.ViewReport
        FormBoletaRpt.Show 1
    Else
       On Error Resume Next
       crReport.PrintOut False
    End If
End Sub
'debb-15/04/2016
Public Sub imprimirReceta_Version2(lIdReceta As Long, lbImpresionIndirecta As Boolean, lnWnd As Long)
    'variables de accedo a datos
    Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
    Dim rsRecetaCabecera As New ADODB.Recordset
    Dim rsRecetaDetalle As New ADODB.Recordset
    Dim rsRecetaDiagnosticos As New ADODB.Recordset
    Dim RsDxTmp As New Recordset
    
    Dim lIdAtencion As Long
    Dim cFechaAtencion As String
    Dim lIdCuentaAtencion As Long
    Dim TipoServicio As String
    Dim numeroFiliacion As String
    Dim cFechaVigencia As String   'debb-24/06/2015
    
    'variables para crystal reports
    Dim crApp As New CRAXDRT.Application
    Dim crReport As New CRAXDRT.Report
    Dim crParameters As CRAXDRT.ParameterFieldDefinitions
    Dim crParameter As CRAXDRT.ParameterFieldDefinition
    
    Dim crReportRecetaDetalle As CRAXDRT.Report
    Dim crReportDiagnostico As CRAXDRT.Report
    Dim crReportIndicaciones As CRAXDRT.Report
    Dim crReportCabecera As CRAXDRT.Report
    Dim crReportCabeceraIndicaciones As CRAXDRT.Report
    
    Dim crSections As CRAXDRT.Sections
    Dim crSection As CRAXDRT.Section
    Dim crReportObject As CRAXDRT.ReportObjects
    Dim crSubReport As CRAXDRT.SubreportObject
    'ruta de los reportes
    Dim rutaRaiz As String, lbAdd As Boolean, lcEPS As String
    
    On Error GoTo ManejadorErrorImprime
    
    rutaRaiz = App.Path
    If InStr(rutaRaiz, "SIGH") > 0 Then
        rutaRaiz = rutaRaiz & "/Plantillas"
    Else
        rutaRaiz = rutaRaiz & "/Plantillas"
        If Dir$(rutaRaiz, vbDirectory) = "" Then
            rutaRaiz = App.Path & "/GalenHos/Plantillas"
        End If
    End If
    
    
    'datos de las recetas
    Set rsRecetaCabecera = mo_ReglasComunes.RecetaCabeceraPoridReceta(lIdReceta)
    If wxParametro351 = "S" Then
       Set rsRecetaCabecera = HCigualDNI_DevuelveRsConHistoriaOCHOdigitos(rsRecetaCabecera, "nroHistoriaClinica")
    End If
    
    
    
    
    Set rsRecetaDetalle = mo_ReglasComunes.RecetaDetallePorIdReceta(lIdReceta)
    
    SeteaOtraImpresoraDefault sighEntidades.ImpresoraDefaultDeEstaPC
    If rsRecetaDetalle.RecordCount > 8 Then
       Set crReport = crApp.OpenReport(rutaRaiz & "/rptRecetaMaestroV.rpt", 1)
    Else
       Set crReport = crApp.OpenReport(rutaRaiz & "/rptRecetaMaestro.rpt", 1)
    End If
    
    lIdAtencion = 0
    If rsRecetaCabecera.RecordCount > 0 Then
        rsRecetaCabecera.MoveFirst
        '
'        lcEPS = ""
'        If Not IsNull(rsRecetaCabecera!EpsPorcentaje) Then
'           If rsRecetaCabecera!EpsPorcentaje > 0 Then
'              lcEPS = " (cubre: " & Trim(Str(rsRecetaCabecera!EpsPorcentaje)) & " % )"
'           End If
'        End If
        Dim lnEpsPorcentaje As Double
        lnEpsPorcentaje = mo_ReporteUtil.DevuelveEpsPorcentaje(rsRecetaCabecera!EpsPorcentaje)
        lcEPS = mo_ReporteUtil.DevuelveEPScubre(lnEpsPorcentaje)
        '
        lIdAtencion = rsRecetaCabecera.Fields!idAtencion
        cFechaAtencion = Format(rsRecetaCabecera.Fields!FechaAtencion, "dd/mm/yyyy")
        lIdCuentaAtencion = rsRecetaCabecera.Fields!idCuentaAtencion
        TipoServicio = rsRecetaCabecera.Fields!TipoServicio
        If rsRecetaCabecera!IdFormaPago = 2 Then
           numeroFiliacion = BuscarNumeroAfiliacion(lIdAtencion, lIdCuentaAtencion)
        Else
           numeroFiliacion = "" + lcEPS
        End If
        cFechaVigencia = IIf(IsNull(rsRecetaCabecera!FechaVigencia), "", rsRecetaCabecera!FechaVigencia)
    End If
    
    If wxParametro545 <> "S" Then
       Set rsRecetaDiagnosticos = mo_ReglasComunes.RecetaDiagnosticosPorIdAtencion(lIdAtencion)
    Else
        With rsRecetaDiagnosticos
              .Fields.Append "Tipo", adVarChar, 30
              .Fields.Append "idDiagnostico", adInteger
              .Fields.Append "codigoCIE10", adVarChar, 20
              .Fields.Append "descripcion", adVarChar, 255
              .LockType = adLockOptimistic
              .Open
        End With
        Set RsDxTmp = mo_ReglasComunes.RecetaDiagnosticosPorIdAtencion(lIdAtencion)
        If RsDxTmp.RecordCount > 0 Then
           RsDxTmp.MoveFirst
           Do While Not RsDxTmp.EOF
              rsRecetaDetalle.MoveFirst
              rsRecetaDetalle.Find "dx='" & RsDxTmp!CodigoCiE10 & "'"
              If Not rsRecetaDetalle.EOF Then
                 lbAdd = True
                 If rsRecetaDiagnosticos.RecordCount > 0 Then
                    rsRecetaDiagnosticos.MoveFirst
                    rsRecetaDiagnosticos.Find "codigoCIE10='" & RsDxTmp!CodigoCiE10 & "'"
                    If Not rsRecetaDiagnosticos.EOF Then
                       lbAdd = False
                    End If
                 End If
                 If lbAdd = True Then
                    rsRecetaDiagnosticos.AddNew
                    rsRecetaDiagnosticos!tipo = IIf(IsNull(RsDxTmp!tipo), "", RsDxTmp!tipo)
                    rsRecetaDiagnosticos!idDiagnostico = RsDxTmp!idDiagnostico
                    rsRecetaDiagnosticos!CodigoCiE10 = RsDxTmp!CodigoCiE10
                    rsRecetaDiagnosticos!descripcion = RsDxTmp!descripcion
                    rsRecetaDiagnosticos.Update
                 End If
              End If
              RsDxTmp.MoveNext
           Loop
        End If
     End If
    
'    crReport.DiscardSavedData
'    crReport.Database.SetDataSource crearRecorsetTemporal(rsReporte)
'    rsReporte.MoveFirst
'    FormBoleta.Caption = lcTipoComprobante
    
    Set crParameters = crReport.ParameterFields
    For Each crParameter In crParameters
        Select Case crParameter.ParameterFieldName
            Case "@lnidReceta":
                crParameter.AddCurrentValue (lIdReceta)
            Case "@NumeroFiliacion":
                crParameter.AddCurrentValue (numeroFiliacion)
            Case "@NombreServicio":
                crParameter.AddCurrentValue (TipoServicio)
            Case "@FechaAtencion":
                crParameter.AddCurrentValue (cFechaAtencion)
            Case "@fechaVigencia":                                    'debb-24/06/2015
                crParameter.AddCurrentValue (cFechaVigencia)         'debb-24/06/2015
        End Select
    Next
    
    
    'setar recorset al reporte principal y a los subreportes
    crReport.Database.SetDataSource rsRecetaCabecera
    
    Set crSections = crReport.Sections
    
    Dim i As Integer, X As Integer
    
    For i = 1 To crSections.Count
        Set crSection = crSections.Item(i)
        Set crReportObject = crSection.ReportObjects
        
        For X = 1 To crReportObject.Count
            If crReportObject.Item(X).Kind = crSubreportObject Then
                Select Case crReportObject.Item(X).SubreportName
                    Case "rptRecetaCabecera.rpt":
                        Set crReportCabecera = crReport.OpenSubreport(crReportObject.Item(X).SubreportName)
                        crReportCabecera.Database.SetDataSource rsRecetaCabecera
                    Case "rptRecetaCabeceraIndicaciones.rpt":
                        Set crReportCabeceraIndicaciones = crReport.OpenSubreport(crReportObject.Item(X).SubreportName)
                        crReportCabeceraIndicaciones.Database.SetDataSource rsRecetaCabecera
                    Case "rptRecetaIndicaciones.rpt":
                        Set crReportIndicaciones = crReport.OpenSubreport(crReportObject.Item(X).SubreportName)
                        crReportIndicaciones.Database.SetDataSource rsRecetaDetalle
                    Case "rptRecetaDiagnostico.rpt":
                        Set crReportDiagnostico = crReport.OpenSubreport(crReportObject.Item(X).SubreportName)
                        crReportDiagnostico.Database.SetDataSource rsRecetaDiagnosticos
                    Case "rptRecetaDetalle.rpt":
                        Set crReportRecetaDetalle = crReport.OpenSubreport(crReportObject.Item(X).SubreportName)
                        crReportRecetaDetalle.Database.SetDataSource rsRecetaDetalle
                End Select
            End If
        Next X
    Next i
    
    crReport.PaperOrientation = crDefaultPaperOrientation
   
   
    
    If lbImpresionIndirecta = False Then
        If wxParametro513 = "S" Then
            crReport.PrinterSetup lnWnd
        End If
        Dim ViewReporte As New FormBoletaRpt
         
        FormBoletaRpt.Caption = "Titulo"
        FormBoletaRpt.CrvBoleta.ReportSource = crReport
        FormBoletaRpt.CrvBoleta.ViewReport
        FormBoletaRpt.Show 1
    Else
       On Error Resume Next       '
       crReport.PrintOut False
       'OrdenMedica.PrintReport
    End If
    Exit Sub
    
ManejadorErrorImprime:
        Select Case Err.Number
        Case 1004
            MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
        Case 8555
            MsgBox "No se puede obtener información de la impresora, seleccione Configuración en el menú Inicio de Windows, haga clic en Impresoras y después ponga una impresora como predeterminada. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
        Case Else
            If Err.Description = "File not found." Then
                MsgBox "No se puede encontrar las plantillas para la impresión de la receta", vbExclamation, "Reporte de historia clínica"
            Else
                MsgBox Err.Description
            End If
            Exit Sub
            Resume
        End Select
End Sub

'Actualizado 29092014
Public Function BuscarNumeroAfiliacion(lIdAtencion As Long, lIdCuentaAtencion As Long) As String
    Dim oConexion As New ADODB.Connection
    Dim mo_DoAtencionDatosAdicionales As New DoAtencionDatosAdicionales
    Dim oRsTmp1 As New Recordset
    Dim oRsAfiliadosSIS As New Recordset
    Dim lcAfiliacionIdSiaSis As String
    Dim lcAfiliacionCodigo As String
    Dim lcSql As String
    Dim wxParametroJAMO As String
    Dim lcFiliacionEnAdmision As String
    
    BuscarNumeroAfiliacion = ""
    wxParametroJAMO = lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
     
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighEntidades.CadenaConexion
    lcAfiliacionIdSiaSis = 0 'Actualizado 02102014
    Set mo_DoAtencionDatosAdicionales = mo_AdminAdmision.AtencionesDatosAdicionalesSeleccionarPorId(lIdAtencion, oConexion)
    If mo_DoAtencionDatosAdicionales.idSiaSis = 0 Then
       Set oRsTmp1 = mo_ReglasSISgalenhos.SisFuaAtencionSeleccionarPorCuenta(lIdCuentaAtencion)
       If oRsTmp1.RecordCount > 0 Then
            If Not IsNull(oRsTmp1.Fields!idSiaSis) Then
                lcAfiliacionIdSiaSis = oRsTmp1.Fields!idSiaSis
                lcAfiliacionCodigo = oRsTmp1.Fields!Codigo
            End If
       End If
    Else
        lcAfiliacionIdSiaSis = mo_DoAtencionDatosAdicionales.idSiaSis
        lcAfiliacionCodigo = mo_DoAtencionDatosAdicionales.SisCodigo
    End If
    lcFiliacionEnAdmision = mo_DoAtencionDatosAdicionales.sisAfiliacion
    If lcAfiliacionCodigo <> "" Then
        lcSql = " where idSiaSis=" & lcAfiliacionIdSiaSis & " and codigo='" & lcAfiliacionCodigo & "'"
        Set oRsAfiliadosSIS = mo_ReglasSISgalenhos.SisFiltraPacientesAfiliados(lcSql, wxParametroJAMO)
        If oRsAfiliadosSIS.RecordCount > 0 Then
           BuscarNumeroAfiliacion = oRsAfiliadosSIS.Fields!cDisa & " " & oRsAfiliadosSIS.Fields!cFormato & " " & oRsAfiliadosSIS.Fields!cnumero
        ElseIf lcFiliacionEnAdmision <> "" Then
           BuscarNumeroAfiliacion = lcFiliacionEnAdmision
        End If
    Else
        BuscarNumeroAfiliacion = lcFiliacionEnAdmision
    End If
    oConexion.Close
    Set oRsTmp1 = Nothing
    Set oRsAfiliadosSIS = Nothing
    Set oConexion = Nothing
    Set mo_DoAtencionDatosAdicionales = Nothing
End Function

'Frank 21082015 Nota Credito
Sub ImpresionNotaCredito(lcSerieNota As String, lcNumeroNota As String, lcRazonSocial As String, lcRucPersona As String, _
                       lcDireccionPersona As String, lcFechaEmision As String, lcTipoDocumento As String, lcSerieComprobante As String, _
                       lcNumeroComprobante As String, lcFechaComprobante As String, lcConceptoNota As String, _
                       lcSubTotal As String, lcTotal As String, lcMotivo As String, lbEsTicket As Boolean)
                       
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Dim lcNombreHospital As String
        Dim lcRuc As String, lcDireccion As String, lcTlefono As String
        Dim lcDevuelveFechaServidor As New SIGHDatos.Parametros
        Dim mrs_Tmp As New Recordset

        With mrs_Tmp
            .Fields.Append "Concepto", adVarChar, 20, adFldIsNullable
            .Fields.Append "SubTotal", adVarChar, 100, adFldIsNullable
            .LockType = adLockOptimistic
            .Open
        End With
        
        lcNombreHospital = lcBuscaParametro.SeleccionaFilaParametro(205)
        lcRuc = lcBuscaParametro.SeleccionaFilaParametro(339)
        lcDireccion = lcBuscaParametro.SeleccionaFilaParametro(206)
        lcTlefono = lcBuscaParametro.SeleccionaFilaParametro(207)

        On Error Resume Next
        '**************************************
        If lbEsTicket = True Then
             Set NotaCredTicket.DataSource = mrs_Tmp
             NotaCredTicket.Sections("Cabecera").Controls("lblEESS").Caption = UCase(lcNombreHospital)
             NotaCredTicket.Sections("Cabecera").Controls("lblEESSRuc").Caption = Trim(lcRuc)
             NotaCredTicket.Sections("Cabecera").Controls("lblEESSdireccion").Caption = UCase(lcDireccion)
             NotaCredTicket.Sections("Cabecera").Controls("lblEESStelefono").Caption = UCase(lcTlefono)
    
             NotaCredTicket.Sections("Cabecera").Controls("lblNroNota").Caption = UCase(lcSerieNota + "-" + lcNumeroNota)
             NotaCredTicket.Sections("Cabecera").Controls("lblRazonSocial").Caption = UCase(lcRazonSocial)
             NotaCredTicket.Sections("Cabecera").Controls("lblRuc").Caption = UCase(lcRucPersona)
             NotaCredTicket.Sections("Cabecera").Controls("lblDireccion").Caption = UCase(lcDireccionPersona)
             NotaCredTicket.Sections("Cabecera").Controls("lblFechaEmision").Caption = Left(lcFechaEmision, 10)
             NotaCredTicket.Sections("Cabecera").Controls("lblTipoDocumento").Caption = UCase(lcTipoDocumento)
             NotaCredTicket.Sections("Cabecera").Controls("lblNroDocumentoModifica").Caption = UCase(lcSerieComprobante + "-" + lcNumeroComprobante)
             NotaCredTicket.Sections("Cabecera").Controls("lblFechaComprobante").Caption = Left(lcFechaComprobante, 10)
             NotaCredTicket.Sections("Cabecera").Controls("lblConceptoNota").Caption = UCase(lcConceptoNota)
             NotaCredTicket.Sections("Cabecera").Controls("lblSubTotal").Caption = UCase(lcSubTotal)
             NotaCredTicket.Sections("Cabecera").Controls("lblIGV").Caption = "--"
             NotaCredTicket.Sections("Cabecera").Controls("lblOtros").Caption = "--"
             NotaCredTicket.Sections("Cabecera").Controls("lblTotal").Caption = UCase(lcTotal)
             NotaCredTicket.Sections("Cabecera").Controls("lblMotivo").Caption = UCase(lcMotivo)
            
             NotaCredTicket.RightMargin = 0
             NotaCredTicket.TopMargin = 0
             NotaCredTicket.LeftMargin = 0
             NotaCredTicket.BottomMargin = 0
             
             NotaCredTicket.Orientation = rptOrientDefault
    '         If lcParametro216 <> "1" Then
                NotaCredTicket.Show 1
    '         Else
    '            NotaCredTicket.PrintReport
    '         End If
        
        Else
             Set NotaCredito.DataSource = mrs_Tmp
             NotaCredito.Sections("Cabecera").Controls("lblEESS").Caption = UCase(lcNombreHospital)
             NotaCredito.Sections("Cabecera").Controls("lblEESSRuc").Caption = Trim(lcRuc)
             NotaCredito.Sections("Cabecera").Controls("lblEESSdireccion").Caption = UCase(lcDireccion)
             NotaCredito.Sections("Cabecera").Controls("lblEESStelefono").Caption = UCase(lcTlefono)
    
             NotaCredito.Sections("Cabecera").Controls("lblNroNota").Caption = UCase(lcSerieNota + "-" + lcNumeroNota)
             NotaCredito.Sections("Cabecera").Controls("lblRazonSocial").Caption = UCase(lcRazonSocial)
             NotaCredito.Sections("Cabecera").Controls("lblRuc").Caption = UCase(lcRucPersona)
             NotaCredito.Sections("Cabecera").Controls("lblDireccion").Caption = UCase(lcDireccionPersona)
             NotaCredito.Sections("Cabecera").Controls("lblFechaEmision").Caption = UCase(lcFechaEmision)
             NotaCredito.Sections("Cabecera").Controls("lblTipoDocumento").Caption = UCase(lcTipoDocumento)
             NotaCredito.Sections("Cabecera").Controls("lblNroDocumentoModifica").Caption = UCase(lcSerieComprobante + "-" + lcNumeroComprobante)
             NotaCredito.Sections("Cabecera").Controls("lblFechaComprobante").Caption = lcFechaComprobante
             NotaCredito.Sections("Cabecera").Controls("lblConceptoNota").Caption = UCase(lcConceptoNota)
             NotaCredito.Sections("Cabecera").Controls("lblSubTotal").Caption = UCase(lcSubTotal)
             NotaCredito.Sections("Cabecera").Controls("lblIGV").Caption = "--"
             NotaCredito.Sections("Cabecera").Controls("lblOtros").Caption = "--"
             NotaCredito.Sections("Cabecera").Controls("lblTotal").Caption = UCase(lcTotal)
             NotaCredito.Sections("Cabecera").Controls("lblMotivo").Caption = UCase(lcMotivo)
            
    '         NotaCredito.RightMargin = 100
    '         NotaCredito.TopMargin = 100
    '         NotaCredito.LeftMargin = 100
    '         NotaCredito.BottomMargin = 100
             
             NotaCredito.Orientation = rptOrientDefault
    '         If lcParametro216 <> "1" Then
                NotaCredito.Show 1
    '         Else
    '            NotaCredito.PrintReport
    '         End If
        End If
        Set mrs_Tmp = Nothing
End Sub



 Sub ImprimeCitas(oDoSiCitas As DoSiCitas, lnArea As sghAreasLaboraEmpleado, lcHistoria As String, _
        lcFinanciamiento As String, lcResponsable As String)
        Dim rsreporte As New Recordset
        Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
        Dim oConexion As New Connection
        Dim oReglasCaja As New SIGHNegocios.ReglasCaja
        Dim lcNombreHospital As String, lcRuc As String, lcDireccion As String, lcTlefono As String
        Dim lcCitas As String, lcCupos As String, lcUsuario As String
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        lcCitas = "": lcCupos = ""
        Set rsreporte = mo_ReglasImagenes.siCitasSeleccionarPorllaveTicket(oDoSiCitas.llaveTicket)
        If lnArea <> sghImageneología Then
            lcCitas = Trim(Str(rsreporte!idCitaSI))
            lcCupos = Trim(Str(rsreporte!Cupo))
        Else
            rsreporte.MoveFirst
            Do While Not rsreporte.EOF
               lcCitas = lcCitas & Trim(Str(rsreporte!idCitaSI)) & " - "
               If Not IsNull(rsreporte!Cupo) Then
                  lcCupos = lcCupos & Trim(Str(rsreporte!Cupo)) & " - "
               End If
               rsreporte.MoveNext
            Loop
            rsreporte.MoveFirst
        End If
        rsreporte.Close
        Set rsreporte = mo_ReglasImagenes.SiCitasDetalleSeleccionarPorLlaveTicket(oDoSiCitas.llaveTicket)

        Set SiPreCita.DataSource = rsreporte
        
        lcNombreHospital = lcBuscaParametro.SeleccionaFilaParametro(205)
        lcRuc = lcBuscaParametro.SeleccionaFilaParametro(339)
        lcDireccion = lcBuscaParametro.SeleccionaFilaParametro(206)
        lcTlefono = lcBuscaParametro.SeleccionaFilaParametro(207)
        
        SiPreCita.Sections("cabecera").Controls("etTitulo").Caption = "TICKET PARA " & IIf(lnArea = sghImageneología, "IMAGENES", "LABORATORIO")
        SiPreCita.Sections("cabecera").Controls("lcHospital1").Caption = UCase(lcNombreHospital)
        SiPreCita.Sections("cabecera").Controls("DatosEESS").Caption = UCase("RUC: " & lcRuc) & vbCrLf & UCase(lcDireccion) & vbCrLf & "Telef: " & UCase(lcTlefono)
        SiPreCita.Sections("cabecera").Controls("lblCupos").Caption = "N° Cupo: " & lcCupos
        SiPreCita.Sections("cabecera").Controls("lblFechaHora").Caption = "Fecha: " & Format(oDoSiCitas.fecha, sighEntidades.DevuelveFechaSoloFormato_DMY) & _
                                                                              "   Hora: " & oDoSiCitas.HoraInicio
        SiPreCita.Sections("cabecera").Controls("lblPuntoCarga").Caption = rsreporte!PuntoCarga
        SiPreCita.Sections("cabecera").Controls("etHistoria").Caption = "N° Historia: " & lcHistoria
        SiPreCita.Sections("cabecera").Controls("lblPaciente").Caption = "Paciente: " & oDoSiCitas.Paciente
        SiPreCita.Sections("cabecera").Controls("etCitas").Caption = "Cita N° " & lcCitas
        SiPreCita.Sections("cabecera").Controls("etTF").Caption = lcFinanciamiento
        SiPreCita.Sections("cabecera").Controls("etResponsable").Caption = "Resp: " & Left(lcResponsable, InStr(lcResponsable, "(") - 2)
        
        lcUsuario = oReglasCaja.SeleccionaDatosCajeroConexion(sighEntidades.Usuario, sghIniciales, oConexion)
        
        SiPreCita.Sections("pie").Controls("etUsuario").Caption = "Usuario: " & lcUsuario & "  F.Imp: " & _
                                                                   Format(Now, sighEntidades.DevuelveFechaSoloFormato_DMY_HMS)
        SiPreCita.Sections("pie").Controls("etAdvert").Caption = lcBuscaParametro.SeleccionaFilaParametro(346)
        
        
        SiPreCita.RightMargin = 0  'getReportRightMargin
        SiPreCita.TopMargin = 0 'getReportTopMargin
        SiPreCita.LeftMargin = 0 'getReportLeftMargin
        SiPreCita.BottomMargin = 0 ' getReportBottomMargin
        SiPreCita.Orientation = rptOrientDefault
        SiPreCita.Show 1
        Set rsreporte = Nothing
        Set mo_ReglasImagenes = Nothing
        Set oConexion = Nothing
        Set oReglasCaja = Nothing
End Sub


