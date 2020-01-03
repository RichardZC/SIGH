VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form rCrystal 
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
   Icon            =   "rCrystal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CrvReportes 
      Height          =   5595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      lastProp        =   500
      _cx             =   5080
      _cy             =   5080
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "rCrystal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Procesos y Vista previa de varios Reportes
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit

'aqui declara los objetos que contendra al rporte
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Private mflgContinuar As Boolean

Dim lc_TextoDelFiltro As String
Dim lc_TipoReporte As String
Dim lnIdPuntoCarga As Long
Dim lnOrdenadoPor As Long
Dim mrs_Tmp As New Recordset
Dim mrs_Tmp1 As New Recordset
Dim mda_FechaInicio As Date
Dim mda_FechaFin As Date
Dim ml_HoraInicio As String
Dim ml_HoraFin As String
Dim mb_ConsiderarSinMovimientos As Boolean
Dim mb_SeMuestraLotes As Boolean
Dim mb_StockMinimoMayorAcantidad As Boolean
Dim ml_idUsuario As Long
Dim ml_idProducto  As Long
Dim lnIdPuntoCargaOrigen As Long
Dim lnIdPuntoCargaDestino As Long
Dim ml_IdConcepto As Long
Dim ml_MovTipo As String
Dim ml_IdEstado  As Long
Dim lc_AlmacenesParaICI As String
Dim ml_IdAnio As Long
Dim ml_IdCuenta As Long
Dim ml_Dias  As Long
Dim ml_Almacen As String
Dim ml_Documento As String
Dim ml_AlmacenO As String
Dim ml_Importe As Double
Dim mb_ConsiderarRecalculo As Boolean
Dim mb_EnArchivoExcel As Boolean
Dim lnIdResponsable As Long
Dim mb_EnResumen As Boolean
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lcArchivoExcel As String
Dim lcTitEESS As String, lcTitDireccion As String, lcTitTelefono As String

Property Let EnResumen(lValue As Boolean)
    mb_EnResumen = lValue
End Property

Property Let EnArchivoExcel(lValue As Boolean)
    mb_EnArchivoExcel = lValue
End Property
Property Let ConsiderarRecalculo(lValue As Boolean)
    mb_ConsiderarRecalculo = lValue
End Property

Property Let Importe(lValue As Double)
    ml_Importe = lValue
End Property
Property Let AlmacenO(lValue As String)
    ml_AlmacenO = lValue
End Property
Property Let Documento(lValue As String)
    ml_Documento = lValue
End Property
Property Let Almacen(lValue As String)
    ml_Almacen = lValue
End Property

Property Let Dias(lValue As Long)
    ml_Dias = lValue
End Property

Property Let IdCuenta(lValue As Long)
    ml_IdCuenta = lValue
End Property
Property Let IdAnio(lValue As Long)
    ml_IdAnio = lValue
End Property

Property Let AlmacenesParaICI(lValue As String)
    lc_AlmacenesParaICI = lValue
End Property

Property Let Estado(lValue As Long)
    ml_IdEstado = lValue
End Property

Property Let Concepto(lValue As Long)
    ml_IdConcepto = lValue
End Property
Property Let MovTipo(lValue As String)
    ml_MovTipo = lValue
End Property


Property Let IdPuntoCargaDestino(iValue As Long)
   lnIdPuntoCargaDestino = iValue
End Property
Property Let IdPuntoCargaOrigen(iValue As Long)
   lnIdPuntoCargaOrigen = iValue
End Property

Property Let idProducto(lValue As Long)
    ml_idProducto = lValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Let StockMinimoMayorAcantidad(lValue As Boolean)
    mb_StockMinimoMayorAcantidad = lValue
End Property
Property Let SeMuestraLotes(lValue As Boolean)
    mb_SeMuestraLotes = lValue
End Property
Property Let ConsiderarSinMovimientos(lValue As Boolean)
    mb_ConsiderarSinMovimientos = lValue
End Property

Property Let HoraInicio(lValue As String)
    ml_HoraInicio = lValue
End Property
Property Let HoraFin(lValue As String)
    ml_HoraFin = lValue
End Property
Property Let FechaInicio(daValue As Date)
    mda_FechaInicio = daValue
End Property
Property Let FechaFin(daValue As Date)
    mda_FechaFin = daValue
End Property

Property Let IdPuntoCarga(iValue As Long)
   lnIdPuntoCarga = iValue
End Property
Property Let OrdenadoPor(iValue As Long)
   lnOrdenadoPor = iValue
End Property

Property Let TipoReporte(iValue As String)
   lc_TipoReporte = iValue
End Property
Property Let TextoDelFiltro(iValue As String)
   lc_TextoDelFiltro = iValue
End Property
Property Let IdResponsable(iValue As Long)
   lnIdResponsable = iValue
End Property

Private Sub Form_Load()
    If Len(lc_TextoDelFiltro) > 250 Then
       lc_TextoDelFiltro = Left(lc_TextoDelFiltro, 250)
    End If

    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crParamDef As CRAXDRT.ParameterFieldDefinition
    Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
    Dim mo_reglasCaja As New SIGHNegocios.ReglasCaja
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_AdminServComunes As New SIGHNegocios.ReglasComunes
    Dim rsReporte As New ADODB.Recordset
    Dim rsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim lnSaldoInicial As Long: Dim lnSaldofinal As Long
    Dim lnIngresos As Long: Dim lnSalidas As Long: Dim lnSalidasImg  As Long
    Dim ldFechaPrincipio As Date
    Dim lcCodigo As String: Dim lcNombre As String: Dim lnIdProducto As Long
    Dim lnPrecio As Double
    Dim oConexion As New ADODB.Connection
    Dim lbPrimeraVez As Boolean
    Dim LcTexto1 As String: Dim LcTexto2 As String: Dim lcTexto3 As String
    Dim lnId1 As Long: Dim lnId2 As Long: Dim lnId3 As Long
    Dim oDoImagMovimientoIngresos As New DoImagMovimientoIngresos
    Dim oDoImagMovimiento As New DoImagMovimiento
    Dim lcOpcs As String
    Dim lnIdMovimientoError As Long
    Dim lnEmergencia As Long    'franklin 2017
    Dim lnImpCE As Double, lnImpEmer As Double, lnImpHosp As Double, lnImpExt As Double
    On Error GoTo ErrHandler
    lnIdMovimientoError = 0
    Screen.MousePointer = vbHourglass
    lcTitEESS = lcBuscaParametro.SeleccionaFilaParametro(205)
    lcTitDireccion = lcBuscaParametro.SeleccionaFilaParametro(206)
    lcTitTelefono = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
    
    mflgContinuar = False
    Select Case lc_TipoReporte
    Case "rMovimientoDiario"
            'Proceso
            ldFechaPrincipio = Format("01/01/1990  00:01", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
            Set rsReporte = mo_ReglasImagenes.ImagMovimientoDetalleSeleccionarPorFechasYpuntoCarga(lnIdPuntoCarga, CDate("01/01/1990 00:01"), mda_FechaFin, 0)
            If rsReporte.RecordCount > 0 Then
                GenerarRecordsetTemporal lc_TipoReporte
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                    lnIdProducto = rsReporte.Fields!idProducto
                    lcCodigo = rsReporte.Fields!codigo
                    lcNombre = rsReporte.Fields!nombre
                    '*******Saldo Inicial********
                    lnSaldoInicial = 0
                    Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And rsReporte.Fields!fecha < mda_FechaInicio
                       If rsReporte.Fields!MovTipo = "S" Then
                           lnSaldoInicial = lnSaldoInicial - rsReporte.Fields!Cantidad
                       Else
                           lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!Cantidad
                       End If
                       rsReporte.MoveNext
                       If rsReporte.EOF Then
                          Exit Do
                       End If
                    Loop
                    '****** Movimientos en el Rango de Fechas***********
                    lnSalidas = 0: lnSalidasImg = 0: lnIngresos = 0
                    If Not rsReporte.EOF Then
                        Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And rsReporte.Fields!fecha <= mda_FechaFin
                            Select Case rsReporte.Fields!IdTipoConcepto
                            Case 1
                               lnIngresos = lnIngresos + rsReporte.Fields!Cantidad
                            Case 2
                               lnSalidas = lnSalidas + rsReporte.Fields!Cantidad
                            Case 3
                               lnSalidasImg = lnSalidasImg + rsReporte.Fields!Cantidad
                            End Select
                            rsReporte.MoveNext
                            If rsReporte.EOF Then
                               Exit Do
                            End If
                        Loop
                    End If
                    lbPrimeraVez = True
                    If mb_ConsiderarSinMovimientos = False Then
                       If (lnIngresos + lnSalidas + lnSalidasImg) = 0 Then
                          lbPrimeraVez = False
                       End If
                    End If
                    If lbPrimeraVez Then
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!codigo = lcCodigo
                        mrs_Tmp.Fields!nombre = lcNombre
                        mrs_Tmp.Fields!saldoI = lnSaldoInicial
                        mrs_Tmp.Fields!Ingresos = lnIngresos
                        mrs_Tmp.Fields!Salidas = lnSalidas
                        mrs_Tmp.Fields!SalidasImg = lnSalidasImg
                        mrs_Tmp.Fields!TotSalidas = lnSalidas + lnSalidasImg
                        mrs_Tmp.Fields!saldoF = lnSaldoInicial + lnIngresos - (lnSalidas + lnSalidasImg)
                        mrs_Tmp.Update
                    End If
                Loop
                'Reporte
                mflgContinuar = True
                Set crReport = crApp.OpenReport(App.Path & "\plantillas\ImagMovimientoDiario.rpt", 1)
                ' Parametros del reporte
                Set crParamDefs = crReport.ParameterFields
                For Each crParamDef In crParamDefs
                    Select Case crParamDef.ParameterFieldName
                        Case "Orden"
                            crParamDef.AddCurrentValue (lnOrdenadoPor)
                        Case "subTitulo"
                            crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                        Case "lcEESS"
                            crParamDef.AddCurrentValue (lcTitEESS)
                        Case "lcEESSdireccion"
                            crParamDef.AddCurrentValue (lcTitDireccion)
                        Case "lcEESStelefono"
                            crParamDef.AddCurrentValue (lcTitTelefono)
                    End Select
                Next
                crReport.Database.SetDataSource mrs_Tmp
            End If
    Case "rKardex"
            'Proceso
            ldFechaPrincipio = Format("01/01/1990  00:01", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
            Set rsReporte = mo_ReglasImagenes.ImagMovimientoDetalleSeleccionarPorFechasYpuntoCarga(lnIdPuntoCarga, CDate("01/01/1990 00:01"), mda_FechaFin, ml_idProducto)
            If rsReporte.RecordCount > 0 Then
                GenerarRecordsetTemporal lc_TipoReporte
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                    lnIdProducto = rsReporte.Fields!idProducto
                    lcCodigo = rsReporte.Fields!codigo
                    lcNombre = rsReporte.Fields!nombre
                    '*******Saldo Inicial********
                    lnSaldoInicial = 0
                    Do While Not rsReporte.EOF And rsReporte.Fields!fecha < mda_FechaInicio
                       If rsReporte.Fields!MovTipo = "S" Then
                           lnSaldoInicial = lnSaldoInicial - rsReporte.Fields!Cantidad
                       Else
                           lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!Cantidad
                       End If
                       rsReporte.MoveNext
                       If rsReporte.EOF Then
                          Exit Do
                       End If
                    Loop
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!fecha = " "
                    mrs_Tmp.Fields!MovTipo = " "
                    mrs_Tmp.Fields!idMovimiento = 0
                    mrs_Tmp.Fields!Ingresos = lnSaldoInicial
                    mrs_Tmp.Fields!Salidas = 0
                    mrs_Tmp.Fields!Saldo = lnSaldoInicial
                    mrs_Tmp.Fields!Concepto = "Saldo Inicial"
                    mrs_Tmp.Update
                    '****** Movimientos en el Rango de Fechas***********
                    If Not rsReporte.EOF Then
                        Do While Not rsReporte.EOF And rsReporte.Fields!fecha <= mda_FechaFin
                            lnSalidas = 0:  lnIngresos = 0
                            Select Case rsReporte.Fields!IdTipoConcepto
                            Case 1
                               lnIngresos = rsReporte.Fields!Cantidad
                               Set rsTmp = mo_ReglasImagenes.ImagMovimientoIngresosSeleccionarPorIdMovimiento(rsReporte.Fields!idMovimiento)
                               LcTexto1 = rsTmp.Fields!NroDocumento
                            Case Else
                               lnSalidas = rsReporte.Fields!Cantidad
                               If rsReporte.Fields!IdTipoConcepto = 3 Then
                                  Set rsTmp = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorIdMovimiento(rsReporte.Fields!idMovimiento)
                                  LcTexto1 = Trim(rsTmp.Fields!NroHistoriaClinica) & " " & Trim(rsTmp.Fields!ApellidoPaterno) & " " & Trim(rsTmp.Fields!ApellidoMaterno) & " " & rsTmp.Fields!PrimerNombre
                                  If rsReporte.Fields!cantidadFallada > 0 Then
                                     LcTexto1 = Trim(LcTexto1) & "     Cant.Fallada: " & Trim(Str(rsReporte.Fields!cantidadFallada))
                                  End If
                               Else
                                  Set rsTmp = mo_ReglasImagenes.ImagMovimientoSalidasSeleccionarPorIdMovimiento(rsReporte.Fields!idMovimiento)
                                  LcTexto1 = rsTmp.Fields!motivo
                               End If
                            End Select
                            lnSaldoInicial = lnSaldoInicial + lnIngresos - lnSalidas
                            mrs_Tmp.AddNew
                            mrs_Tmp.Fields!fecha = Format(rsReporte.Fields!fecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
                            mrs_Tmp.Fields!MovTipo = rsReporte.Fields!MovTipo
                            mrs_Tmp.Fields!idMovimiento = rsReporte.Fields!idMovimiento
                            mrs_Tmp.Fields!Ingresos = lnIngresos
                            mrs_Tmp.Fields!Salidas = lnSalidas
                            mrs_Tmp.Fields!Saldo = lnSaldoInicial
                            mrs_Tmp.Fields!Concepto = rsReporte.Fields!Concepto
                            mrs_Tmp.Fields!Observacion = LcTexto1
                            mrs_Tmp.Update
                            rsReporte.MoveNext
                            If rsReporte.EOF Then
                               Exit Do
                            End If
                        Loop
                    End If
                Loop
                'Reporte
                mflgContinuar = True
                Set crReport = crApp.OpenReport(App.Path & "\plantillas\ImagKardex.rpt", 1)
                ' Parametros del reporte
                Set crParamDefs = crReport.ParameterFields
                For Each crParamDef In crParamDefs
                    Select Case crParamDef.ParameterFieldName
                        Case "Orden"
                            crParamDef.AddCurrentValue (lnOrdenadoPor)
                        Case "subTitulo"
                            crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                        Case "lcEESS"
                            crParamDef.AddCurrentValue (lcTitEESS)
                        Case "lcEESSdireccion"
                            crParamDef.AddCurrentValue (lcTitDireccion)
                        Case "lcEESStelefono"
                            crParamDef.AddCurrentValue (lcTitTelefono)
                    End Select
                Next
                crReport.Database.SetDataSource mrs_Tmp
            End If
    Case "rProduccion"
            'Proceso
            If mb_EnResumen = True Then
               'A nivel de CPT
                Set rsReporte = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorFechasPuntoCarga(lnIdPuntoCarga, mda_FechaInicio, mda_FechaFin, sghPorIdProductoMasFecha)
                If rsReporte.RecordCount > 0 Then
                    If lnIdResponsable > 0 Then
                       rsReporte.Filter = "idPersonaTomaImagen=" & lnIdResponsable
                    End If
                    If rsReporte.RecordCount > 0 Then
                        GenerarRecordsetTemporal lc_TipoReporte
                        rsReporte.MoveFirst
                        Do While Not rsReporte.EOF
                            lcOpcs = mo_AdminServComunes.OPCsDevuelveCodigoOPCporCodigoCPT(rsReporte.Fields!codigo)
                            lcOpcs = IIf(lcOpcs = "", "", "  (" & lcOpcs & ")")
                        
                            lnIdProducto = rsReporte.Fields!idProductoCpt
                            lcCodigo = rsReporte.Fields!codigo
                            lcNombre = rsReporte.Fields!nombreMINSA
                            '*******Saldo Inicial********
                            lnSalidas = 0: lnPrecio = 0
                            lnSalidasImg = 0: lnSaldoInicial = 0: lnSaldofinal = 0: lnEmergencia = 0
                            lnImpCE = 0: lnImpEmer = 0: lnImpHosp = 0: lnImpExt = 0
                            Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProductoCpt
                               lcTexto3 = IIf(IsNull(rsReporte!nombrePlan), "Externo", rsReporte!nombrePlan)
                               lnPrecio = lnPrecio + rsReporte.Fields!total
                               lnSalidas = lnSalidas + rsReporte.Fields!Cantidad
                               Select Case rsReporte.Fields!idTipoServicio
                               Case 1  'ce
                                  lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!Cantidad
                                  lnImpCE = lnImpCE + rsReporte!total
                               Case 2  'emergencia
                                  lnEmergencia = lnEmergencia + rsReporte.Fields!Cantidad
                                  lnImpEmer = lnImpEmer + rsReporte!total
                               Case 3  'Hospitalizacion
                                  lnSaldofinal = lnSaldofinal + rsReporte.Fields!Cantidad
                                  lnImpHosp = lnImpHosp + rsReporte!total
                               Case Else
                                  lnSalidasImg = lnSalidasImg + rsReporte.Fields!Cantidad
                                  lnImpExt = lnImpExt + rsReporte!total
                               End Select
                               rsReporte.MoveNext
                               If rsReporte.EOF Then
                                  Exit Do
                               End If
                            Loop
                            mrs_Tmp.AddNew
                            mrs_Tmp.Fields!codigo = Trim(lcCodigo) & lcOpcs
                            mrs_Tmp.Fields!nombre = lcNombre
                            mrs_Tmp.Fields!buenos = lnSalidasImg      'Externos
                            mrs_Tmp.Fields!fallados = lnSaldoInicial  'CE
                            mrs_Tmp.Fields!repetidos = lnSaldofinal   'Hosp
                            mrs_Tmp.Fields!total = lnSalidas
                            mrs_Tmp.Fields!Importe = lnPrecio
                            mrs_Tmp.Fields!emergencia = lnEmergencia   'emergencia
                            mrs_Tmp.Fields!ImpCE = lnImpCE
                            mrs_Tmp.Fields!ImpEmer = lnImpEmer
                            mrs_Tmp.Fields!ImpHosp = lnImpHosp
                            mrs_Tmp.Fields!ImpExt = lnImpExt
                            mrs_Tmp.Fields!Financ = lcTexto3
                            mrs_Tmp.Update
                        Loop
                        'Reporte
                        mflgContinuar = True
                        Set crReport = crApp.OpenReport(App.Path & "\plantillas\ImagProductividadCPT.rpt", 1)
                        ' Parametros del reporte
                        Set crParamDefs = crReport.ParameterFields
                        For Each crParamDef In crParamDefs
                            Select Case crParamDef.ParameterFieldName
                                Case "Orden"
                                    crParamDef.AddCurrentValue (lnOrdenadoPor)
                                Case "subTitulo"
                                    crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                                Case "lcEESS"
                                    crParamDef.AddCurrentValue (lcTitEESS)
                                Case "lcEESSdireccion"
                                    crParamDef.AddCurrentValue (lcTitDireccion)
                                Case "lcEESStelefono"
                                    crParamDef.AddCurrentValue (lcTitTelefono)
                            End Select
                        Next
                        crReport.Database.SetDataSource mrs_Tmp
                     End If
                End If
            Else
               'A nivel de INSUMOS
                Set rsReporte = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorFechasPToInsumos(lnIdPuntoCarga, mda_FechaInicio, mda_FechaFin, sghPorIdProductoMasFecha)
                If rsReporte.RecordCount > 0 Then
                    If lnIdResponsable > 0 Then
                       rsReporte.Filter = "idPersonaTomaImagen=" & lnIdResponsable
                    End If
                    If rsReporte.RecordCount > 0 Then
                        GenerarRecordsetTemporal lc_TipoReporte
                        rsReporte.MoveFirst
                        Do While Not rsReporte.EOF
                            lnIdProducto = rsReporte.Fields!idProducto
                            lcCodigo = rsReporte.Fields!codigo
                            lcNombre = rsReporte.Fields!nombre
                            '*******Saldo Inicial********
                            lnSalidas = 0: lnSalidasImg = 0
                            Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto
                               lnSalidas = lnSalidas + rsReporte.Fields!cantidadFallada
                               lnSalidasImg = lnSalidasImg + (rsReporte.Fields!Cantidad - rsReporte.Fields!cantidadFallada)
                               rsReporte.MoveNext
                               If rsReporte.EOF Then
                                  Exit Do
                               End If
                            Loop
                            mrs_Tmp.AddNew
                            mrs_Tmp.Fields!codigo = lcCodigo
                            mrs_Tmp.Fields!nombre = lcNombre
                            mrs_Tmp.Fields!buenos = lnSalidasImg
                            mrs_Tmp.Fields!fallados = lnSalidas
                            mrs_Tmp.Fields!repetidos = 0
                            mrs_Tmp.Fields!total = lnSalidas + lnSalidasImg
                            mrs_Tmp.Update
                        Loop
                        lnSalidasImg = mrs_Tmp.RecordCount
                        'Repeticion de Examen
                        Set rsReporte = mo_ReglasImagenes.ImagMovimientoSalidasSeleccionarProductosPorFechasYpuntoCarga(lnIdPuntoCarga, mda_FechaInicio, mda_FechaFin)
                        If rsReporte.RecordCount > 0 Then
                            If lnIdResponsable > 0 Then
                               rsReporte.Filter = "idResponsable=" & lnIdResponsable
                            End If
                            If rsReporte.RecordCount > 0 Then
                                rsReporte.MoveFirst
                                Do While Not rsReporte.EOF
                                    lnIdProducto = rsReporte.Fields!idProducto
                                    lcCodigo = rsReporte.Fields!codigo
                                    lcNombre = rsReporte.Fields!nombre
                                    lnSalidas = 0
                                    Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto
                                       If rsReporte.Fields!idMotivoSalida = 1 Then
                                          lnSalidas = lnSalidas + rsReporte.Fields!Cantidad
                                       End If
                                       rsReporte.MoveNext
                                       If rsReporte.EOF Then
                                          Exit Do
                                       End If
                                    Loop
                                    If lnSalidas > 0 Then
                                        lbPrimeraVez = True
                                        If lnSalidasImg > 0 Then
                                           mrs_Tmp.MoveFirst
                                           mrs_Tmp.Find "codigo='" & lcCodigo & "'"
                                           If Not mrs_Tmp.EOF Then
                                              lbPrimeraVez = False
                                           End If
                                        End If
                                        If lbPrimeraVez = True Then
                                            mrs_Tmp.AddNew
                                            mrs_Tmp.Fields!codigo = lcCodigo
                                            mrs_Tmp.Fields!nombre = lcNombre
                                            mrs_Tmp.Fields!buenos = 0
                                            mrs_Tmp.Fields!fallados = 0
                                            mrs_Tmp.Fields!repetidos = lnSalidas
                                            mrs_Tmp.Fields!total = lnSalidas
                                        Else
                                            lnIngresos = mrs_Tmp.Fields!buenos + mrs_Tmp.Fields!fallados + lnSalidas
                                            mrs_Tmp.Fields!repetidos = lnSalidas
                                            mrs_Tmp.Fields!total = lnIngresos
                                        End If
                                        mrs_Tmp.Update
                                    End If
                                Loop
                            End If
                        End If
                        
                        'Reporte
                        mflgContinuar = True
                        Set crReport = crApp.OpenReport(App.Path & "\plantillas\ImagProductividad.rpt", 1)
                        ' Parametros del reporte
                        Set crParamDefs = crReport.ParameterFields
                        For Each crParamDef In crParamDefs
                            Select Case crParamDef.ParameterFieldName
                                Case "Orden"
                                    crParamDef.AddCurrentValue (lnOrdenadoPor)
                                Case "subTitulo"
                                    crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                                Case "lcEESS"
                                    crParamDef.AddCurrentValue (lcTitEESS)
                                Case "lcEESSdireccion"
                                    crParamDef.AddCurrentValue (lcTitDireccion)
                                Case "lcEESStelefono"
                                    crParamDef.AddCurrentValue (lcTitTelefono)
                            End Select
                        Next
                        crReport.Database.SetDataSource mrs_Tmp
                    End If
                End If
            End If
    Case "rProduccionTF"
            'Proceso
            If mb_EnResumen = True Then
               'A nivel de CPT
                Set rsReporte = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorFechasPuntoCarga(lnIdPuntoCarga, mda_FechaInicio, mda_FechaFin, sghPorIdProductoMasFecha)
                If rsReporte.RecordCount > 0 Then
                    If lnIdResponsable > 0 Then
                       rsReporte.Filter = "idPersonaTomaImagen=" & lnIdResponsable
                    End If
                    If rsReporte.RecordCount > 0 Then
                        GenerarRecordsetTemporal "rProduccion"
                        rsReporte.MoveFirst
                        Do While Not rsReporte.EOF
                            lcOpcs = mo_AdminServComunes.OPCsDevuelveCodigoOPCporCodigoCPT(rsReporte.Fields!codigo)
                            lcOpcs = IIf(lcOpcs = "", "", "  (" & lcOpcs & ")")
                        
                            lnIdProducto = rsReporte.Fields!idProductoCpt
                            lcCodigo = rsReporte.Fields!codigo
                            lcNombre = rsReporte.Fields!nombreMINSA
                            '*******Saldo Inicial********
                            lnSalidasImg = 0: lnSaldoInicial = 0: lnSaldofinal = 0: lnEmergencia = 0
                            lnImpCE = 0: lnImpEmer = 0: lnImpHosp = 0: lnImpExt = 0
                            Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProductoCpt
                               lcTexto3 = rsReporte!nombrePlan
                               lnSalidas = 0: lnPrecio = 0
                               Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProductoCpt And lcTexto3 = rsReporte!nombrePlan
                                    lnPrecio = lnPrecio + rsReporte.Fields!total
                                    lnSalidas = lnSalidas + rsReporte.Fields!Cantidad
                                    rsReporte.MoveNext
                                    If rsReporte.EOF Then
                                       Exit Do
                                    End If
                                Loop
                                mrs_Tmp.AddNew
                                mrs_Tmp.Fields!codigo = Trim(lcCodigo) & lcOpcs
                                mrs_Tmp.Fields!nombre = Left(lcCodigo & " - " & lcNombre, 255)
                                mrs_Tmp.Fields!total = lnSalidas
                                mrs_Tmp.Fields!Importe = lnPrecio
                                mrs_Tmp.Fields!Financ = IIf(lcTexto3 = "", "Externo", lcTexto3)
                                mrs_Tmp.Update
                                If rsReporte.EOF Then
                                   Exit Do
                                End If
                            Loop
                        Loop
                        'Reporte
                        mflgContinuar = True
                        Set crReport = crApp.OpenReport(App.Path & "\plantillas\ImagProducCPT_FF.rpt", 1)
                        ' Parametros del reporte
                        Set crParamDefs = crReport.ParameterFields
                        For Each crParamDef In crParamDefs
                            Select Case crParamDef.ParameterFieldName
                                Case "Orden"
                                    crParamDef.AddCurrentValue (lnOrdenadoPor)
                                Case "subTitulo"
                                    crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                                Case "lcEESS"
                                    crParamDef.AddCurrentValue (lcTitEESS)
                                Case "lcEESSdireccion"
                                    crParamDef.AddCurrentValue (lcTitDireccion)
                                Case "lcEESStelefono"
                                    crParamDef.AddCurrentValue (lcTitTelefono)
                            End Select
                        Next
                        crReport.Database.SetDataSource mrs_Tmp
                     End If
                End If
            End If
    
    Case "rEcogGen"
            'Proceso
            Set rsReporte = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorFechasPuntoCarga(lnIdPuntoCarga, mda_FechaInicio, mda_FechaFin, sghPorFechaYhora)
            If rsReporte.RecordCount > 0 Then
                GenerarRecordsetTemporal lc_TipoReporte
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                    If mb_EnResumen = True Then
                        lbPrimeraVez = True
                        If mrs_Tmp.RecordCount > 0 Then
                           mrs_Tmp.MoveFirst
                           mrs_Tmp.Find "plan='" & rsReporte!nombrePlan & "'"
                           If Not mrs_Tmp.EOF Then
                              lbPrimeraVez = False
                           End If
                        End If
                        If lbPrimeraVez = True Then
                            mrs_Tmp.AddNew
                            mrs_Tmp.Fields!Plan = rsReporte!nombrePlan
                            mrs_Tmp.Fields!numero = 1
                            mrs_Tmp.Fields!total = rsReporte!total
                        Else
                            mrs_Tmp.Fields!numero = mrs_Tmp.Fields!numero + 1
                            mrs_Tmp.Fields!total = mrs_Tmp.Fields!total + rsReporte!total
                        End If
                        mrs_Tmp.Update
                    Else
                        lcOpcs = mo_AdminServComunes.OPCsDevuelveCodigoOPCporCodigoCPT(rsReporte.Fields!codigo)
                        lcOpcs = IIf(lcOpcs = "", "", "  (" & lcOpcs & ")")
                        lnIdMovimientoError = rsReporte!idMovimiento
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!fecha = Format(rsReporte!fecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
                        mrs_Tmp.Fields!fechaHr = Format(rsReporte.Fields!fecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
                        mrs_Tmp.Fields!idMovimiento = rsReporte!idMovimiento
                        mrs_Tmp.Fields!Paciente = rsReporte!NroHistoriaClinica & " - " & Trim(rsReporte!Paciente)
                        mrs_Tmp.Fields!sexo = IIf(rsReporte!idTipoSexoMov = 1, "M", "F")
                        If Not IsNull(rsReporte!FechaNacimientoMov) Then
                           mrs_Tmp.Fields!Edad = EdadActual(rsReporte!FechaNacimientoMov, rsReporte.Fields!fecha)
                        End If
                        mrs_Tmp.Fields!NroCuenta = rsReporte!idCuentaAtencion
                        mrs_Tmp.Fields!Procedencia = rsReporte!dServicio
                        mrs_Tmp.Fields!Plan = rsReporte!nombrePlan
                        mrs_Tmp.Fields!Recibo = rsReporte!NroSerie & " " & rsReporte!NroDocumento
                        mrs_Tmp.Fields!Resultado = rsReporte!ResultadoFinal
                        mrs_Tmp.Fields!codigo = Trim(rsReporte!codigo) & lcOpcs
                        mrs_Tmp.Fields!nombre = rsReporte!nombreMINSA
                        mrs_Tmp.Fields!total = rsReporte!total
                        mrs_Tmp.Update
                    End If
                    rsReporte.MoveNext
                Loop
                
                'Reporte
                mflgContinuar = True
                If mb_EnResumen = True Then
                   lc_TextoDelFiltro = "ECOGRAFIA GENERAL" & Chr(13) & Chr(13) & lc_TextoDelFiltro
                   Set crReport = crApp.OpenReport(App.Path & "\plantillas\ImagEcogGenResumen.rpt", 1)
                Else
                   Set crReport = crApp.OpenReport(App.Path & "\plantillas\ImagEcogGen.rpt", 1)
                End If
                ' Parametros del reporte
                Set crParamDefs = crReport.ParameterFields
                For Each crParamDef In crParamDefs
                    Select Case crParamDef.ParameterFieldName
                        Case "subTitulo"
                            crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                        Case "lcEESS"
                            crParamDef.AddCurrentValue (lcTitEESS)
                        Case "lcEESSdireccion"
                            crParamDef.AddCurrentValue (lcTitDireccion)
                        Case "lcEESStelefono"
                            crParamDef.AddCurrentValue (lcTitTelefono)
                    End Select
                Next
                crReport.Database.SetDataSource mrs_Tmp
            End If
    Case "rRayosX"
            'Proceso
            Set rsReporte = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorFechasPuntoCarga(lnIdPuntoCarga, mda_FechaInicio, mda_FechaFin, sghPorFechaYhora)
            If rsReporte.RecordCount > 0 Then
                GenerarRecordsetTemporal lc_TipoReporte
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                    If mb_EnResumen = True Then
                        lbPrimeraVez = True
                        If mrs_Tmp.RecordCount > 0 Then
                           mrs_Tmp.MoveFirst
                           mrs_Tmp.Find "plan='" & rsReporte!nombrePlan & "'"
                           If Not mrs_Tmp.EOF Then
                              lbPrimeraVez = False
                           End If
                        End If
                        If lbPrimeraVez = True Then
                            mrs_Tmp.AddNew
                            mrs_Tmp.Fields!Plan = rsReporte!nombrePlan
                            mrs_Tmp.Fields!numero = 1
                            mrs_Tmp.Fields!total = rsReporte!total
                        Else
                            mrs_Tmp.Fields!numero = mrs_Tmp.Fields!numero + 1
                            mrs_Tmp.Fields!total = mrs_Tmp.Fields!total + rsReporte!total
                        End If
                        mrs_Tmp.Update
                    Else
                        If IsNull(rsReporte.Fields!codigo) Then
                           lcOpcs = ""
                        Else
                           lcOpcs = mo_AdminServComunes.OPCsDevuelveCodigoOPCporCodigoCPT(rsReporte.Fields!codigo)
                        End If
                        lcOpcs = IIf(lcOpcs = "", "", "  (" & lcOpcs & ")")
                        lnIdMovimientoError = rsReporte!idMovimiento
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!fecha = Format(rsReporte!fecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
                        mrs_Tmp.Fields!fechaHr = Format(rsReporte.Fields!fecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
                        mrs_Tmp.Fields!idMovimiento = rsReporte!idMovimiento
                        mrs_Tmp.Fields!Paciente = rsReporte!NroHistoriaClinica & " - " & Trim(rsReporte!Paciente)
                        mrs_Tmp.Fields!sexo = IIf(rsReporte!idTipoSexoMov = 1, "M", "F")
                        If Not IsNull(rsReporte!FechaNacimientoMov) Then
                            mrs_Tmp.Fields!Edad = EdadActual(rsReporte!FechaNacimientoMov, rsReporte.Fields!fecha)
                        End If
                        mrs_Tmp.Fields!NroCuenta = rsReporte!idCuentaAtencion
                        mrs_Tmp.Fields!Procedencia = rsReporte!dServicio
                        mrs_Tmp.Fields!Plan = rsReporte!nombrePlan
                        mrs_Tmp.Fields!Recibo = rsReporte!NroSerie & " " & rsReporte!NroDocumento
                        mrs_Tmp.Fields!Resultado = rsReporte!ResultadoFinal
                        If Not IsNull(rsReporte!idPersonaRecoge) Then
                           mrs_Tmp.Fields!Recoje = mo_ReglasImagenes.ImagRecojeExamenSeleccionaPorID(rsReporte!idPersonaRecoge)
                        End If
                        mrs_Tmp.Fields!Responsable = mo_reglasCaja.SeleccionaDatosCajero(rsReporte.Fields!IdPersonaTomaImagen, sghIniciales)
                        mrs_Tmp.Fields!Zona = rsReporte!ZonaRayosX
                        mrs_Tmp.Fields!codigo = Trim(rsReporte!codigo) & lcOpcs
                        mrs_Tmp.Fields!nombre = rsReporte!nombreMINSA
                        mrs_Tmp.Fields!total = rsReporte!total
                        mrs_Tmp.Update
                    End If
                    rsReporte.MoveNext
                Loop
                
                'Reporte
                mflgContinuar = True
                If mb_EnResumen = True Then
                   lc_TextoDelFiltro = "RAYOS  X" & Chr(13) & Chr(13) & lc_TextoDelFiltro
                   Set crReport = crApp.OpenReport(App.Path & "\plantillas\ImagEcogGenResumen.rpt", 1)
                Else
                   Set crReport = crApp.OpenReport(App.Path & "\plantillas\ImagRayosX.rpt", 1)
                End If
                ' Parametros del reporte
                Set crParamDefs = crReport.ParameterFields
                For Each crParamDef In crParamDefs
                    Select Case crParamDef.ParameterFieldName
                        Case "subTitulo"
                            crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                        Case "lcEESS"
                            crParamDef.AddCurrentValue (lcTitEESS)
                        Case "lcEESSdireccion"
                            crParamDef.AddCurrentValue (lcTitDireccion)
                        Case "lcEESStelefono"
                            crParamDef.AddCurrentValue (lcTitTelefono)
                    End Select
                Next
                crReport.Database.SetDataSource mrs_Tmp
            End If
    Case "rTomografia"
            'Proceso
            Set rsReporte = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorFechasPuntoCarga(lnIdPuntoCarga, mda_FechaInicio, mda_FechaFin, sghPorFechaYhora)
            If rsReporte.RecordCount > 0 Then
                GenerarRecordsetTemporal lc_TipoReporte
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                    If mb_EnResumen = True Then
                        lbPrimeraVez = True
                        If mrs_Tmp.RecordCount > 0 Then
                           mrs_Tmp.MoveFirst
                           mrs_Tmp.Find "plan='" & rsReporte!nombrePlan & "'"
                           If Not mrs_Tmp.EOF Then
                              lbPrimeraVez = False
                           End If
                        End If
                        If lbPrimeraVez = True Then
                            mrs_Tmp.AddNew
                            mrs_Tmp.Fields!Plan = rsReporte!nombrePlan
                            mrs_Tmp.Fields!numero = 1
                            mrs_Tmp.Fields!total = rsReporte!total
                        Else
                            mrs_Tmp.Fields!numero = mrs_Tmp.Fields!numero + 1
                            mrs_Tmp.Fields!total = mrs_Tmp.Fields!total + rsReporte!total
                        End If
                        mrs_Tmp.Update
                    Else
                        lcOpcs = mo_AdminServComunes.OPCsDevuelveCodigoOPCporCodigoCPT(rsReporte.Fields!codigo)
                        lcOpcs = IIf(lcOpcs = "", "", "  (" & lcOpcs & ")")
                        lnIdMovimientoError = rsReporte!idMovimiento
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!fecha = Format(rsReporte!fecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
                        mrs_Tmp.Fields!fechaHr = Format(rsReporte.Fields!fecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
                        mrs_Tmp.Fields!idMovimiento = rsReporte!idMovimiento
                        mrs_Tmp.Fields!Paciente = rsReporte!NroHistoriaClinica & " - " & Trim(rsReporte!Paciente)
                        mrs_Tmp.Fields!sexo = IIf(rsReporte!idTipoSexoMov = 1, "M", "F")
                        If Not IsNull(rsReporte!FechaNacimientoMov) Then
                           mrs_Tmp.Fields!Edad = EdadActual(rsReporte!FechaNacimientoMov, rsReporte.Fields!fecha)
                        End If
                        mrs_Tmp.Fields!NroCuenta = rsReporte!idCuentaAtencion
                        mrs_Tmp.Fields!Procedencia = rsReporte!dServicio
                        mrs_Tmp.Fields!Plan = rsReporte!nombrePlan
                        mrs_Tmp.Fields!Recibo = rsReporte!NroSerie & " " & rsReporte!NroDocumento
                        mrs_Tmp.Fields!Resultado = rsReporte!ResultadoFinal
                        mrs_Tmp.Fields!EsContraste = IIf(rsReporte!EsContraste = 1, "Si", "")
                        mrs_Tmp.Fields!EsIonico = IIf(rsReporte!EsContrasteIonico = 1, "Si", "")
                        mrs_Tmp.Fields!codigo = Trim(rsReporte!codigo) & lcOpcs
                        mrs_Tmp.Fields!nombre = rsReporte!nombreMINSA
                        mrs_Tmp.Fields!Cantidad = rsReporte!Cantidad
                        mrs_Tmp.Fields!total = rsReporte!total
                        mrs_Tmp.Fields!totalIR = rsReporte!total + (rsReporte!total * rsReporte!PorcInformeRadiolog) / 100
                        mrs_Tmp.Update
                    End If
                    rsReporte.MoveNext
                Loop
                
                'Reporte
                mflgContinuar = True
                If mb_EnResumen = True Then
                   lc_TextoDelFiltro = "TOMOGRAFIA" & Chr(13) & Chr(13) & lc_TextoDelFiltro
                   Set crReport = crApp.OpenReport(App.Path & "\plantillas\ImagEcogGenResumen.rpt", 1)
                Else
                   Set crReport = crApp.OpenReport(App.Path & "\plantillas\ImagTomografia.rpt", 1)
                End If
                ' Parametros del reporte
                Set crParamDefs = crReport.ParameterFields
                For Each crParamDef In crParamDefs
                    Select Case crParamDef.ParameterFieldName
                        Case "subTitulo"
                            crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                        Case "lcEESS"
                            crParamDef.AddCurrentValue (lcTitEESS)
                        Case "lcEESSdireccion"
                            crParamDef.AddCurrentValue (lcTitDireccion)
                        Case "lcEESStelefono"
                            crParamDef.AddCurrentValue (lcTitTelefono)
                    End Select
                Next
                crReport.Database.SetDataSource mrs_Tmp
            End If
    Case "rEcogObs"
            oConexion.Open SIGHEntidades.CadenaConexion
            oConexion.CursorLocation = adUseClient
            'Proceso
            Set rsReporte = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorFechasPuntoCarga(lnIdPuntoCarga, mda_FechaInicio, mda_FechaFin, sghPorFechaYhora)
            If rsReporte.RecordCount > 0 Then
                GenerarRecordsetTemporal lc_TipoReporte
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                    If mb_EnResumen = True Then
                        lbPrimeraVez = True
                        If mrs_Tmp.RecordCount > 0 Then
                           mrs_Tmp.MoveFirst
                           mrs_Tmp.Find "plan='" & rsReporte!nombrePlan & "'"
                           If Not mrs_Tmp.EOF Then
                              lbPrimeraVez = False
                           End If
                        End If
                        If lbPrimeraVez = True Then
                            mrs_Tmp.AddNew
                            mrs_Tmp.Fields!Plan = rsReporte!nombrePlan
                            mrs_Tmp.Fields!numero = 1
                            mrs_Tmp.Fields!total = rsReporte!total
                        Else
                            mrs_Tmp.Fields!numero = mrs_Tmp.Fields!numero + 1
                            mrs_Tmp.Fields!total = mrs_Tmp.Fields!total + rsReporte!total
                        End If
                        mrs_Tmp.Update
                    Else
                        lcNombre = ""
                        If Not IsNull(rsReporte!idCuentaAtencion) Then
                           Set rsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(rsReporte!idCuentaAtencion, oConexion)
                           If rsTmp.RecordCount > 0 Then
                              lcNombre = IIf(IsNull(rsTmp.Fields!idTipoReferenciaOrigen), "", "Si")
                           End If
                        End If
                        lcOpcs = mo_AdminServComunes.OPCsDevuelveCodigoOPCporCodigoCPT(rsReporte.Fields!codigo)
                        lcOpcs = IIf(lcOpcs = "", "", "  (" & lcOpcs & ")")
                        lnIdMovimientoError = rsReporte!idMovimiento
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!fecha = Format(rsReporte!fecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
                        mrs_Tmp.Fields!fechaHr = Format(rsReporte.Fields!fecha, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
                        mrs_Tmp.Fields!idMovimiento = rsReporte!idMovimiento
                        mrs_Tmp.Fields!Paciente = rsReporte!NroHistoriaClinica & " - " & Trim(rsReporte!Paciente)
                        mrs_Tmp.Fields!sexo = IIf(rsReporte!idTipoSexoMov = 1, "M", "F")
                        If Not IsNull(rsReporte!FechaNacimientoMov) Then
                           mrs_Tmp.Fields!Edad = EdadActual(rsReporte!FechaNacimientoMov, rsReporte.Fields!fecha)
                        End If
                        mrs_Tmp.Fields!NroCuenta = rsReporte!idCuentaAtencion
                        mrs_Tmp.Fields!Procedencia = rsReporte!dServicio
                        mrs_Tmp.Fields!Plan = rsReporte!nombrePlan
                        mrs_Tmp.Fields!Recibo = rsReporte!NroSerie & " " & rsReporte!NroDocumento
                        mrs_Tmp.Fields!Resultado = rsReporte!ResultadoFinal
                        mrs_Tmp.Fields!Referido = lcNombre
                        mrs_Tmp.Fields!codigo = Trim(rsReporte!codigo) & lcOpcs
                        mrs_Tmp.Fields!nombre = rsReporte!nombreMINSA
                        mrs_Tmp.Fields!total = rsReporte!total
                        mrs_Tmp.Update
                    End If
                    rsReporte.MoveNext
                Loop
                
                'Reporte
                mflgContinuar = True
                If mb_EnResumen = True Then
                   lc_TextoDelFiltro = "ECOGRAFIA OBSTETRICA" & Chr(13) & Chr(13) & lc_TextoDelFiltro
                   Set crReport = crApp.OpenReport(App.Path & "\plantillas\ImagEcogGenResumen.rpt", 1)
                Else
                   Set crReport = crApp.OpenReport(App.Path & "\plantillas\ImagEcogObs.rpt", 1)
                End If
                ' Parametros del reporte
                Set crParamDefs = crReport.ParameterFields
                For Each crParamDef In crParamDefs
                    Select Case crParamDef.ParameterFieldName
                        Case "subTitulo"
                            crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                        Case "lcEESS"
                            crParamDef.AddCurrentValue (lcTitEESS)
                        Case "lcEESSdireccion"
                            crParamDef.AddCurrentValue (lcTitDireccion)
                        Case "lcEESStelefono"
                            crParamDef.AddCurrentValue (lcTitTelefono)
                    End Select
                Next
                crReport.Database.SetDataSource mrs_Tmp
            End If
            oConexion.Close
    Case "rAuditoriaImg"
        Set mrs_Tmp1 = mo_ReglasFarmacia.AuditoriaPorFechasUsuario(ml_idUsuario, mda_FechaInicio, mda_FechaFin)
        mrs_Tmp1.Filter = "idListItem>=1315 and idListItem<=1320"
        If mrs_Tmp1.RecordCount = 0 Then
             mflgContinuar = False
        Else
             GenerarRecordsetTemporal lc_TipoReporte
             mrs_Tmp1.MoveFirst
             Do While Not mrs_Tmp1.EOF
                If UCase(Left(mrs_Tmp1.Fields!tabla, 14)) = "IMAGMOVIMIENTO" Then
                        lcTexto3 = IIf(mrs_Tmp1.Fields!accion = "A", "Agregó", IIf(mrs_Tmp1.Fields!accion = "M", "Modificó", "Anuló"))
                        LcTexto1 = Mid(mrs_Tmp1.Fields!tabla, 16, 1)  'MOVTIPO
                        LcTexto2 = ""
                        lnSalidasImg = Val(Mid(mrs_Tmp1.Fields!tabla, 18, 15))   'NumeroMovimiento
                        Set oDoImagMovimiento = mo_ReglasImagenes.ImagMovimientoSeleccionarPorId(lnSalidasImg)
                        If lnIdPuntoCarga = oDoImagMovimiento.IdPuntoCarga Then
                            If LcTexto1 = "E" Then
                                'Ingresos
                                oDoImagMovimientoIngresos.NroDocumento = ""
                                Set oDoImagMovimientoIngresos = mo_ReglasImagenes.ImagMovimientoIngresosSeleccionarPorId(lnSalidasImg)
                                If Not IsNull(oDoImagMovimientoIngresos.NroDocumento) Then
                                    LcTexto2 = oDoImagMovimientoIngresos.NroDocumento
                                End If
                            Else
                                If oDoImagMovimiento.IdTipoConcepto = 2 Then
                                   'salidas por Deterioro/vencimiento/etc
                                    Set rsReporte = mo_ReglasImagenes.ImagMovimientoSalidasSeleccionarPorIdMovimiento(lnSalidasImg)
                                    If rsReporte.RecordCount > 0 Then
                                       LcTexto2 = rsReporte.Fields!motivo
                                    End If
                                Else
                                   'salida  por toma de imagen
                                   Set rsReporte = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorIdMovimiento(lnSalidasImg)
                                   If rsReporte.RecordCount > 0 Then
                                      If Not IsNull(rsReporte.Fields!NroHistoriaClinica) Then
                                         LcTexto2 = Trim(Str(rsReporte.Fields!NroHistoriaClinica)) & " " & Trim(rsReporte.Fields!ApellidoPaterno) & " " & Trim(rsReporte.Fields!ApellidoMaterno) & " " & rsReporte.Fields!PrimerNombre
                                      End If
                                   End If
                                End If
                            End If
                            mrs_Tmp.AddNew
                            mrs_Tmp.Fields!FechaCreacion = Format(mrs_Tmp1.Fields!fechaHora, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
                            mrs_Tmp.Fields!HoraCreacion = Format(mrs_Tmp1.Fields!fechaHora, SIGHEntidades.DevuelveHoraSoloFormato_HM)
                            mrs_Tmp.Fields!MovTipo = LcTexto1
                            mrs_Tmp.Fields!movNumero = lnSalidasImg
                            mrs_Tmp.Fields!Concepto = LcTexto2
                            mrs_Tmp.Fields!Lote = lcTexto3
                            mrs_Tmp.Fields!fDestino = Trim(mrs_Tmp1.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp1.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp1.Fields!nombres) & "   (Pc: " & Trim(mrs_Tmp1.Fields!nombrePC) & ")"
                            mrs_Tmp.Update
                       End If
                End If
                mrs_Tmp1.MoveNext
             Loop
             mrs_Tmp.Sort = "fechaCreacion,HoraCreacion,movNumero"
             'Reporte
             mflgContinuar = True
             Set crReport = crApp.OpenReport(App.Path & "\plantillas\ImagAuditoria.rpt", 1)
             ' Parametros del reporte
             Set crParamDefs = crReport.ParameterFields
             For Each crParamDef In crParamDefs
                 Select Case crParamDef.ParameterFieldName
                     Case "subTitulo"
                         crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                    Case "lcEESS"
                        crParamDef.AddCurrentValue (lcTitEESS)
                    Case "lcEESSdireccion"
                        crParamDef.AddCurrentValue (lcTitDireccion)
                    Case "lcEESStelefono"
                        crParamDef.AddCurrentValue (lcTitTelefono)
                 End Select
             Next
             crReport.Database.SetDataSource mrs_Tmp
        End If
    Case "rInsumoPorTipoServ"
        Set rsReporte = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorFechasPToInsumos(lnIdPuntoCarga, mda_FechaInicio, mda_FechaFin, sghPorIdProductoMasFecha)
        If rsReporte.RecordCount > 0 Then
            If mb_ConsiderarRecalculo = True Then   'solo cantidades falladas
               rsReporte.Filter = "cantidadFallada>0"
            End If
            If rsReporte.RecordCount > 0 Then
                GenerarRecordsetTemporal lc_TipoReporte
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                    lnIdProducto = rsReporte.Fields!idProducto
                    lcCodigo = rsReporte.Fields!codigo
                    lcNombre = rsReporte.Fields!nombre
                    '*******Saldo Inicial********
                    lnSalidas = 0: lnSalidasImg = 0: lnSaldoInicial = 0
                    Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto
                       If mb_ConsiderarRecalculo = True Then
                            'solo cantidades falladas
                            If rsReporte.Fields!idTipoServicio = 1 Then
                               'CE
                               lnSalidas = lnSalidas + rsReporte.Fields!cantidadFallada
                            ElseIf IsNull(rsReporte.Fields!idTipoServicio) Then
                               'Externos
                               lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!cantidadFallada
                            Else
                               'Hospit/emerg
                               lnSalidasImg = lnSalidasImg + rsReporte.Fields!cantidadFallada
                            End If
                       Else
                            'solo cantidades buenas
                            If rsReporte.Fields!idTipoServicio = 1 Then
                               'CE
                               lnSalidas = lnSalidas + (rsReporte.Fields!Cantidad - rsReporte.Fields!cantidadFallada)
                            ElseIf IsNull(rsReporte.Fields!idTipoServicio) Then
                               'Externos
                               lnSaldoInicial = lnSaldoInicial + (rsReporte.Fields!Cantidad - rsReporte.Fields!cantidadFallada)
                            Else
                               'Hospit/emerg
                               lnSalidasImg = lnSalidasImg + (rsReporte.Fields!Cantidad - rsReporte.Fields!cantidadFallada)
                            End If
                       End If
                       rsReporte.MoveNext
                       If rsReporte.EOF Then
                          Exit Do
                       End If
                    Loop
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!codigo = lcCodigo
                    mrs_Tmp.Fields!nombre = lcNombre
                    mrs_Tmp.Fields!CE = lnSalidas
                    mrs_Tmp.Fields!Hospitalizados = lnSalidasImg
                    mrs_Tmp.Fields!Externos = lnSaldoInicial
                    mrs_Tmp.Fields!total = lnSalidas + lnSalidasImg + lnSaldoInicial
                    mrs_Tmp.Update
                Loop
             End If
             mrs_Tmp.Sort = "total desc"
             'Reporte
             mflgContinuar = True
             Set crReport = crApp.OpenReport(App.Path & "\plantillas\ImagInsumoXtipoServicio.rpt", 1)
             'Parametros del reporte
             Set crParamDefs = crReport.ParameterFields
             For Each crParamDef In crParamDefs
                 Select Case crParamDef.ParameterFieldName
                     Case "subTitulo"
                         crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                    Case "lcEESS"
                        crParamDef.AddCurrentValue (lcTitEESS)
                    Case "lcEESSdireccion"
                        crParamDef.AddCurrentValue (lcTitDireccion)
                    Case "lcEESStelefono"
                        crParamDef.AddCurrentValue (lcTitTelefono)
                 End Select
             Next
             crReport.Database.SetDataSource mrs_Tmp
        End If
    Case "rInsumoPorServicio"
        Set rsReporte = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorFechasPToInsumos(lnIdPuntoCarga, mda_FechaInicio, mda_FechaFin, sghPorIdProductoMasIdServiciopaciente)
        If rsReporte.RecordCount > 0 Then
            Select Case lnOrdenadoPor
            Case 0   'solo Pacientes EXTERNOS (pagantes)
                 rsReporte.Filter = "idTipoServicio=null"
            Case 1   'solo CE
                 rsReporte.Filter = "idTipoServicio=1"
            Case 2   'Hospitalizados                                      'debb-29/03/2012
                 rsReporte.Filter = "idTipoServicio=3"
            Case 3   'Emergencia
                 rsReporte.Filter = "idTipoServicio=2  or idTipoServicio=4"
            Case Else  'solo Pacientes Externos (con algun Seguro)    'debb-29/03/2012
                 rsReporte.Filter = "idTipoServicio>4"
            End Select
            If rsReporte.RecordCount > 0 Then
                GenerarRecordsetTemporal lc_TipoReporte
                rsReporte.MoveFirst
                Do While Not rsReporte.EOF
                    lnIdProducto = rsReporte.Fields!idProducto
                    lcCodigo = rsReporte.Fields!codigo
                    lcNombre = rsReporte.Fields!nombre
                    If lnOrdenadoPor = 0 Then
                        lnSaldoInicial = 0
                        LcTexto1 = "Externo"
                        lnSalidas = 0
                        Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto
                           lnSalidas = lnSalidas + (rsReporte.Fields!Cantidad - rsReporte.Fields!cantidadFallada)
                           rsReporte.MoveNext
                           If rsReporte.EOF Then
                              Exit Do
                           End If
                        Loop
                    Else
                        lnSaldoInicial = rsReporte.Fields!idServicioPaciente
                        LcTexto1 = rsReporte.Fields!dServicio
                        lnSalidas = 0
                        Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProducto And lnSaldoInicial = rsReporte.Fields!idServicioPaciente
                           lnSalidas = lnSalidas + (rsReporte.Fields!Cantidad - rsReporte.Fields!cantidadFallada)
                           rsReporte.MoveNext
                           If rsReporte.EOF Then
                              Exit Do
                           End If
                        Loop
                    End If
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!codigo = lcCodigo
                    mrs_Tmp.Fields!nombre = lcNombre
                    mrs_Tmp.Fields!idServicio = lnSaldoInicial
                    mrs_Tmp.Fields!dServicio = LcTexto1
                    mrs_Tmp.Fields!total = lnSalidas
                    mrs_Tmp.Update
                 Loop
                 'Reporte
                 mflgContinuar = True
                 Set crReport = crApp.OpenReport(App.Path & "\plantillas\ImagInsumoXServicio.rpt", 1)
                 'Parametros del reporte
                 Set crParamDefs = crReport.ParameterFields
                 For Each crParamDef In crParamDefs
                     Select Case crParamDef.ParameterFieldName
                         Case "subTitulo"
                             crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                        Case "lcEESS"
                            crParamDef.AddCurrentValue (lcTitEESS)
                        Case "lcEESSdireccion"
                            crParamDef.AddCurrentValue (lcTitDireccion)
                        Case "lcEESStelefono"
                            crParamDef.AddCurrentValue (lcTitTelefono)
                     End Select
                 Next
                 crReport.Database.SetDataSource mrs_Tmp
             End If
        End If
    Case "rProducPagoDeuda"
        Set rsReporte = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorFechasPuntoCarga(lnIdPuntoCarga, mda_FechaInicio, mda_FechaFin, sghPorIdFuenteFinanciamientoIdTipoServicio)
        If rsReporte.RecordCount > 0 Then
            GenerarRecordsetTemporal lc_TipoReporte
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
               lnId1 = rsReporte.Fields!idFuenteFinanciamiento
               LcTexto1 = rsReporte.Fields!nombrePlan
               lnId3 = rsReporte.Fields!IdOrden
               If IsNull(rsReporte.Fields!idTipoServicio) Then
                    lnId2 = 0
                    LcTexto2 = "EXTERNO"
                    lnSalidas = 0: lnPrecio = 0
                    Do While Not rsReporte.EOF And lnId1 = rsReporte.Fields!idFuenteFinanciamiento
                       lnSalidas = lnSalidas + rsReporte.Fields!Cantidad
                       lnPrecio = lnPrecio + rsReporte.Fields!total
                       rsReporte.MoveNext
                       If rsReporte.EOF Or (Not IsNull(rsReporte.Fields!idTipoServicio)) Then
                          Exit Do
                       End If
                    Loop
               Else
                    lnId2 = rsReporte.Fields!idTipoServicio
                    LcTexto2 = IIf(lnId2 = 1, "CE", IIf(lnId2 = 3, "HOSP", "EMERG"))
                    lnSalidas = 0: lnPrecio = 0
                    Do While Not rsReporte.EOF And lnId1 = rsReporte.Fields!idFuenteFinanciamiento And lnId2 = rsReporte.Fields!idTipoServicio
                       lnSalidas = lnSalidas + rsReporte.Fields!Cantidad
                       lnPrecio = lnPrecio + rsReporte.Fields!total
                       rsReporte.MoveNext
                       If rsReporte.EOF Then
                          Exit Do
                       End If
                    Loop
                End If
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!idPlan = lnId1
                mrs_Tmp.Fields!Plan = LcTexto1
                mrs_Tmp.Fields!dTServicio = LcTexto2
                mrs_Tmp.Fields!Cantidad = lnSalidas
                mrs_Tmp.Fields!total = lnPrecio
                mrs_Tmp.Update
            Loop
        End If
        'Reporte
        mflgContinuar = True
        Set crReport = crApp.OpenReport(App.Path & "\plantillas\ImagProducPagoDeuda.rpt", 1)
        ' Parametros del reporte
        Set crParamDefs = crReport.ParameterFields
        For Each crParamDef In crParamDefs
            Select Case crParamDef.ParameterFieldName
            Case "subTitulo"
                crParamDef.AddCurrentValue (lc_TextoDelFiltro)
            Case "lcEESS"
                crParamDef.AddCurrentValue (lcTitEESS)
            Case "lcEESSdireccion"
                crParamDef.AddCurrentValue (lcTitDireccion)
            Case "lcEESStelefono"
                crParamDef.AddCurrentValue (lcTitTelefono)
            End Select
        Next
        crReport.Database.SetDataSource mrs_Tmp
    End Select
    If mflgContinuar = True Then
       If mb_EnArchivoExcel = True Then
          If lcBuscaParametro.SeleccionaFilaParametro(284) = "S" Then
                Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
                Select Case lc_TipoReporte
                Case "rMovimientoDiario"
                    mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Movimiento Diario de Entradas y Salida", lc_TextoDelFiltro, "", Me.hwnd
                Case "rKardex"
                    mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Kardex", lc_TextoDelFiltro, "", Me.hwnd
                Case "rProduccion"
                    mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Producción por Fechas", lc_TextoDelFiltro, "", Me.hwnd
                Case "rEcogGen"
                    mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Ecografía General por Fechas", lc_TextoDelFiltro, "", Me.hwnd
                Case "rRayosX"
                    mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Rayos X por Fechas", lc_TextoDelFiltro, "", Me.hwnd
                Case "rTomografia"
                    mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Tomografía por Fechas", lc_TextoDelFiltro, "", Me.hwnd
                Case "rEcogObs"
                    mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Ecografía Obstétrica por Fechas", lc_TextoDelFiltro, "", Me.hwnd
                Case "rAuditoriaImg"
                    mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Auditoria", lc_TextoDelFiltro, "", Me.hwnd
                Case "rInsumoPorTipoServ"
                    mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Consumo de Insumos por Tipo de Servicio", lc_TextoDelFiltro, "", Me.hwnd
                Case "rInsumoPorServicio"
                    mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Consumo de Insumos por Servicio", lc_TextoDelFiltro, "", Me.hwnd
                Case "rProducPagoDeuda"
                    mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Producción, Pagos y Deudas  por Fechas", lc_TextoDelFiltro, "", Me.hwnd
                End Select
                Set mo_ReglasReportes = Nothing
          Else
                lcArchivoExcel = lcBuscaParametro.SeleccionaFilaParametro(269)
                crReport.ExportOptions.DestinationType = crEDTDiskFile
                crReport.ExportOptions.FormatType = crEFTExcel70
                crReport.ExportOptions.DiskFileName = lcArchivoExcel
                crReport.Export (False)
                MsgBox "Se generó el archivo " & lcArchivoExcel
          End If
        End If
        CrvReportes.ReportSource = crReport
        CrvReportes.ViewReport
        CrvReportes.Zoom 120
        '
        mo_AdminServComunes.grabaTablaAuditoria crReport.Database.Tables.Item(1).Name & " " & _
                             Mid(lc_TextoDelFiltro, IIf(InStr(lc_TextoDelFiltro, "FILTROS: ") > 0, 10, 1))   'debb-27/05/2015

    End If
    Screen.MousePointer = vbDefault
    Set crParamDefs = Nothing
    Set crParamDef = Nothing
    Set oConexion = Nothing
    Set mo_ReglasImagenes = Nothing
    Set rsReporte = Nothing
    Set rsTmp = Nothing
    Set oDoImagMovimientoIngresos = Nothing
    Set oDoImagMovimiento = Nothing
    LimpiarVariablesDeMemoria
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
'Resume
    If Err.Number = -2147206461 Then
        MsgBox "El archivo de reporte no se encuentra, restáurelo de los discos de instalación", _
            vbCritical + vbOKOnly
    Else
        MsgBox Err.Description & Chr(13) & "Movimiento: " & lnIdMovimientoError
    End If
    mflgContinuar = False
    Screen.MousePointer = vbDefault
    Exit Sub
    Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set crReport = Nothing
    Set crApp = Nothing
    LimpiarVariablesDeMemoria
End Sub





Private Sub Form_Resize()
    CrvReportes.Top = 0
    CrvReportes.Left = 0
    CrvReportes.Height = ScaleHeight
    CrvReportes.Width = ScaleWidth
End Sub

Sub GenerarRecordsetTemporal(lcReporte As String)
    With mrs_Tmp
         Select Case lcReporte
         Case "rMovimientoDiario"
                .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
                .Fields.Append "Nombre", adVarChar, 255, adFldIsNullable            'debb-18/05/2016
                .Fields.Append "saldoI", adInteger, 4, adFldIsNullable
                .Fields.Append "Ingresos", adInteger, 4, adFldIsNullable
                .Fields.Append "Salidas", adInteger, 4, adFldIsNullable
                .Fields.Append "SalidasImg", adInteger, 4, adFldIsNullable
                .Fields.Append "TotSalidas", adInteger, 4, adFldIsNullable
                .Fields.Append "saldoF", adInteger, 4, adFldIsNullable
         Case "rKardex"
                .Fields.Append "Fecha", adVarChar, 16, adFldIsNullable
                .Fields.Append "MovTipo", adVarChar, 1, adFldIsNullable
                .Fields.Append "idMovimiento", adInteger, 4, adFldIsNullable
                .Fields.Append "Ingresos", adInteger, 4, adFldIsNullable
                .Fields.Append "salidas", adInteger, 4, adFldIsNullable
                .Fields.Append "saldo", adInteger, 4, adFldIsNullable
                .Fields.Append "Concepto", adVarChar, 100, adFldIsNullable
                .Fields.Append "Observacion", adVarChar, 100, adFldIsNullable
         Case "rProduccion"
                .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
                .Fields.Append "Nombre", adVarChar, 255, adFldIsNullable        'debb-18/05/2016
                .Fields.Append "Buenos", adInteger, 4, adFldIsNullable
                .Fields.Append "Fallados", adInteger, 4, adFldIsNullable
                .Fields.Append "Repetidos", adInteger, 4, adFldIsNullable
                .Fields.Append "Total", adInteger, 4, adFldIsNullable
                .Fields.Append "Importe", adDouble
                .Fields.Append "Emergencia", adInteger, 4, adFldIsNullable          'franklin 2017
                .Fields.Append "ImpCE", adDouble
                .Fields.Append "ImpEmer", adDouble
                .Fields.Append "ImpHosp", adDouble
                .Fields.Append "ImpExt", adDouble
                .Fields.Append "Financ", adVarChar, 150, adFldIsNullable
         Case "rEcogGen"
                If mb_EnResumen = True Then
                    .Fields.Append "plan", adVarChar, 50, adFldIsNullable
                    .Fields.Append "numero", adInteger, 4, adFldIsNullable
                    .Fields.Append "Total", adDouble
                Else
                    .Fields.Append "Fecha", adVarChar, 10, adFldIsNullable
                    .Fields.Append "FechaHr", adVarChar, 16, adFldIsNullable
                    .Fields.Append "idMovimiento", adInteger, 4, adFldIsNullable
                    .Fields.Append "paciente", adVarChar, 160, adFldIsNullable
                    .Fields.Append "sexo", adVarChar, 1, adFldIsNullable
                    .Fields.Append "edad", adInteger, 4, adFldIsNullable
                    .Fields.Append "NroCuenta", adInteger, 4, adFldIsNullable
                    .Fields.Append "Procedencia", adVarChar, 50, adFldIsNullable
                    .Fields.Append "plan", adVarChar, 50, adFldIsNullable
                    .Fields.Append "Recibo", adVarChar, 15, adFldIsNullable
                    .Fields.Append "Resultado", adVarChar, 3000, adFldIsNullable
                    .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
                    .Fields.Append "Nombre", adVarChar, 255, adFldIsNullable            'debb-18/05/2016
                    .Fields.Append "Total", adDouble
                End If
         Case "rRayosX"
                If mb_EnResumen = True Then
                    .Fields.Append "plan", adVarChar, 50, adFldIsNullable
                    .Fields.Append "numero", adInteger, 4, adFldIsNullable
                    .Fields.Append "Total", adDouble
                Else
                    .Fields.Append "Fecha", adVarChar, 10, adFldIsNullable
                    .Fields.Append "FechaHr", adVarChar, 16, adFldIsNullable
                    .Fields.Append "idMovimiento", adInteger, 4, adFldIsNullable
                    .Fields.Append "paciente", adVarChar, 160, adFldIsNullable
                    .Fields.Append "sexo", adVarChar, 1, adFldIsNullable
                    .Fields.Append "edad", adInteger, 4, adFldIsNullable
                    .Fields.Append "NroCuenta", adInteger, 4, adFldIsNullable
                    .Fields.Append "Procedencia", adVarChar, 50, adFldIsNullable
                    .Fields.Append "plan", adVarChar, 50, adFldIsNullable
                    .Fields.Append "Recibo", adVarChar, 15, adFldIsNullable
                    .Fields.Append "Resultado", adVarChar, 3000, adFldIsNullable
                    .Fields.Append "Recoje", adVarChar, 30, adFldIsNullable
                    .Fields.Append "Responsable", adVarChar, 3, adFldIsNullable
                    .Fields.Append "Zona", adVarChar, 50, adFldIsNullable
                    .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
                    .Fields.Append "Nombre", adVarChar, 255, adFldIsNullable            'debb-18/05/2016
                    .Fields.Append "Total", adDouble
                End If
         Case "rTomografia"
                If mb_EnResumen = True Then
                    .Fields.Append "plan", adVarChar, 50, adFldIsNullable
                    .Fields.Append "numero", adInteger, 4, adFldIsNullable
                    .Fields.Append "Total", adDouble
                Else
                    .Fields.Append "Fecha", adVarChar, 10, adFldIsNullable
                    .Fields.Append "FechaHr", adVarChar, 16, adFldIsNullable
                    .Fields.Append "idMovimiento", adInteger, 4, adFldIsNullable
                    .Fields.Append "paciente", adVarChar, 160, adFldIsNullable
                    .Fields.Append "sexo", adVarChar, 1, adFldIsNullable
                    .Fields.Append "edad", adInteger, 4, adFldIsNullable
                    .Fields.Append "NroCuenta", adInteger, 4, adFldIsNullable
                    .Fields.Append "Procedencia", adVarChar, 50, adFldIsNullable
                    .Fields.Append "plan", adVarChar, 50, adFldIsNullable
                    .Fields.Append "Recibo", adVarChar, 15, adFldIsNullable
                    .Fields.Append "Resultado", adVarChar, 3000, adFldIsNullable
                    .Fields.Append "EsContraste", adVarChar, 2, adFldIsNullable
                    .Fields.Append "EsIonico", adVarChar, 2, adFldIsNullable
                    .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
                    .Fields.Append "Nombre", adVarChar, 255, adFldIsNullable            'debb-18/05/2016
                    .Fields.Append "Cantidad", adInteger, 4, adFldIsNullable
                    .Fields.Append "Total", adDouble
                    .Fields.Append "TotalIR", adDouble
                End If
         Case "rEcogObs"
                If mb_EnResumen = True Then
                    .Fields.Append "plan", adVarChar, 50, adFldIsNullable
                    .Fields.Append "numero", adInteger, 4, adFldIsNullable
                    .Fields.Append "Total", adDouble
                Else
                    .Fields.Append "Fecha", adVarChar, 10, adFldIsNullable
                    .Fields.Append "FechaHr", adVarChar, 16, adFldIsNullable
                    .Fields.Append "idMovimiento", adInteger, 4, adFldIsNullable
                    .Fields.Append "paciente", adVarChar, 160, adFldIsNullable
                    .Fields.Append "sexo", adVarChar, 1, adFldIsNullable
                    .Fields.Append "edad", adInteger, 4, adFldIsNullable
                    .Fields.Append "NroCuenta", adInteger, 4, adFldIsNullable
                    .Fields.Append "Procedencia", adVarChar, 50, adFldIsNullable
                    .Fields.Append "plan", adVarChar, 50, adFldIsNullable
                    .Fields.Append "Recibo", adVarChar, 15, adFldIsNullable
                    .Fields.Append "Resultado", adVarChar, 3000, adFldIsNullable
                    .Fields.Append "Referido", adVarChar, 2, adFldIsNullable
                    .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
                    .Fields.Append "Nombre", adVarChar, 255, adFldIsNullable        'debb-18/05/2016
                    .Fields.Append "Total", adDouble
                End If
         Case "rAuditoriaImg"
                .Fields.Append "FechaCreacion", adDate, 10, adFldIsNullable
                .Fields.Append "HoraCreacion", adVarChar, 5, adFldIsNullable
                .Fields.Append "MovTipo", adVarChar, 1, adFldIsNullable
                .Fields.Append "MovNumero", adVarChar, 10, adFldIsNullable
                .Fields.Append "Concepto", adVarChar, 200, adFldIsNullable
                .Fields.Append "fOrigen", adVarChar, 200, adFldIsNullable
                .Fields.Append "Lote", adVarChar, 20, adFldIsNullable
                .Fields.Append "FechaVencimiento", adDate, 10, adFldIsNullable
                .Fields.Append "fDestino", adVarChar, 200, adFldIsNullable
                .Fields.Append "Estado", adVarChar, 30, adFldIsNullable
                .Fields.Append "Total", adDouble
         Case "rInsumoPorTipoServ"
                .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
                .Fields.Append "Nombre", adVarChar, 255, adFldIsNullable            'debb-18/05/2016
                .Fields.Append "CE", adInteger, 4, adFldIsNullable
                .Fields.Append "Hospitalizados", adInteger, 4, adFldIsNullable
                .Fields.Append "Externos", adInteger, 4, adFldIsNullable
                .Fields.Append "Total", adInteger, 4, adFldIsNullable
         Case "rInsumoPorServicio"
                .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
                .Fields.Append "Nombre", adVarChar, 255, adFldIsNullable                'debb-18/05/2016
                .Fields.Append "idServicio", adInteger, 4, adFldIsNullable
                .Fields.Append "dServicio", adVarChar, 100, adFldIsNullable
                .Fields.Append "Total", adInteger, 4, adFldIsNullable
         Case "rProducPagoDeuda"
                .Fields.Append "idPlan", adInteger, 4, adFldIsNullable
                .Fields.Append "Plan", adVarChar, 50, adFldIsNullable
                .Fields.Append "dTServicio", adVarChar, 10, adFldIsNullable
                .Fields.Append "Cantidad", adInteger, 4, adFldIsNullable
                .Fields.Append "Total", adDouble
         End Select
         .LockType = adLockOptimistic
         .Open
    End With
End Sub



Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mrs_Tmp = Nothing
    Set mrs_Tmp1 = Nothing
End Sub

