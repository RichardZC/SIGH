VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptEtipoTarifa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para tipo tarifa
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim moConexion As New Connection
Dim lnIdServiciosHospitalarios As Long, lnIdReembolsosServicios As Long
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim lnIdTipoTarifaFarmacia As Long

Sub EjecutaFormulario()
    Dim oFormulario As New EconTipoTarifa
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub




Sub CreaDatosReporte(lnIdCajero As Long, lnIdTipoTarifaDetalle As Long, lbEnExcel As Boolean, lcTitulo As String, _
                     lcSubTitulo As String, mda_FechaInicio As Date, mda_FechaFin As Date, ml_IdPartidaFiltro As Long, _
                     lbConProrrateoEnExoneraciones As Boolean, lnHwnd As Long, lbConsiderarTodosItems As Boolean)
        Dim mrs_Tmp As New Recordset
        Dim lcPie As String
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        mo_ReglasCaja.ReporteTipoTarifa mrs_Tmp, sghSoloParaReporte, lnIdCajero, lnIdTipoTarifaDetalle, _
                                        mda_FechaInicio, mda_FechaFin, ml_IdPartidaFiltro, _
                                        lbConProrrateoEnExoneraciones, lbConsiderarTodosItems
        '
        If mrs_Tmp.RecordCount = 0 Then
             MsgBox "No existe Datos", vbInformation, "Reporte"
        Else
             If lbConProrrateoEnExoneraciones = True Then
                lcSubTitulo = lcSubTitulo & "  (p.ex.) "
             End If
        
             If ml_IdPartidaFiltro > 0 Then
                mrs_Tmp.Sort = "Descripcion"
             End If
             lcPie = ""
             If lbEnExcel = True Then
                mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, lcSubTitulo, lcPie, lnHwnd
             Else
                Set RpEPartidaResumen.DataSource = mrs_Tmp
                RpEPartidaResumen.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
                RpEPartidaResumen.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
                RpEPartidaResumen.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
                RpEPartidaResumen.Sections("cabecera").Controls("lblhora").Caption = lcBuscaParametro.RetornaHoraServidorSQL
                RpEPartidaResumen.Sections("cabecera").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaServidorSQL
                RpEPartidaResumen.Sections("cabecera").Controls("lblPc").Caption = "PC: " & sighentidades.RetornaNombrePC
                RpEPartidaResumen.Sections("cabecera").Controls("lblUsuario").Caption = "Usuario: " & lcBuscaParametro.RetornaLoginUsuario(sighentidades.Usuario)
                RpEPartidaResumen.Sections("cabecera").Controls("lblTitulo").Caption = lcTitulo
                RpEPartidaResumen.Sections("cabecera").Controls("lblSubTitulo").Caption = lcSubTitulo
                RpEPartidaResumen.Sections("cabecera").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaHoraServidorSQL
                Set RpEPartidaResumen.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
                RpEPartidaResumen.Sections("pie").Controls("lblPie").Caption = lcPie
                RpEPartidaResumen.RightMargin = 100
                RpEPartidaResumen.TopMargin = 100
                RpEPartidaResumen.LeftMargin = 100
                RpEPartidaResumen.BottomMargin = 100
                RpEPartidaResumen.Orientation = rptOrientDefault
                '
                RpEPartidaResumen.Show 1
             End If
        End If

End Sub


Sub CreaDatosReporteXitem(lnIdCajero As Long, lnIdTipoTarifaDetalle As Long, lbEnExcel As Boolean, lcTitulo As String, _
                     lcSubTitulo As String, mda_FechaInicio As Date, mda_FechaFin As Date, ml_IdPartidaFiltro As Long, _
                     lbConProrrateoEnExoneraciones As Boolean, lnHwnd As Long, lbConsiderarTodosItems As Boolean)
        Dim mrs_Tmp As New Recordset
        Dim lcPie As String
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        mo_ReglasCaja.ReporteTipoTarifa mrs_Tmp, sghSoloParaReporte, lnIdCajero, lnIdTipoTarifaDetalle, _
                                        mda_FechaInicio, mda_FechaFin, ml_IdPartidaFiltro, _
                                        lbConProrrateoEnExoneraciones, lbConsiderarTodosItems
        '
        If mrs_Tmp.RecordCount = 0 Then
             MsgBox "No existe Datos", vbInformation, "Reporte"
        Else
             If lbConProrrateoEnExoneraciones = True Then
                lcSubTitulo = lcSubTitulo & "  (p.ex.) "
             End If
        
             If ml_IdPartidaFiltro > 0 Then
                mrs_Tmp.Sort = "Descripcion"
             End If
             lcPie = ""
             If lbEnExcel = True Then
                mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, lcSubTitulo, lcPie, lnHwnd
             Else
                Set RpETipoTarifaXitem.DataSource = mrs_Tmp
                RpETipoTarifaXitem.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
                RpETipoTarifaXitem.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
                RpETipoTarifaXitem.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
                RpETipoTarifaXitem.Sections("cabecera").Controls("lblhora").Caption = lcBuscaParametro.RetornaHoraServidorSQL
                RpETipoTarifaXitem.Sections("cabecera").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaServidorSQL
                RpETipoTarifaXitem.Sections("cabecera").Controls("lblPc").Caption = "PC: " & sighentidades.RetornaNombrePC
                RpETipoTarifaXitem.Sections("cabecera").Controls("lblUsuario").Caption = "Usuario: " & lcBuscaParametro.RetornaLoginUsuario(sighentidades.Usuario)
                RpETipoTarifaXitem.Sections("cabecera").Controls("lblTitulo").Caption = lcTitulo
                RpETipoTarifaXitem.Sections("cabecera").Controls("lblSubTitulo").Caption = lcSubTitulo
                RpETipoTarifaXitem.Sections("cabecera").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaHoraServidorSQL
                Set RpETipoTarifaXitem.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
                RpETipoTarifaXitem.Sections("pie").Controls("lblPie").Caption = lcPie
                RpETipoTarifaXitem.RightMargin = 100
                RpETipoTarifaXitem.TopMargin = 100
                RpETipoTarifaXitem.LeftMargin = 100
                RpETipoTarifaXitem.BottomMargin = 100
                RpETipoTarifaXitem.Orientation = rptOrientDefault
                '
                RpETipoTarifaXitem.Show 1
             End If
        End If

End Sub

Sub CreaDatosReporteXitemIncluyeSeguros(lnIdCajero As Long, lnIdTipoTarifaDetalle As Long, lbEnExcel As Boolean, lcTitulo As String, _
                     lcSubTitulo As String, mda_FechaInicio As Date, mda_FechaFin As Date, ml_IdPartidaFiltro As Long, _
                     lbConProrrateoEnExoneraciones As Boolean, lnHwnd As Long, lbConsiderarTodosItems As Boolean, _
                     lnIdFuenteFinanciamiento As Long, lnIdMedico As Long)
        Dim mrs_Tmp As New Recordset
        Dim oRsTmp11 As New Recordset
        Dim oRsTmp12 As New Recordset
        Dim oRsTmp13 As New Recordset
        Dim oRsTmp14 As New Recordset
        Dim oRsTmp15 As New Recordset
        Dim oRsTmp16 As New Recordset
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
        Dim mo_ReglasAdmision As New ReglasAdmision
        Dim mo_ReporteUtil As New ReporteUtil
        Dim mo_ReglasServiciosHosp As New ReglasServiciosHosp
        Dim mo_ReglasFarmacia As New ReglasFarmacia
        Dim mo_ReglasCaja As New ReglasCaja
        Dim mo_ReglasLaboratorio As New ReglasLaboratorio
        Dim mo_ReglasImagenes As New ReglasImagenes
        Dim mo_ReglasFacturacion As New ReglasFacturacion
        Dim oConexion As New Connection
        Dim lcPie As String, lnIdFuenteFinanciamiento1 As Long, lnIdProducto As Long, lcFuenteFinanciamiento As String
        Dim lnCantidad As Long, lnPrecio As Double, lnTotal As Double, lbEsNuevo As Boolean, lcProducto As String
        Dim lcCodigo  As String, lnIdEspecialidad As Long, lnIdTipoFinanciamiento1 As Long, lbEncontroCta As Boolean
        Dim lnTotalFarmacia As Double, lnTotalLaboratorio As Double, lnTotalImagenes As Double
        Dim lnTotalProcedimientos As Double, lnAtendido As Long, lbContinuaCCCTA As Boolean
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        'pacientes pagantes
        If lnIdFuenteFinanciamiento = 0 Or lnIdFuenteFinanciamiento = 1 Then
            mo_ReglasCaja.ReporteTipoTarifa mrs_Tmp, sghSoloParaReporte, lnIdCajero, lnIdTipoTarifaDetalle, _
                                            mda_FechaInicio, mda_FechaFin, ml_IdPartidaFiltro, _
                                            lbConProrrateoEnExoneraciones, lbConsiderarTodosItems
        Else
            With mrs_Tmp
               .Fields.Append "Identificador", adInteger
               .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
               .Fields.Append "Descripcion", adVarChar, 300, adFldIsNullable
               .Fields.Append "ImpAnulado", adDouble
               .Fields.Append "ImpExonerado", adDouble
               .Fields.Append "ImpNormal", adDouble
               .Fields.Append "ImpCancelado", adDouble
               .Fields.Append "ImpSubTotal", adDouble
               .Fields.Append "cantidad", adInteger
               .Fields.Append "precio", adDouble
               .Fields.Append "total", adDouble
               .Fields.Append "idFuenteFinanciamiento", adInteger
               .Fields.Append "FuenteFinanciamiento", adVarChar, 100, adFldIsNullable
               .Fields.Append "puntoCarga", adVarChar, 50
               .LockType = adLockOptimistic
               .Open
            End With
        End If
        'pacientes con seguros
        Set oRsTmp11 = mo_ReglasFacturacion.FacturacionServicioFinanciamientosDespachos(mda_FechaInicio, _
                                                                    mda_FechaFin, lnIdFuenteFinanciamiento)
        If lnIdFuenteFinanciamiento = 0 Then
           oRsTmp11.Filter = "idFuenteFinanciamiento<>5"
        End If
        If oRsTmp11.RecordCount > 0 Then
           oRsTmp11.MoveFirst
           Do While Not oRsTmp11.EOF
              lnIdFuenteFinanciamiento1 = oRsTmp11!idfuentefinanciamiento
              lcFuenteFinanciamiento = oRsTmp11!descripcion
              lnIdProducto = oRsTmp11!idProducto
              lcProducto = oRsTmp11!Nombre
              lcCodigo = oRsTmp11!Codigo
              lnCantidad = 0:  lnTotal = 0
              Do While Not oRsTmp11.EOF And lnIdProducto = oRsTmp11!idProducto And lnIdFuenteFinanciamiento1 = oRsTmp11!idfuentefinanciamiento
                    lnCantidad = lnCantidad + oRsTmp11!cantidadFinanciada
                    lnTotal = lnTotal + oRsTmp11!TotalFinanciado
                    oRsTmp11.MoveNext
                    If oRsTmp11.EOF Then
                       Exit Do
                    End If
              Loop
              'kike 2017
              If lnCantidad > 0 Then
              lnPrecio = Round(lnTotal / lnCantidad, 2)
              Else
              lnPrecio = 0
              End If
              
              mrs_Tmp.AddNew
              mrs_Tmp.Fields!identificador = lnIdProducto
              mrs_Tmp.Fields!descripcion = lcProducto
              mrs_Tmp.Fields!Codigo = lcCodigo
              mrs_Tmp.Fields!precio = lnPrecio
              mrs_Tmp.Fields!ImpCancelado = lnTotal
              mrs_Tmp.Fields!Cantidad = lnCantidad
              mrs_Tmp.Fields!idfuentefinanciamiento = lnIdFuenteFinanciamiento1
              mrs_Tmp.Fields!fuenteFinanciamiento = lcFuenteFinanciamiento
              mrs_Tmp.Update
           Loop
        End If
        oRsTmp11.Close
        'elimina los CPT que corresponden a CITAS (pagantes, sis, soat, etc)
        Set oRsTmp12 = mo_ReglasServiciosHosp.EspecialidadCESeleccionarTodos
        If oRsTmp12.RecordCount > 0 Then
            If mrs_Tmp.RecordCount > 0 Then
               mrs_Tmp.MoveFirst
               Do While Not mrs_Tmp.EOF
                  oRsTmp12.MoveFirst
                  oRsTmp12.Find "idproductoConsulta=" & mrs_Tmp!identificador
                  If Not oRsTmp12.EOF Then
                     mrs_Tmp.Delete
                     mrs_Tmp.Update
                  End If
                  mrs_Tmp.MoveNext
               Loop
            End If
        End If
        'agrega los CPT que corresponden a CITAS  (pagantes, sis, soat, etc)
        If oRsTmp12.RecordCount > 0 Then
           Set oRsTmp11 = mo_ReglasAdmision.AtencionesSoloCEconFuenteFinanciamiento(mda_FechaInicio, _
                                                                    mda_FechaFin, lnIdFuenteFinanciamiento)
           If oRsTmp11.RecordCount > 0 Then
              oRsTmp11.MoveFirst
              Do While Not oRsTmp11.EOF
                    lnIdTipoFinanciamiento1 = oRsTmp11!IdFormaPago
                    lnIdFuenteFinanciamiento1 = oRsTmp11!idfuentefinanciamiento
                    lcFuenteFinanciamiento = oRsTmp11!descripcion
                    lnIdEspecialidad = oRsTmp11!IdEspecialidad
                    lnCantidad = 0:  lnTotal = 0: lnAtendido = 0
                    lnTotalFarmacia = 0: lnTotalLaboratorio = 0: lnTotalImagenes = 0: lnTotalProcedimientos = 0
                    Do While Not oRsTmp11.EOF And lnIdFuenteFinanciamiento1 = oRsTmp11!idfuentefinanciamiento And lnIdEspecialidad = oRsTmp11!IdEspecialidad
                       lbContinuaCCCTA = True
                       If lnIdMedico > 0 Then
                            If oRsTmp11!idMedicoIngreso <> lnIdMedico Then
                               lbContinuaCCCTA = False
                            End If
                       End If
                       If lbContinuaCCCTA = True Then
                          If oRsTmp15.State = 1 Then oRsTmp15.Close
                          Set oRsTmp15 = mo_ReglasCaja.CajaComprobantesPagoXcuenta(oRsTmp11!idCuentaAtencion, oConexion)
                          'consumo en Farmacia
                          If lnIdTipoFinanciamiento1 = sghTipoFinanciamiento.sghPacienteNormal Then
                             oRsTmp15.Filter = "idEstadoComprobante=4 and IdTipoOrden=2"
                             If oRsTmp15.RecordCount > 0 Then
                                  oRsTmp15.MoveFirst
                                  Do While Not oRsTmp15.EOF
                                     lnTotalFarmacia = lnTotalFarmacia + oRsTmp15!Total
                                     oRsTmp15.MoveNext
                                  Loop
                             End If
                          Else
                             Set oRsTmp14 = mo_ReglasFarmacia.FarmMovimientoVentasDetalleSeleccionarPorCuenta(oRsTmp11!idCuentaAtencion, oConexion)
                             If oRsTmp14.RecordCount > 0 Then
                               oRsTmp14.MoveFirst
                               Do While Not oRsTmp14.EOF
                                  lnTotalFarmacia = lnTotalFarmacia + oRsTmp14!Total
                                  oRsTmp14.MoveNext
                               Loop
                             End If
                             oRsTmp14.Close
                          End If
                          'consumo en Laboratorio
                          If lnIdTipoFinanciamiento1 = sghTipoFinanciamiento.sghPacienteNormal Then
                               oRsTmp15.Filter = "idEstadoComprobante=4 and IdTipoOrden<>2"
                               If oRsTmp15.RecordCount > 0 Then
                                    oRsTmp15.MoveFirst
                                    Do While Not oRsTmp15.EOF
                                       Set oRsTmp16 = mo_ReglasLaboratorio.LabMovimientoLaboratorioSeleccionarPorIdComprobantePago(oRsTmp15!IdComprobantePago, oConexion)
                                       If oRsTmp16.RecordCount > 0 Then
                                          lnTotalLaboratorio = lnTotalLaboratorio + oRsTmp15!Total
                                       End If
                                       oRsTmp15.MoveNext
                                    Loop
                               End If
                          Else
                               Set oRsTmp14 = mo_ReglasLaboratorio.LabMovimientoCPTXcuenta(oRsTmp11!idCuentaAtencion, oConexion)
                               If oRsTmp14.RecordCount > 0 Then
                                 oRsTmp14.MoveFirst
                                 Do While Not oRsTmp14.EOF
                                    lnTotalLaboratorio = lnTotalLaboratorio + oRsTmp14!Importe
                                    oRsTmp14.MoveNext
                                 Loop
                               End If
                               oRsTmp14.Close
                          End If
                          'consumo en Imagenes
                          If lnIdTipoFinanciamiento1 = sghTipoFinanciamiento.sghPacienteNormal Then
                               oRsTmp15.Filter = "idEstadoComprobante=4 and IdTipoOrden<>2"
                               If oRsTmp15.RecordCount > 0 Then
                                    oRsTmp15.MoveFirst
                                    Do While Not oRsTmp15.EOF
                                       Set oRsTmp16 = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorIdComprobantePago(oRsTmp15!IdComprobantePago, oConexion)
                                       If oRsTmp16.RecordCount > 0 Then
                                          lnTotalImagenes = lnTotalImagenes + oRsTmp15!Total
                                       End If
                                       oRsTmp15.MoveNext
                                    Loop
                               End If
                          Else
                               Set oRsTmp14 = mo_ReglasImagenes.ImagMovimientoCPTXcuenta(oRsTmp11!idCuentaAtencion, oConexion)
                               If oRsTmp14.RecordCount > 0 Then
                                 oRsTmp14.MoveFirst
                                 Do While Not oRsTmp14.EOF
                                    lnTotalImagenes = lnTotalImagenes + oRsTmp14!Importe
                                    oRsTmp14.MoveNext
                                 Loop
                               End If
                               oRsTmp14.Close
                          End If
                          'consumo en CPT en el mismo CONSULTORIO
                          Set oRsTmp14 = mo_ReglasFacturacion.FacturacionServicioDespachoSeleccionarPorCuenta(oRsTmp11!idCuentaAtencion, oConexion)
                          oRsTmp14.Filter = "idPuntoCarga=1 and IdEstadoFacturacion<>9"
                          If oRsTmp14.RecordCount > 0 Then
                                oRsTmp14.MoveFirst
                                Do While Not oRsTmp14.EOF
                                   lnTotalProcedimientos = lnTotalProcedimientos + oRsTmp14!Total
                                   oRsTmp14.MoveNext
                                Loop
                          End If
                          oRsTmp14.Close
                          
                          lnCantidad = lnCantidad + 1
                          If Not IsNull(oRsTmp11!fechaEgreso) Then lnAtendido = lnAtendido + 1
                          oRsTmp11.MoveNext
                          If oRsTmp11.EOF Then
                             Exit Do
                          End If
                      Else
                          oRsTmp11.MoveNext
                          If oRsTmp11.EOF Then
                             Exit Do
                          End If
                      End If
                    Loop
                    
                    lnIdProducto = 0
                    lcProducto = ""
                    lcCodigo = ""
                    lnPrecio = 0
                    oRsTmp12.MoveFirst
                    oRsTmp12.Find "idEspecialidad=" & lnIdEspecialidad
                    If Not oRsTmp12.EOF Then
                       lnIdProducto = oRsTmp12!idProductoConsulta
                       
                    End If
                    Set oRsTmp13 = mo_ReglasFacturacion.FactCatalogoServiciosHospSeleccionarPorIdYtipoFinanciamiento(lnIdProducto, lnIdTipoFinanciamiento1)
                    If oRsTmp13.RecordCount > 0 Then
                       lcProducto = oRsTmp13!Nombre
                       lcCodigo = oRsTmp13!Codigo
                       lnPrecio = oRsTmp13!PrecioUnitario
                    Else
                       oRsTmp13.Close
                       Set oRsTmp13 = mo_ReglasFacturacion.FactCatalogoServiciosXIdProducto(lnIdProducto, oConexion)
                       If oRsTmp13.RecordCount > 0 Then
                          lcProducto = oRsTmp13!Nombre
                          lcCodigo = oRsTmp13!Codigo
                       End If
                    End If
                    oRsTmp13.Close
                    Set oRsTmp13 = mo_ReglasServiciosHosp.EspecialidadesSeleccionarXid(lnIdEspecialidad, oConexion)
                    If oRsTmp13.RecordCount > 0 Then
                       lcProducto = lcProducto & " (" & Trim(oRsTmp13!Nombre) & ")"
                    End If
                    oRsTmp13.Close
                    lnTotal = lnCantidad * lnPrecio
                    
                    If lnCantidad > 0 Then
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!identificador = lnIdProducto
                        mrs_Tmp.Fields!descripcion = lcProducto
                        mrs_Tmp.Fields!Codigo = lcCodigo
                        mrs_Tmp.Fields!precio = 0
                        mrs_Tmp.Fields!ImpCancelado = 0
                        mrs_Tmp.Fields!Cantidad = lnCantidad
                        mrs_Tmp.Fields!idfuentefinanciamiento = lnIdFuenteFinanciamiento1
                        mrs_Tmp.Fields!fuenteFinanciamiento = lcFuenteFinanciamiento
                        mrs_Tmp.Fields!PuntoCarga = "CITAS"
                        mrs_Tmp.Update
                        
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!identificador = lnIdProducto
                        mrs_Tmp.Fields!descripcion = lcProducto
                        mrs_Tmp.Fields!Codigo = lcCodigo + "A"
                        mrs_Tmp.Fields!precio = lnPrecio
                        mrs_Tmp.Fields!ImpCancelado = lnPrecio * lnAtendido
                        mrs_Tmp.Fields!Cantidad = lnAtendido
                        mrs_Tmp.Fields!idfuentefinanciamiento = lnIdFuenteFinanciamiento1
                        mrs_Tmp.Fields!fuenteFinanciamiento = lcFuenteFinanciamiento
                        mrs_Tmp.Fields!PuntoCarga = "CITAS_ATENDIDAS"
                        mrs_Tmp.Update
                    End If
                    
                    If lnTotalFarmacia > 0 Then
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!identificador = lnIdProducto
                        mrs_Tmp.Fields!descripcion = lcProducto
                        mrs_Tmp.Fields!Codigo = lcCodigo + "F"
                        mrs_Tmp.Fields!precio = 0
                        mrs_Tmp.Fields!ImpCancelado = lnTotalFarmacia
                        mrs_Tmp.Fields!Cantidad = 0
                        mrs_Tmp.Fields!idfuentefinanciamiento = lnIdFuenteFinanciamiento1
                        mrs_Tmp.Fields!fuenteFinanciamiento = lcFuenteFinanciamiento
                        mrs_Tmp.Fields!PuntoCarga = "FARMACIA"
                        mrs_Tmp.Update
                    End If
                    If lnTotalLaboratorio > 0 Then
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!identificador = lnIdProducto
                        mrs_Tmp.Fields!descripcion = lcProducto
                        mrs_Tmp.Fields!Codigo = lcCodigo + "L"
                        mrs_Tmp.Fields!precio = 0
                        mrs_Tmp.Fields!ImpCancelado = lnTotalLaboratorio
                        mrs_Tmp.Fields!Cantidad = 0
                        mrs_Tmp.Fields!idfuentefinanciamiento = lnIdFuenteFinanciamiento1
                        mrs_Tmp.Fields!fuenteFinanciamiento = lcFuenteFinanciamiento
                        mrs_Tmp.Fields!PuntoCarga = "LABORATORIO"
                        mrs_Tmp.Update
                    End If
                    If lnTotalImagenes > 0 Then
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!identificador = lnIdProducto
                        mrs_Tmp.Fields!descripcion = lcProducto
                        mrs_Tmp.Fields!Codigo = lcCodigo + "I"
                        mrs_Tmp.Fields!precio = 0
                        mrs_Tmp.Fields!ImpCancelado = lnTotalImagenes
                        mrs_Tmp.Fields!Cantidad = 0
                        mrs_Tmp.Fields!idfuentefinanciamiento = lnIdFuenteFinanciamiento1
                        mrs_Tmp.Fields!fuenteFinanciamiento = lcFuenteFinanciamiento
                        mrs_Tmp.Fields!PuntoCarga = "IMAGENES"
                        mrs_Tmp.Update
                    End If
                    If lnTotalProcedimientos > 0 Then
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!identificador = lnIdProducto
                        mrs_Tmp.Fields!descripcion = lcProducto
                        mrs_Tmp.Fields!Codigo = lcCodigo + "T"
                        mrs_Tmp.Fields!precio = 0
                        mrs_Tmp.Fields!ImpCancelado = lnTotalProcedimientos
                        mrs_Tmp.Fields!Cantidad = 0
                        mrs_Tmp.Fields!idfuentefinanciamiento = lnIdFuenteFinanciamiento1
                        mrs_Tmp.Fields!fuenteFinanciamiento = lcFuenteFinanciamiento
                        mrs_Tmp.Fields!PuntoCarga = "PROCEDIMIENTOS"
                        mrs_Tmp.Update
                    End If
              Loop
           End If
           oRsTmp11.Close
        End If
        '
        If mrs_Tmp.RecordCount = 0 Then
             MsgBox "No existe Datos", vbInformation, "Reporte"
        Else
             If lbConProrrateoEnExoneraciones = True Then
                lcSubTitulo = lcSubTitulo & "  (p.ex.) "
             End If
             If lnIdFuenteFinanciamiento > 0 Then
                    '**** Una FuenteFinanciamiento (menos PARTICULAR)
                    If ml_IdPartidaFiltro > 0 Then
                       mrs_Tmp.Sort = "Descripcion,fuenteFinanciamiento"
                    End If
                    
                    lcPie = ""
                    If lbEnExcel = True Then
                       mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, lcSubTitulo, lcPie, lnHwnd
                    Else
                       Set RpETipoTarifaXitem.DataSource = mrs_Tmp
                       RpETipoTarifaXitem.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
                       RpETipoTarifaXitem.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
                       RpETipoTarifaXitem.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
                       RpETipoTarifaXitem.Sections("cabecera").Controls("lblhora").Caption = lcBuscaParametro.RetornaHoraServidorSQL
                       RpETipoTarifaXitem.Sections("cabecera").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaServidorSQL
                       RpETipoTarifaXitem.Sections("cabecera").Controls("lblPc").Caption = "PC: " & sighentidades.RetornaNombrePC
                       RpETipoTarifaXitem.Sections("cabecera").Controls("lblUsuario").Caption = "Usuario: " & lcBuscaParametro.RetornaLoginUsuario(sighentidades.Usuario)
                       RpETipoTarifaXitem.Sections("cabecera").Controls("lblTitulo").Caption = lcTitulo
                       RpETipoTarifaXitem.Sections("cabecera").Controls("lblSubTitulo").Caption = lcSubTitulo
                       RpETipoTarifaXitem.Sections("cabecera").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaHoraServidorSQL
                       Set RpETipoTarifaXitem.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
                       RpETipoTarifaXitem.Sections("pie").Controls("lblPie").Caption = lcPie
                       RpETipoTarifaXitem.RightMargin = 100
                       RpETipoTarifaXitem.TopMargin = 100
                       RpETipoTarifaXitem.LeftMargin = 100
                       RpETipoTarifaXitem.BottomMargin = 100
                       RpETipoTarifaXitem.Orientation = rptOrientDefault
                       '
                       RpETipoTarifaXitem.Show 1
                    End If
             Else
                    '******* todos  *****
                    mrs_Tmp.Sort = "descripcion,fuentefinanciamiento,puntoCarga"
                    Dim lbEsOpenOffice As Boolean, iFila As Long, iFila1 As Long, lnTotalG As Double, lnCant1 As Long
                    lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
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
                        'Dim lnHWnd As Long
                        'Abre el archivo ExcelOpenOffice
                        lcArchivoExcel = App.Path + "\Plantillas\HojaLibre.ods"
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
                        Dim oExcel As Excel.Application
                        Dim oWorkBookPlantilla As Workbook
                        Dim oWorkBook As Workbook
                        Dim oWorkSheet As Worksheet
                        Dim oRange As range
                        Dim range As Excel.range
                        Dim borders As Excel.borders
                        'Crea nueva hoja
                        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
                        Set oWorkBook = oExcel.Workbooks.Add
                        'Abre, copia y cierra la plantilla
                        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HojaLibre.xls")
                        oWorkBookPlantilla.Worksheets("Hoja_libre").Copy Before:=oWorkBook.Sheets(1)
                        oWorkBookPlantilla.Close
                        'Activa la primera hoja
                        Set oWorkSheet = oWorkBook.Sheets(1)
                        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
                        
                   End If
                

                   
                   
                   If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(0, 0).setFormula(lcTitulo)
                        Call Feuille.getcellbyposition(0, 1).setFormula(lcSubTitulo)
                        Call Feuille.getcellbyposition(0, 4).setFormula("Código")
                        Call Feuille.getcellbyposition(1, 4).setFormula("Descripción")
                        Call Feuille.getcellbyposition(4, 4).setFormula("Precio")
                        Call Feuille.getcellbyposition(5, 4).setFormula("Cantidad")
                        Call Feuille.getcellbyposition(6, 4).setFormula("Importe")
                        Call Feuille.getcellbyposition(7, 4).setFormula("Anulaciones")
                        Call Feuille.getcellbyposition(8, 4).setFormula("Exoneraciones")
                        Call Feuille.getcellbyposition(9, 4).setFormula("T.Generado")
                        Call Feuille.getcellbyposition(10, 4).setFormula("Total")
                        iFila = 4
                       Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(1) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(11) & CStr(iFila))
                       mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                       iFila = 6
                       lnTotalG = 0
                       mrs_Tmp.MoveFirst
                       Do While Not mrs_Tmp.EOF
                          lcCodigo = mrs_Tmp!Codigo
                          lcProducto = mrs_Tmp!descripcion
                          iFila1 = iFila
                          
                          Call Feuille.getcellbyposition(0, iFila - 1).setFormula(mrs_Tmp!Codigo)
                          Call Feuille.getcellbyposition(1, iFila - 1).setFormula(mrs_Tmp!descripcion)

                          
                          lnIdProducto = mrs_Tmp!identificador
                          lnTotal = 0: lnCantidad = 0: lnCant1 = 0
                          Do While Not mrs_Tmp.EOF And lcProducto = mrs_Tmp!descripcion
                                lnCantidad = lnCantidad + 1
                                iFila = iFila + 1
                                Call Feuille.getcellbyposition(2, iFila - 1).setFormula(mrs_Tmp!fuenteFinanciamiento)
                                Call Feuille.getcellbyposition(4, iFila - 1).setFormula(mrs_Tmp!precio)
                                Call Feuille.getcellbyposition(5, iFila - 1).setFormula(mrs_Tmp!Cantidad)
                                Call Feuille.getcellbyposition(6, iFila - 1).setFormula(mrs_Tmp!Total)
                                Call Feuille.getcellbyposition(7, iFila - 1).setFormula(mrs_Tmp!impAnulado)
                                Call Feuille.getcellbyposition(8, iFila - 1).setFormula(mrs_Tmp!ImpExonerado)
                                Call Feuille.getcellbyposition(9, iFila - 1).setFormula(mrs_Tmp!ImpCancelado)
                                lnTotal = lnTotal + mrs_Tmp!ImpCancelado
                                lnCant1 = lnCant1 + mrs_Tmp!Cantidad
                                mrs_Tmp.MoveNext
                                
                                If mrs_Tmp.EOF Then
                                   Exit Do
                                End If
                          Loop
                          iFila = iFila1
                          Call Feuille.getcellbyposition(5, iFila - 1).setFormula(lnCant1)
                          Call Feuille.getcellbyposition(10, iFila - 1).setFormula(lnTotal)
                          lnTotalG = lnTotalG + lnTotal
                          iFila = iFila + lnCantidad + 1
                       Loop
                       
                       Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(1) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(11) & CStr(iFila))
                       mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
                       
                       iFila = iFila + 1
                       Call Feuille.getcellbyposition(10, iFila - 1).setFormula(lnTotalG)
                       Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                        PrintArea(0).Sheet = 0
                        PrintArea(0).startcolumn = 1
                        PrintArea(0).StartRow = 0
                        PrintArea(0).EndColumn = 11
                        PrintArea(0).EndRow = iFila
                        Call Feuille.SetPrintAreas(PrintArea())
                        Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
                   Else
                       oWorkSheet.Cells(1, 1).Value = lcTitulo
                       oWorkSheet.Cells(1, 1).Font.Bold = True
                       oWorkSheet.Cells(2, 1).Value = lcSubTitulo
                       oWorkSheet.Cells(5, 1).Value = "Código"
                       oWorkSheet.Cells(5, 2).Value = "Descripción"
                       oWorkSheet.Cells(5, 5).Value = "Precio"
                       oWorkSheet.Cells(5, 6).Value = "Cantidad"
                       oWorkSheet.Cells(5, 7).Value = "Importe"
                       oWorkSheet.Cells(5, 8).Value = "Anulaciones"
                       oWorkSheet.Cells(5, 9).Value = "Exoneraciones"
                       oWorkSheet.Cells(5, 10).Value = "T.Generado"
                       oWorkSheet.Cells(5, 11).Value = "Total"
                       mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, 5, 1, 5, 11
    
                       iFila = 6
                       lnTotalG = 0
                       mrs_Tmp.MoveFirst
                       Do While Not mrs_Tmp.EOF
                          lcCodigo = mrs_Tmp!Codigo
                          lcProducto = mrs_Tmp!descripcion
                          iFila1 = iFila
                          oWorkSheet.Cells(iFila, 1).Value = mrs_Tmp!Codigo
                          oWorkSheet.Cells(iFila, 2).Value = mrs_Tmp!descripcion
                          oWorkSheet.Cells(iFila, 2).Font.Name = "Arial Narrow"
                          lnIdProducto = mrs_Tmp!identificador
                          lnTotal = 0: lnCantidad = 0: lnCant1 = 0
                          Do While Not mrs_Tmp.EOF And lcProducto = mrs_Tmp!descripcion
                                lnCantidad = lnCantidad + 1
                                iFila = iFila + 1
                                oWorkSheet.Cells(iFila, 3).Value = mrs_Tmp!fuenteFinanciamiento
                                oWorkSheet.Cells(iFila, 4).Value = mrs_Tmp!PuntoCarga
                                oWorkSheet.Cells(iFila, 5).Value = mrs_Tmp!precio
                                oWorkSheet.Cells(iFila, 6).Value = mrs_Tmp!Cantidad
                                oWorkSheet.Cells(iFila, 7).Value = mrs_Tmp!Total
                                oWorkSheet.Cells(iFila, 8).Value = mrs_Tmp!impAnulado
                                oWorkSheet.Cells(iFila, 9).Value = mrs_Tmp!ImpExonerado
                                oWorkSheet.Cells(iFila, 10).Value = mrs_Tmp!ImpCancelado
                                lnTotal = lnTotal + mrs_Tmp!ImpCancelado
                                lnCant1 = lnCant1 + mrs_Tmp!Cantidad
                                mrs_Tmp.MoveNext
                                
                                If mrs_Tmp.EOF Then
                                   Exit Do
                                End If
                          Loop
                          iFila = iFila1
                          'oWorkSheet.Cells(iFila, 6).Value = lnCant1
                          oWorkSheet.Cells(iFila, 11).Value = lnTotal
                          lnTotalG = lnTotalG + lnTotal
                          iFila = iFila + lnCantidad + 1
                       Loop
                       mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 1, iFila, 11
                       iFila = iFila + 1
                       oWorkSheet.Cells(iFila, 11).Value = lnTotalG
                       oWorkSheet.PageSetup.Orientation = xlLandscape
                       oWorkSheet.PageSetup.PrintTitleRows = "$1:$7"
                       If oWorkSheet.PageSetup.PrintArea <> "" Then oWorkSheet.PageSetup.PrintArea = "$A$1:$R$" & (iFila + 2) 'sighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
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
                        'liberar memoria
                        Set oExcel = Nothing
                        Set oWorkBookPlantilla = Nothing
                        Set oWorkBook = Nothing
                        Set oWorkSheet = Nothing
                   End If

             End If
        End If
        oConexion.Close
        Set mrs_Tmp = Nothing
        Set oRsTmp11 = Nothing
        Set lcBuscaParametro = Nothing
        Set mo_CabeceraReportes = Nothing
        Set mo_ReporteUtil = Nothing
        Set oRsTmp12 = Nothing
        Set mo_ReglasServiciosHosp = Nothing
        Set mo_ReglasAdmision = Nothing
        Set oRsTmp13 = Nothing
        Set oConexion = Nothing
        Set oRsTmp14 = Nothing
        Set oRsTmp15 = Nothing
        Set oRsTmp16 = Nothing
        Set mo_ReglasFarmacia = Nothing
        Set mo_ReglasCaja = Nothing
        Set mo_ReglasLaboratorio = Nothing
        Set mo_ReglasImagenes = Nothing
        Set mo_ReglasFacturacion = Nothing
        
End Sub



