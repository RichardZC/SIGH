VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dllFactUcEstadoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para Estado de Cuenta
'        Programado por: Barrantes D
'        Fecha: Enero 2011
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_MensajeError As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad

'***************daniel barrantes**************
'***************Grabar Cantidades/precios/Importes registrados por SIS/SOAT/
'***************/ASISTENTA SOCIAL/CONVENIO FOSPOLIS
Sub GrabaCantidadesPreciosRegistradosEnSisSoatExoConvenio(ml_idUsuarioConPermisoEnSISoEXOoSOAT As Long, _
                        ucFacturacionServicios As ADODB.Recordset, ucFacturacionBienes As ADODB.Recordset, _
                        ml_IdUsuario As Long, lnIdCuentaAtencion As Long, lnIdPaciente As Long, _
                        mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, ml_lnIdPlanActual As Long, _
                        lcPlanActual As String, lnEstadoFacturacionAtendidoOpreventa As sghEstadoFacturacion, _
                        lcMotivoAnulacion As String)
          Dim rsActualizar As New Recordset
          Dim rsBusqueda As New Recordset
          Dim oRsTmp1 As New Recordset
          Dim oRsTmp2 As New Recordset
          Dim lnIdFacturacionServicio As Long
          Dim oFacturacionServicio As New sighcomun.DOFacturacionServicios
          Dim oBuscaFacturacionServicioPorID As New SIGHDatos.FacturacionServicios
          Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
          Dim oConexion As New ADODB.Connection
10        Dim lnTotalPorPagar As Double: Dim lnDctos As Double
          Dim lnCant As Long
          Dim sSql As String
          Dim lbContinuar As Boolean
          Dim lnImporteExonerado As Double
          Dim lnIdOrden As Long
          Dim lbGeneraPagos As Boolean
          Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
          Dim oRsOrdenesPago As New Recordset
          Dim lcAccion As String
          Dim lbHuboDetallePago As Boolean
          Dim oRsTmp3 As New Recordset
          Dim lnCorrelativo As Long
20        On Error GoTo ErrAceptar
          '
30        With oRsOrdenesPago
40              .Fields.Append "IdOrden", adInteger
50              .Fields.Append "MovNumero", adVarChar, 9, adFldIsNullable
60              .Fields.Append "ImporteExonerado", adCurrency
70              .CursorType = adOpenDynamic
80              .LockType = adLockOptimistic
90              .Open
100       End With
         
110       oConexion.CommandTimeout = 300
120       oConexion.CursorLocation = adUseClient
130       oConexion.Open sighentidades.CadenaConexion
140       oConexion.BeginTrans
150       Set oBuscaFacturacionServicioPorID.Conexion = oConexion
          
          '*********************************Servicios*****************************
          Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
          Dim oFacturacionServicioPagos As New FacturacionServicioPagos
          Dim oFacturacionServicioFinan As New FacturacionServicioFin
          Dim oFactOrdenesServicio As New FactOrdenServicio
          Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
          Dim oDoFacturacionServicioFinanc As New DoFacturacionServicioFin
          Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
          Dim oDoFactOrdenServ As New DoFactOrdenServ
          Dim oDoFactOrdenServicioPagos As New DoFactOrdenServPagos
          Dim oFactServicioPagos As New FactOrdenServicioPagos
          '
160       lcAccion = ""
170       Select Case ml_idUsuarioConPermisoEnSISoEXOoSOAT
          Case 2   'SIS
180           lcAccion = "Modificó SIS"
190       Case 3   'SOAT
200           lcAccion = "Modificó SOAT"
210       Case 4   'Convenios
220           lcAccion = "Modificó CONVENIO"
230       Case 9   'EXO
240           lcAccion = "Modificó SERV.SOCIAL"
250       End Select
          
          '
260       Set rsBusqueda = ucFacturacionServicios.Clone()
270       If rsBusqueda.RecordCount > 0 Then
280            rsBusqueda.Filter = "idEstadoFacturacion = " & sghAtendido & " or idEstadoFacturacion = " & sghConPreVenta

290       End If
300       If rsBusqueda.RecordCount > 0 Then
310          Set oFacturacionServicioPagos.Conexion = oConexion
320          Set oFacturacionServicioFinan.Conexion = oConexion
330          Set oFacturacionServicioDespacho.Conexion = oConexion
340          Set oFactOrdenesServicio.Conexion = oConexion
350          Set oFactServicioPagos.Conexion = oConexion
360          rsBusqueda.Sort = "idOrden"
370          lbGeneraPagos = False
380          rsBusqueda.MoveFirst
390          Do While Not rsBusqueda.EOF
400               lnIdOrden = rsBusqueda.Fields!IdOrden
410               lnImporteExonerado = 0
420               Do While Not rsBusqueda.EOF And lnIdOrden = rsBusqueda.Fields!IdOrden
430   If Val(rsBusqueda.Fields!Codigo) = 82101 Then
440   sSql = ""
450   End If
460                   If (rsBusqueda.Fields!IdEstadoFacturacion = sghAtendido Or rsBusqueda.Fields!IdEstadoFacturacion = sghConPreVenta) And rsBusqueda!cantidad > 0 Then
470                      lbGeneraPagos = True
480                   End If
490                   lnImporteExonerado = lnImporteExonerado + rsBusqueda.Fields!importeEXO
500                   rsBusqueda.MoveNext
510                   If rsBusqueda.EOF Then
520                      Exit Do
530                   End If
540               Loop
550               oRsOrdenesPago.AddNew
560               oRsOrdenesPago.Fields!IdOrden = lnIdOrden
570               oRsOrdenesPago.Fields!ImporteExonerado = lnImporteExonerado
580               oRsOrdenesPago.Update
590          Loop
600          rsBusqueda.MoveFirst
610          Do While Not rsBusqueda.EOF
620              lbHuboDetallePago = False
If rsBusqueda.Fields!IdOrden = 327777 Then
lnIdOrden = 0
End If
630              lnIdOrden = rsBusqueda.Fields!IdOrden
                 '
640              oRsOrdenesPago.MoveFirst
650              oRsOrdenesPago.Find "idOrden=" & lnIdOrden
                 'debb 26/08/2013
660              Set oRsTmp1 = mo_ReglasFacturacion.FactOrdenServicioPagosSeleccionarPorIdOrden(lnIdOrden, oConexion)
670              If oRsTmp1.RecordCount > 0 Then
680                 oDoFactOrdenServicioPagos.idOrdenPago = oRsTmp1.Fields!idOrdenPago
690                 oDoFactOrdenServicioPagos.IdOrden = lnIdOrden
700                 If Not oFactServicioPagos.SeleccionarPorId(oDoFactOrdenServicioPagos) Then
710                       ms_MensajeError = oFactServicioPagos.MensajeError: GoTo ErrAceptar
720                 End If
730                 oDoFactOrdenServicioPagos.IdUsuarioAuditoria = ml_IdUsuario
740                 oDoFactOrdenServicioPagos.ImporteExonerado = oRsOrdenesPago.Fields!ImporteExonerado
750                 If ml_idUsuarioConPermisoEnSISoEXOoSOAT = 9 Then
760                    oDoFactOrdenServicioPagos.idUsuarioExonera = ml_IdUsuario
770                 End If
                    
780                 If Not oFactServicioPagos.Modificar(oDoFactOrdenServicioPagos) Then
790                       ms_MensajeError = oFactServicioPagos.MensajeError: GoTo ErrAceptar
800                 End If
                    
810              ElseIf lbGeneraPagos = True Then
820                 oDoFactOrdenServicioPagos.fechaCreacion = lcBuscaParametro.RetornaFechaServidorSQL
830                 oDoFactOrdenServicioPagos.IdEstadoFacturacion = lnEstadoFacturacionAtendidoOpreventa
840                 oDoFactOrdenServicioPagos.IdOrden = lnIdOrden
850                 oDoFactOrdenServicioPagos.idUsuario = ml_IdUsuario
860                 oDoFactOrdenServicioPagos.IdUsuarioAuditoria = ml_IdUsuario
870                 oDoFactOrdenServicioPagos.ImporteExonerado = oRsOrdenesPago.Fields!ImporteExonerado
880                 If ml_idUsuarioConPermisoEnSISoEXOoSOAT = 9 Then
890                    oDoFactOrdenServicioPagos.idUsuarioExonera = ml_IdUsuario
900                 End If
910                 If Not oFactServicioPagos.Insertar(oDoFactOrdenServicioPagos) Then
920                       ms_MensajeError = oFactServicioPagos.MensajeError: GoTo ErrAceptar
930                 End If
940              End If
                 '
950              oDoFacturacionServicioPagos.idOrdenPago = oDoFactOrdenServicioPagos.idOrdenPago
960              If Not oFacturacionServicioPagos.Eliminar(oDoFacturacionServicioPagos) Then
970                  ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo ErrAceptar
980              End If
                 '
990              oDoFacturacionServicioFinanc.IdOrden = lnIdOrden
1000             If ml_idUsuarioConPermisoEnSISoEXOoSOAT = 9 Then
1010                  oDoFacturacionServicioFinanc.IdTipoFinanciamiento = ml_idUsuarioConPermisoEnSISoEXOoSOAT
1020                  If Not oFacturacionServicioFinan.EliminarPorIdTipoFinanciamiento(oDoFacturacionServicioFinanc) Then
1030                      ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo ErrAceptar
1040                  End If
1050             Else
1060                  oDoFacturacionServicioFinanc.IdTipoFinanciamiento = rsBusqueda.Fields!IdTipoFinanciamiento    ' ml_idUsuarioConPermisoEnSISoEXOoSOAT
1070                  If Not oFacturacionServicioFinan.Eliminar(oDoFacturacionServicioFinanc) Then
1080                      ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo ErrAceptar
1090                  End If
1100             End If
                 '
1110             Do While Not rsBusqueda.EOF And lnIdOrden = rsBusqueda.Fields!IdOrden
1120                  If (rsBusqueda.Fields!IdEstadoFacturacion = sghAtendido Or rsBusqueda.Fields!IdEstadoFacturacion = sghConPreVenta) And rsBusqueda!cantidadPagar > 0 Then
1130                     lbContinuar = True
1140                     Select Case ml_idUsuarioConPermisoEnSISoEXOoSOAT
                         Case 2   'SIS
1150                         If IsNull(rsBusqueda.Fields!ImporteSIS) Or rsBusqueda.Fields!ImporteSIS < 0 Then
1160                            lbContinuar = False
1170                         End If
1180                     Case 3   'SOAT
1190                         If IsNull(rsBusqueda.Fields!ImporteSOAT) Or rsBusqueda.Fields!ImporteSOAT < 0 Then
1200                            lbContinuar = False
1210                         End If
1220                     Case 9   'EXO
1230                         If IsNull(rsBusqueda.Fields!importeEXO) Or rsBusqueda.Fields!importeEXO < 0 Then
1240                            lbContinuar = False
1250                         End If
1260                     End Select
1270                     lbContinuar = True
1280                     If lbContinuar Then
1290                           lnDctos = 0: lnCant = 0
1300                           If Not IsNull(rsBusqueda.Fields!ImporteSIS) Then
1310                              lnCant = lnCant + rsBusqueda.Fields!CantidadSIS
1320                           End If
1330                           If Not IsNull(rsBusqueda.Fields!ImporteSOAT) Then
1340                              lnCant = lnCant + rsBusqueda.Fields!CantidadSOAT
1350                           End If
1360                           Select Case ml_idUsuarioConPermisoEnSISoEXOoSOAT
                               Case 2   'SIS
1370                               If rsBusqueda.Fields!CantidadSIS > 0 Then
1380                                    oDoFacturacionServicioFinanc.CantidadFinanciada = rsBusqueda.Fields!CantidadSIS
1390                                    oDoFacturacionServicioFinanc.FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
1400                                    oDoFacturacionServicioFinanc.idProducto = rsBusqueda.Fields!idProducto
1410                                    oDoFacturacionServicioFinanc.IdTipoFinanciamiento = ml_idUsuarioConPermisoEnSISoEXOoSOAT
1420                                    oDoFacturacionServicioFinanc.IdUsuarioAuditoria = ml_IdUsuario
1430                                    oDoFacturacionServicioFinanc.IdUsuarioAutoriza = ml_IdUsuario
1440                                    oDoFacturacionServicioFinanc.PrecioFinanciado = rsBusqueda.Fields!precioSIS
1450                                    oDoFacturacionServicioFinanc.TotalFinanciado = rsBusqueda.Fields!ImporteSIS
1460                                    oDoFacturacionServicioFinanc.idFuenteFinanciamiento = ml_lnIdPlanActual
1470                                    oDoFacturacionServicioFinanc.IdEstadoFacturacion = 1
                                        'debb-26/08/2013
1480                                    Set oRsTmp2 = mo_ReglasFacturacion.FacturacionServicioFinanciamFiltrarXIdOrdenProductoTFinanc(oDoFacturacionServicioFinanc.IdOrden, _
                                              oDoFacturacionServicioFinanc.idProducto, oDoFacturacionServicioFinanc.IdTipoFinanciamiento, oConexion)
1490                                    If oRsTmp2.RecordCount > 0 Then
1500                                          If Not oFacturacionServicioFinan.Modificar(oDoFacturacionServicioFinanc) Then
1510                                             ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo ErrAceptar
1520                                          End If
1530                                    Else
1540                                          If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then
1550                                             ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo ErrAceptar
1560                                          End If
1570                                    End If
1580                                    oRsTmp2.Close
1590                               End If
1600                               If rsBusqueda.Fields!cantidad > 0 Then
1610                                    oDoFacturacionServicioPagos.idOrdenPago = oDoFactOrdenServicioPagos.idOrdenPago
1620                                    oDoFacturacionServicioPagos.cantidad = rsBusqueda.Fields!cantidad
1630                                    oDoFacturacionServicioPagos.idProducto = rsBusqueda.Fields!idProducto
1640                                    oDoFacturacionServicioPagos.IdUsuarioAuditoria = ml_IdUsuario
1650                                    oDoFacturacionServicioPagos.Precio = rsBusqueda.Fields!PrecioUnitario
1660                                    oDoFacturacionServicioPagos.total = Round((rsBusqueda.Fields!cantidad * rsBusqueda.Fields!PrecioUnitario), 2) - lnDctos
1670                                    If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
1680                                       ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo ErrAceptar
1690                                    End If
1700                                    lbHuboDetallePago = True
1710                               End If
1720                           Case 3   'SOAT
1730                               If rsBusqueda.Fields!CantidadSOAT > 0 Then
1740                                    oDoFacturacionServicioFinanc.CantidadFinanciada = rsBusqueda.Fields!CantidadSOAT
1750                                    oDoFacturacionServicioFinanc.FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
1760                                    oDoFacturacionServicioFinanc.idProducto = rsBusqueda.Fields!idProducto
1770                                    oDoFacturacionServicioFinanc.IdTipoFinanciamiento = ml_idUsuarioConPermisoEnSISoEXOoSOAT
1780                                    oDoFacturacionServicioFinanc.IdUsuarioAuditoria = ml_IdUsuario
1790                                    oDoFacturacionServicioFinanc.IdUsuarioAutoriza = ml_IdUsuario
1800                                    oDoFacturacionServicioFinanc.PrecioFinanciado = rsBusqueda.Fields!precioSOAT
1810                                    oDoFacturacionServicioFinanc.TotalFinanciado = rsBusqueda.Fields!ImporteSOAT
1820                                    oDoFacturacionServicioFinanc.idFuenteFinanciamiento = ml_lnIdPlanActual
1830                                    oDoFacturacionServicioFinanc.IdEstadoFacturacion = 1
1840                                    Set oRsTmp2 = mo_ReglasFacturacion.FacturacionServicioFinanciamFiltrarXIdOrdenProductoTFinanc(oDoFacturacionServicioFinanc.IdOrden, _
                                              oDoFacturacionServicioFinanc.idProducto, oDoFacturacionServicioFinanc.IdTipoFinanciamiento, oConexion)
1850                                    If oRsTmp2.RecordCount > 0 Then
1860                                          If Not oFacturacionServicioFinan.Modificar(oDoFacturacionServicioFinanc) Then
1870                                             ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo ErrAceptar
1880                                          End If
1890                                    Else
1900                                          If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then
1910                                             ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo ErrAceptar
1920                                          End If
1930                                    End If
1940                                    oRsTmp2.Close
1950                               End If
1960                               If rsBusqueda.Fields!cantidad > 0 Then
1970                                    oDoFacturacionServicioPagos.idOrdenPago = oDoFactOrdenServicioPagos.idOrdenPago
1980                                    oDoFacturacionServicioPagos.cantidad = rsBusqueda.Fields!cantidad
1990                                    oDoFacturacionServicioPagos.idProducto = rsBusqueda.Fields!idProducto
2000                                    oDoFacturacionServicioPagos.IdUsuarioAuditoria = ml_IdUsuario
2010                                    oDoFacturacionServicioPagos.Precio = rsBusqueda.Fields!PrecioUnitario
2020                                    oDoFacturacionServicioPagos.total = Round((rsBusqueda.Fields!cantidad * rsBusqueda.Fields!PrecioUnitario), 2) - lnDctos
2030                                    If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
2040                                       ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo ErrAceptar
2050                                    End If
2060                                    lbHuboDetallePago = True
2070                               End If
2080                           Case 9   'EXONERACIONES
2090                               If rsBusqueda.Fields!importeEXO > 0 Then
2100                                    oDoFacturacionServicioFinanc.CantidadFinanciada = rsBusqueda.Fields!cantidadPagar
2110                                    oDoFacturacionServicioFinanc.FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
2120                                    oDoFacturacionServicioFinanc.idProducto = rsBusqueda.Fields!idProducto
2130                                    oDoFacturacionServicioFinanc.IdTipoFinanciamiento = ml_idUsuarioConPermisoEnSISoEXOoSOAT
2140                                    oDoFacturacionServicioFinanc.IdUsuarioAuditoria = ml_IdUsuario
2150                                    oDoFacturacionServicioFinanc.IdUsuarioAutoriza = ml_IdUsuario
2160                                    oDoFacturacionServicioFinanc.PrecioFinanciado = Round(rsBusqueda.Fields!importeEXO / rsBusqueda.Fields!cantidadPagar, 2)
2170                                    oDoFacturacionServicioFinanc.TotalFinanciado = rsBusqueda.Fields!importeEXO
2180                                    oDoFacturacionServicioFinanc.idFuenteFinanciamiento = 0
2190                                    oDoFacturacionServicioFinanc.IdEstadoFacturacion = 1
                                        
2200                                    Set oRsTmp2 = mo_ReglasFacturacion.FacturacionServicioFinanciamFiltrarXIdOrdenProductoTFinanc(oDoFacturacionServicioFinanc.IdOrden, _
                                              oDoFacturacionServicioFinanc.idProducto, oDoFacturacionServicioFinanc.IdTipoFinanciamiento, oConexion)
2210                                    If oRsTmp2.RecordCount > 0 Then
2220                                          If Not oFacturacionServicioFinan.Modificar(oDoFacturacionServicioFinanc) Then
2230                                             ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo ErrAceptar
2240                                          End If
2250                                    Else
2260                                          If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then
2270                                             ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo ErrAceptar
2280                                          End If
2290                                    End If
2300                                    oRsTmp2.Close
2310                               End If
2320                               If rsBusqueda.Fields!cantidad > 0 Then
2330                                    oDoFacturacionServicioPagos.idOrdenPago = oDoFactOrdenServicioPagos.idOrdenPago
2340                                    oDoFacturacionServicioPagos.cantidad = rsBusqueda.Fields!cantidad
2350                                    oDoFacturacionServicioPagos.idProducto = rsBusqueda.Fields!idProducto
2360                                    oDoFacturacionServicioPagos.IdUsuarioAuditoria = ml_IdUsuario
2370                                    oDoFacturacionServicioPagos.Precio = rsBusqueda.Fields!PrecioUnitario
2380                                    oDoFacturacionServicioPagos.total = Round((rsBusqueda.Fields!cantidad * rsBusqueda.Fields!PrecioUnitario), 2) - lnDctos
2390                                    If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
2400                                       ms_MensajeError = oFactServicioPagos.MensajeError: GoTo ErrAceptar
2410                                    End If
2420                                    lbHuboDetallePago = True
2430                               End If
2440                           Case 4   'Convenios
2450                               If rsBusqueda.Fields!esConvenio = "Si" Then
2460                                    If rsBusqueda.Fields!CantidadConv > 0 Then
2470                                         oDoFacturacionServicioFinanc.CantidadFinanciada = rsBusqueda.Fields!CantidadConv
2480                                         oDoFacturacionServicioFinanc.FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
2490                                         oDoFacturacionServicioFinanc.idProducto = rsBusqueda.Fields!idProducto
2500                                         oDoFacturacionServicioFinanc.IdTipoFinanciamiento = ml_idUsuarioConPermisoEnSISoEXOoSOAT
2510                                         oDoFacturacionServicioFinanc.IdUsuarioAuditoria = ml_IdUsuario
2520                                         oDoFacturacionServicioFinanc.IdUsuarioAutoriza = ml_IdUsuario
2530                                         oDoFacturacionServicioFinanc.PrecioFinanciado = rsBusqueda.Fields!PrecioConv
2540                                         oDoFacturacionServicioFinanc.TotalFinanciado = Round(rsBusqueda.Fields!CantidadConv * rsBusqueda.Fields!PrecioConv, 2)
2550                                         oDoFacturacionServicioFinanc.idFuenteFinanciamiento = ml_lnIdPlanActual
2560                                         oDoFacturacionServicioFinanc.IdEstadoFacturacion = 1
                                             
2570                                          Set oRsTmp2 = mo_ReglasFacturacion.FacturacionServicioFinanciamFiltrarXIdOrdenProductoTFinanc(oDoFacturacionServicioFinanc.IdOrden, _
                                              oDoFacturacionServicioFinanc.idProducto, oDoFacturacionServicioFinanc.IdTipoFinanciamiento, oConexion)
2580                                          If oRsTmp2.RecordCount > 0 Then
2590                                                If Not oFacturacionServicioFinan.Modificar(oDoFacturacionServicioFinanc) Then
2600                                                   ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo ErrAceptar
2610                                                End If
2620                                          Else
2630                                                If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then
2640                                                   ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo ErrAceptar
2650                                                End If
2660                                          End If
2670                                          oRsTmp2.Close
2680                                    End If
2690                               Else
2700                                    If rsBusqueda.Fields!cantidad > 0 Then
2710                                         oDoFacturacionServicioPagos.idOrdenPago = oDoFactOrdenServicioPagos.idOrdenPago
2720                                         oDoFacturacionServicioPagos.cantidad = rsBusqueda.Fields!cantidad
2730                                         oDoFacturacionServicioPagos.idProducto = rsBusqueda.Fields!idProducto
2740                                         oDoFacturacionServicioPagos.IdUsuarioAuditoria = ml_IdUsuario
2750                                         oDoFacturacionServicioPagos.Precio = rsBusqueda.Fields!PrecioUnitario
2760                                         oDoFacturacionServicioPagos.total = Round(rsBusqueda.Fields!cantidad * rsBusqueda.Fields!PrecioUnitario, 2)
2770                                         If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
2780                                            ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo ErrAceptar
2790                                         End If
2800                                         lbHuboDetallePago = True
2810                                    End If
2820                               End If
2830                           End Select
2840                     End If
2850                  End If
2860                  If lbHuboDetallePago = True Then
2870                      oDoFactOrdenServicioPagos.IdOrden = lnIdOrden
2880                      If Not oFactServicioPagos.Modificar(oDoFactOrdenServicioPagos) Then
2890                            ms_MensajeError = oFactServicioPagos.MensajeError: GoTo ErrAceptar
2900                      End If
2910                  End If
2920                  rsBusqueda.MoveNext
2930                  If rsBusqueda.EOF Then
2940                     Exit Do
2950                  End If
2960             Loop
2970         Loop
                     'Yamill junio
2980          If ml_idUsuarioConPermisoEnSISoEXOoSOAT = 9 Then
2990              Set oRsTmp3 = mo_ReglasFacturacion.FacturacionCuentasAtencionExoneracion(lnIdCuentaAtencion, oConexion)
3000              If oRsTmp3.RecordCount = 0 Then
3010                 lnCorrelativo = CLng(lcBuscaParametro.SeleccionaFilaParametro(337)) + 1
3020                 If mo_ReglasFacturacion.FacturacionCuentasAtencionExonAgregar(lnIdCuentaAtencion, lnCorrelativo, _
                                                    lcMotivoAnulacion, oConexion) = True Then
3030                      mo_ReglasComunes.ParametrosActualizaValorTextoXid 337, CStr(lnCorrelativo)
3040                 End If
                  Else
                     If mo_ReglasFacturacion.FacturacionCuentasAtencionExonmodificar(lnIdCuentaAtencion, _
                                                oRsTmp3!numeroExoneracion, lcMotivoAnulacion, oConexion) = True Then
                     End If
                     
3050              End If

3060          End If
3070      End If
3080      rsBusqueda.Close
          
          '****************************************Farmacia*************************
          Dim oFactOrdenesBienes As New FactOrdenesBienes
          Dim oDoFactOrdenesBienes As New DoFactOrdenesBienes
          Dim oFacturacionBienesPagos As New FacturacionBienesPagos
          Dim oDoFacturacionBienesPagos As New DoFacturacionBienesPagos
          Dim oFacturacionBienesFinan As New FacturacionBienesFinanciam
          Dim oDoFacturacionBienesFinan As New DoFacturacionBienesFinanciam
          Dim lcMovNumero As String
          
3090      Set rsBusqueda = ucFacturacionBienes.Clone()
3100      If rsBusqueda.RecordCount > 0 Then
3110          rsBusqueda.Filter = "idEstadoFacturacion = 1"
3120      End If
3130      If rsBusqueda.RecordCount > 0 Then
3140         rsBusqueda.Sort = "movnumero"
3150         lbGeneraPagos = False
3160         Do While Not rsBusqueda.EOF
3170            lbHuboDetallePago = False
3180            lcMovNumero = rsBusqueda.Fields!MovNumero
3190            lnImporteExonerado = 0
3200            Do While Not rsBusqueda.EOF And lcMovNumero = rsBusqueda.Fields!MovNumero
3210  If rsBusqueda.Fields!idProducto = 896 Then
3220  sSql = ""
3230  End If
3240                  If rsBusqueda.Fields!IdEstadoFacturacion = 1 And rsBusqueda!cantidad > 0 Then
3250                     lbGeneraPagos = True
3260                  End If
3270                  lnImporteExonerado = lnImporteExonerado + rsBusqueda.Fields!importeEXO
3280                  rsBusqueda.MoveNext
3290                  If rsBusqueda.EOF Then
3300                     Exit Do
3310                  End If
3320              Loop
3330              oRsOrdenesPago.AddNew
3340              oRsOrdenesPago.Fields!MovNumero = lcMovNumero
3350              oRsOrdenesPago.Fields!ImporteExonerado = lnImporteExonerado
3360              oRsOrdenesPago.Update
3370         Loop
3380         Set oFactOrdenesBienes.Conexion = oConexion
3390         Set oFacturacionBienesPagos.Conexion = oConexion
3400         Set oFacturacionBienesFinan.Conexion = oConexion
3410         rsBusqueda.MoveFirst
3420         Do While Not rsBusqueda.EOF
3430            lcMovNumero = rsBusqueda.Fields!MovNumero
                 '
3440             oRsOrdenesPago.MoveFirst
3450             oRsOrdenesPago.Find "MovNumero='" & lcMovNumero & "'"
                 '
3460            oDoFactOrdenesBienes.MovTipo = "S"
3470            oDoFactOrdenesBienes.idPuntoCarga = 0
3480            oDoFactOrdenesBienes.MovNumero = rsBusqueda.Fields!MovNumero
3490            If Not oFactOrdenesBienes.SeleccionarPorMovNumero(oDoFactOrdenesBienes) Then
3500                    ms_MensajeError = oFactOrdenesBienes.MensajeError: GoTo ErrAceptar
3510            End If
3520            If lbGeneraPagos = True Then
3530                  If oDoFactOrdenesBienes.idPuntoCarga <> 5 Then
3540                    oDoFactOrdenesBienes.fechaCreacion = lcBuscaParametro.RetornaFechaServidorSQL
3550                    oDoFactOrdenesBienes.IdCuentaAtencion = lnIdCuentaAtencion
3560                    oDoFactOrdenesBienes.IdEstadoFacturacion = 1   'registrado
3570                    oDoFactOrdenesBienes.idPaciente = lnIdPaciente
3580                    oDoFactOrdenesBienes.idPreVenta = 0
3590                    oDoFactOrdenesBienes.idPuntoCarga = 5
3600                    oDoFactOrdenesBienes.idUsuario = ml_IdUsuario
3610                    oDoFactOrdenesBienes.IdUsuarioAuditoria = ml_IdUsuario
3620                    oDoFactOrdenesBienes.ImporteExonerado = oRsOrdenesPago.Fields!ImporteExonerado
3630                    If ml_idUsuarioConPermisoEnSISoEXOoSOAT = 9 Then
3640                         oDoFactOrdenesBienes.idUsuarioExonera = ml_IdUsuario
3650                    End If
3660                    If Not oFactOrdenesBienes.Insertar(oDoFactOrdenesBienes) Then
3670                        ms_MensajeError = oFactOrdenesBienes.MensajeError: GoTo ErrAceptar
3680                    End If
3690                 Else
3700                    oDoFactOrdenesBienes.ImporteExonerado = oRsOrdenesPago.Fields!ImporteExonerado
3710                    oDoFactOrdenesBienes.IdUsuarioAuditoria = ml_IdUsuario
3720                    If ml_idUsuarioConPermisoEnSISoEXOoSOAT = 9 Then
3730                         oDoFactOrdenesBienes.idUsuarioExonera = ml_IdUsuario
3740                    End If
3750                    If Not oFactOrdenesBienes.Modificar(oDoFactOrdenesBienes) Then
3760                        ms_MensajeError = oFactOrdenesBienes.MensajeError: GoTo ErrAceptar
3770                    End If
3780                 End If
3790            End If
                '
3800            If oDoFactOrdenesBienes.IdOrden > 0 Then
3810                  oDoFacturacionBienesPagos.IdOrden = oDoFactOrdenesBienes.IdOrden
3820                  Set oFacturacionBienesPagos.Conexion = oConexion
3830                  If Not oFacturacionBienesPagos.Eliminar(oDoFacturacionBienesPagos) Then
3840                     ms_MensajeError = oFacturacionBienesPagos.MensajeError: GoTo ErrAceptar
3850                  End If
3860            End If
                '
3870            oDoFacturacionBienesFinan.MovTipo = "S"
3880            oDoFacturacionBienesFinan.MovNumero = rsBusqueda.Fields!MovNumero
3890            oDoFacturacionBienesFinan.IdTipoFinanciamiento = rsBusqueda.Fields!IdTipoFinanciamiento         'ml_idUsuarioConPermisoEnSISoEXOoSOAT
                If rsBusqueda.Fields!IdTipoFinanciamiento = 2 And rsBusqueda.Fields!IdTipoFinanciamiento <> ml_idUsuarioConPermisoEnSISoEXOoSOAT Then
                    oDoFacturacionBienesFinan.IdTipoFinanciamiento = ml_idUsuarioConPermisoEnSISoEXOoSOAT
                    If Not oFacturacionBienesFinan.EliminarPorTipoFinanciamiento(oDoFacturacionBienesFinan) Then
                    End If
                Else
3900                If Not oFacturacionBienesFinan.Eliminar(oDoFacturacionBienesFinan) Then
3910                     ms_MensajeError = oFacturacionBienesFinan.MensajeError: GoTo ErrAceptar
3920                End If
                End If
3930            Do While Not rsBusqueda.EOF And lcMovNumero = rsBusqueda.Fields!MovNumero
3940                  If rsBusqueda.Fields!IdEstadoFacturacion = 1 And rsBusqueda!cantidadPagar > 0 Then
3950                     lbContinuar = True
3960                     Select Case ml_idUsuarioConPermisoEnSISoEXOoSOAT
                         Case 2   'SIS
3970                         If IsNull(rsBusqueda.Fields!ImporteSIS) Or rsBusqueda.Fields!ImporteSIS < 0 Then
3980                            lbContinuar = False
3990                         End If
4000                     Case 3   'SOAT
4010                         If IsNull(rsBusqueda.Fields!ImporteSOAT) Or rsBusqueda.Fields!ImporteSOAT < 0 Then
4020                            lbContinuar = False
4030                         End If
4040                     Case 9   'EXO
4050                         If IsNull(rsBusqueda.Fields!importeEXO) Or rsBusqueda.Fields!importeEXO < 0 Then
4060                            lbContinuar = False
4070                         End If
4080                     End Select
4090                     lbContinuar = True
4100                     If lbContinuar Then
4110                           lnDctos = 0: lnCant = 0
4120                           If Not IsNull(rsBusqueda.Fields!ImporteSIS) Then
4130                              lnCant = lnCant + rsBusqueda.Fields!CantidadSIS
4140                           End If
4150                           If Not IsNull(rsBusqueda.Fields!ImporteSOAT) Then
4160                              lnCant = lnCant + rsBusqueda.Fields!CantidadSOAT
4170                           End If
4180                           Select Case ml_idUsuarioConPermisoEnSISoEXOoSOAT
                               Case 2   'SIS
4190                               If rsBusqueda.Fields!CantidadSIS > 0 Then
4200                                    oDoFacturacionBienesFinan.CantidadFinanciada = rsBusqueda.Fields!CantidadSIS
4210                                    oDoFacturacionBienesFinan.FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
4220                                    oDoFacturacionBienesFinan.idProducto = rsBusqueda.Fields!idProducto
4230                                    oDoFacturacionBienesFinan.IdTipoFinanciamiento = ml_idUsuarioConPermisoEnSISoEXOoSOAT
4240                                    oDoFacturacionBienesFinan.IdUsuarioAuditoria = ml_IdUsuario
4250                                    oDoFacturacionBienesFinan.IdUsuarioAutoriza = ml_IdUsuario
4260                                    oDoFacturacionBienesFinan.PrecioFinanciado = rsBusqueda.Fields!precioSIS
4270                                    oDoFacturacionBienesFinan.TotalFinanciado = rsBusqueda.Fields!ImporteSIS
4280                                    oDoFacturacionBienesFinan.idFuenteFinanciamiento = ml_lnIdPlanActual
4290                                    oDoFacturacionBienesFinan.IdEstadoFacturacion = 1
                                        
4300                                    Set oRsTmp2 = mo_ReglasFacturacion.FacturacionBienesFinanciamFiltrarXmovNumeroProductoTFinanc(oDoFacturacionBienesFinan.MovNumero, _
                                                                        oDoFacturacionBienesFinan.idProducto, _
                                                                        oDoFacturacionBienesFinan.IdTipoFinanciamiento, _
                                                                        oDoFacturacionBienesFinan.MovTipo, oConexion)
4310                                    If oRsTmp2.RecordCount > 0 Then
4320                                          If Not oFacturacionBienesFinan.Modificar(oDoFacturacionBienesFinan) Then
4330                                             ms_MensajeError = oFacturacionBienesFinan.MensajeError: GoTo ErrAceptar
4340                                          End If
4350                                    Else
4360                                          If Not oFacturacionBienesFinan.Insertar(oDoFacturacionBienesFinan) Then
4370                                             ms_MensajeError = oFacturacionBienesFinan.MensajeError: GoTo ErrAceptar
4380                                          End If
4390                                    End If
4400                                    oRsTmp2.Close
                                        
4410                               End If
4420                               If rsBusqueda.Fields!cantidad > 0 Then
4430                                    oDoFacturacionBienesPagos.IdOrden = oDoFactOrdenesBienes.IdOrden
4440                                    oDoFacturacionBienesPagos.cantidadPagar = rsBusqueda.Fields!cantidad
4450                                    oDoFacturacionBienesPagos.idProducto = rsBusqueda.Fields!idProducto
4460                                    oDoFacturacionBienesPagos.IdUsuarioAuditoria = ml_IdUsuario
4470                                    oDoFacturacionBienesPagos.PrecioVenta = rsBusqueda.Fields!PrecioUnitario
4480                                    oDoFacturacionBienesPagos.TotalPagar = Round((rsBusqueda.Fields!cantidad * rsBusqueda.Fields!PrecioUnitario), 2) - lnDctos
4490                                    If Not oFacturacionBienesPagos.Insertar(oDoFacturacionBienesPagos) Then
4500                                       ms_MensajeError = oFacturacionBienesPagos.MensajeError: GoTo ErrAceptar
4510                                    End If
4520                                    lbHuboDetallePago = True
4530                               End If
4540                           Case 3   'SOAT
4550                               If rsBusqueda.Fields!CantidadSOAT > 0 Then
4560                                    oDoFacturacionBienesFinan.MovTipo = "S"
4570                                    oDoFacturacionBienesFinan.CantidadFinanciada = rsBusqueda.Fields!CantidadSOAT
4580                                    oDoFacturacionBienesFinan.FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
4590                                    oDoFacturacionBienesFinan.idProducto = rsBusqueda.Fields!idProducto
4600                                    oDoFacturacionBienesFinan.IdTipoFinanciamiento = ml_idUsuarioConPermisoEnSISoEXOoSOAT
4610                                    oDoFacturacionBienesFinan.IdUsuarioAuditoria = ml_IdUsuario
4620                                    oDoFacturacionBienesFinan.IdUsuarioAutoriza = ml_IdUsuario
4630                                    oDoFacturacionBienesFinan.PrecioFinanciado = rsBusqueda.Fields!precioSOAT
4640                                    oDoFacturacionBienesFinan.TotalFinanciado = rsBusqueda.Fields!ImporteSOAT
4650                                    oDoFacturacionBienesFinan.idFuenteFinanciamiento = ml_lnIdPlanActual
4660                                    oDoFacturacionBienesFinan.IdEstadoFacturacion = 1
                                        
4670                                    Set oRsTmp2 = mo_ReglasFacturacion.FacturacionBienesFinanciamFiltrarXmovNumeroProductoTFinanc(oDoFacturacionBienesFinan.MovNumero, _
                                                                        oDoFacturacionBienesFinan.idProducto, _
                                                                        oDoFacturacionBienesFinan.IdTipoFinanciamiento, _
                                                                        oDoFacturacionBienesFinan.MovTipo, oConexion)
4680                                    If oRsTmp2.RecordCount > 0 Then
4690                                          If Not oFacturacionBienesFinan.Modificar(oDoFacturacionBienesFinan) Then
4700                                             ms_MensajeError = oFacturacionBienesFinan.MensajeError: GoTo ErrAceptar
4710                                          End If
4720                                    Else
4730                                          If Not oFacturacionBienesFinan.Insertar(oDoFacturacionBienesFinan) Then
4740                                             ms_MensajeError = oFacturacionBienesFinan.MensajeError: GoTo ErrAceptar
4750                                          End If
4760                                    End If
4770                                    oRsTmp2.Close
4780                               End If
4790                               If rsBusqueda.Fields!cantidad > 0 Then
4800                                    oDoFacturacionBienesPagos.IdOrden = oDoFactOrdenesBienes.IdOrden
4810                                    oDoFacturacionBienesPagos.cantidadPagar = rsBusqueda.Fields!cantidad
4820                                    oDoFacturacionBienesPagos.idProducto = rsBusqueda.Fields!idProducto
4830                                    oDoFacturacionBienesPagos.IdUsuarioAuditoria = ml_IdUsuario
4840                                    oDoFacturacionBienesPagos.PrecioVenta = rsBusqueda.Fields!PrecioUnitario
4850                                    oDoFacturacionBienesPagos.TotalPagar = Round((rsBusqueda.Fields!cantidad * rsBusqueda.Fields!PrecioUnitario), 2) - lnDctos
4860                                    If Not oFacturacionBienesPagos.Insertar(oDoFacturacionBienesPagos) Then
4870                                       ms_MensajeError = oFacturacionBienesPagos.MensajeError: GoTo ErrAceptar
4880                                    End If
4890                                    lbHuboDetallePago = True
4900                               End If
4910                           Case 9   'EXONERACIONES
4920                               If rsBusqueda.Fields!importeEXO > 0 Then
4930                                    oDoFacturacionBienesFinan.MovTipo = "S"
4940                                    oDoFacturacionBienesFinan.CantidadFinanciada = rsBusqueda.Fields!cantidad
4950                                    oDoFacturacionBienesFinan.FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
4960                                    oDoFacturacionBienesFinan.idProducto = rsBusqueda.Fields!idProducto
4970                                    oDoFacturacionBienesFinan.IdTipoFinanciamiento = ml_idUsuarioConPermisoEnSISoEXOoSOAT
4980                                    oDoFacturacionBienesFinan.IdUsuarioAuditoria = ml_IdUsuario
4990                                    oDoFacturacionBienesFinan.IdUsuarioAutoriza = ml_IdUsuario
5000                                    oDoFacturacionBienesFinan.PrecioFinanciado = Round(rsBusqueda.Fields!importeEXO / rsBusqueda.Fields!cantidad, 2)
5010                                    oDoFacturacionBienesFinan.TotalFinanciado = rsBusqueda.Fields!importeEXO
5020                                    oDoFacturacionBienesFinan.idFuenteFinanciamiento = 0
5030                                    oDoFacturacionBienesFinan.IdEstadoFacturacion = 1
5040                                    Set oRsTmp2 = mo_ReglasFacturacion.FacturacionBienesFinanciamFiltrarXmovNumeroProductoTFinanc(oDoFacturacionBienesFinan.MovNumero, _
                                                                        oDoFacturacionBienesFinan.idProducto, _
                                                                        oDoFacturacionBienesFinan.IdTipoFinanciamiento, _
                                                                        oDoFacturacionBienesFinan.MovTipo, oConexion)
5050                                    If oRsTmp2.RecordCount > 0 Then
5060                                          If Not oFacturacionBienesFinan.Modificar(oDoFacturacionBienesFinan) Then
5070                                             ms_MensajeError = oFacturacionBienesFinan.MensajeError: GoTo ErrAceptar
5080                                          End If
5090                                    Else
5100                                          If Not oFacturacionBienesFinan.Insertar(oDoFacturacionBienesFinan) Then
5110                                             ms_MensajeError = oFacturacionBienesFinan.MensajeError: GoTo ErrAceptar
5120                                          End If
5130                                    End If
5140                                    oRsTmp2.Close
                                   
5150                               End If
5160                               If rsBusqueda.Fields!cantidad > 0 Then
5170                                    oDoFacturacionBienesPagos.IdOrden = oDoFactOrdenesBienes.IdOrden
5180                                    oDoFacturacionBienesPagos.cantidadPagar = rsBusqueda.Fields!cantidad
5190                                    oDoFacturacionBienesPagos.idProducto = rsBusqueda.Fields!idProducto
5200                                    oDoFacturacionBienesPagos.IdUsuarioAuditoria = ml_IdUsuario
5210                                    oDoFacturacionBienesPagos.PrecioVenta = rsBusqueda.Fields!PrecioUnitario
5220                                    oDoFacturacionBienesPagos.TotalPagar = Round((rsBusqueda.Fields!cantidad * rsBusqueda.Fields!PrecioUnitario), 2) - lnDctos
5230                                    If Not oFacturacionBienesPagos.Insertar(oDoFacturacionBienesPagos) Then
5240                                       ms_MensajeError = oFacturacionBienesPagos.MensajeError: GoTo ErrAceptar
5250                                    End If
5260                                    lbHuboDetallePago = True
5270                               End If
5280                           Case 4   'Convenios
5290                               If rsBusqueda.Fields!esConvenio = "Si" Then
5300                                    If rsBusqueda.Fields!CantidadConv > 0 Then
5310                                         oDoFacturacionBienesFinan.MovTipo = "S"
5320                                         oDoFacturacionBienesFinan.CantidadFinanciada = rsBusqueda.Fields!CantidadConv
5330                                         oDoFacturacionBienesFinan.FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
5340                                         oDoFacturacionBienesFinan.idProducto = rsBusqueda.Fields!idProducto
5350                                         oDoFacturacionBienesFinan.IdTipoFinanciamiento = ml_idUsuarioConPermisoEnSISoEXOoSOAT
5360                                         oDoFacturacionBienesFinan.IdUsuarioAuditoria = ml_IdUsuario
5370                                         oDoFacturacionBienesFinan.IdUsuarioAutoriza = ml_IdUsuario
5380                                         oDoFacturacionBienesFinan.PrecioFinanciado = rsBusqueda.Fields!PrecioConv
5390                                         oDoFacturacionBienesFinan.TotalFinanciado = Round(rsBusqueda.Fields!CantidadConv * rsBusqueda.Fields!PrecioConv, 2)
5400                                         oDoFacturacionBienesFinan.idFuenteFinanciamiento = ml_lnIdPlanActual
5410                                         oDoFacturacionBienesFinan.IdEstadoFacturacion = 1
                                             
                                             
5420                                         Set oRsTmp2 = mo_ReglasFacturacion.FacturacionBienesFinanciamFiltrarXmovNumeroProductoTFinanc(oDoFacturacionBienesFinan.MovNumero, _
                                                                        oDoFacturacionBienesFinan.idProducto, _
                                                                        oDoFacturacionBienesFinan.IdTipoFinanciamiento, _
                                                                        oDoFacturacionBienesFinan.MovTipo, oConexion)
5430                                          If oRsTmp2.RecordCount > 0 Then
5440                                                If Not oFacturacionBienesFinan.Modificar(oDoFacturacionBienesFinan) Then
5450                                                   ms_MensajeError = oFacturacionBienesFinan.MensajeError: GoTo ErrAceptar
5460                                                End If
5470                                          Else
5480                                                If Not oFacturacionBienesFinan.Insertar(oDoFacturacionBienesFinan) Then
5490                                                   ms_MensajeError = oFacturacionBienesFinan.MensajeError: GoTo ErrAceptar
5500                                                End If
5510                                          End If
5520                                          oRsTmp2.Close
5530                                    End If
5540                               Else
5550                                    If rsBusqueda.Fields!cantidad > 0 Then
5560                                         oDoFacturacionBienesPagos.IdOrden = oDoFactOrdenesBienes.IdOrden
5570                                         oDoFacturacionBienesPagos.cantidadPagar = rsBusqueda.Fields!cantidad
5580                                         oDoFacturacionBienesPagos.idProducto = rsBusqueda.Fields!idProducto
5590                                         oDoFacturacionBienesPagos.IdUsuarioAuditoria = ml_IdUsuario
5600                                         oDoFacturacionBienesPagos.PrecioVenta = rsBusqueda.Fields!PrecioUnitario
5610                                         oDoFacturacionBienesPagos.TotalPagar = Round(rsBusqueda.Fields!cantidad * rsBusqueda.Fields!PrecioUnitario, 2)
5620                                         If Not oFacturacionBienesPagos.Insertar(oDoFacturacionBienesPagos) Then
5630                                            ms_MensajeError = oFacturacionBienesPagos.MensajeError: GoTo ErrAceptar
5640                                         End If
5650                                         lbHuboDetallePago = True
5660                                    End If
5670                               End If
5680                           End Select
5690                     End If
5700                  End If
5710                  If lbHuboDetallePago = True Then
5720                      oDoFactOrdenesBienes.MovNumero = lcMovNumero
5730                      oDoFactOrdenesBienes.MovTipo = "S"
5740                      If Not oFactOrdenesBienes.Modificar(oDoFactOrdenesBienes) Then
5750                          ms_MensajeError = oFactOrdenesBienes.MensajeError: GoTo ErrAceptar
5760                      End If
5770                  End If
5780                  rsBusqueda.MoveNext
5790                  If rsBusqueda.EOF Then
5800                     Exit Do
5810                  End If
5820            Loop
5830         Loop
5840      End If
5850      rsBusqueda.Close
5860      Set rsBusqueda = Nothing
5870      Set rsActualizar = Nothing
5880      Set oRsTmp3 = Nothing
          
          '
5890      Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_IdUsuario, "M", lnIdCuentaAtencion, lcAccion, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPlanActual)        'ListBarItems.idListItem
5900      oConexion.CommitTrans
5910      oConexion.Close
5920      Set oConexion = Nothing
5930      Set mo_ReglasComunes = Nothing
         
5940      Exit Sub
ErrAceptar:
      'Resume
5950      oConexion.RollbackTrans
      'Resume
5960      If Err.Description <> "" Then
5970          ms_MensajeError = ms_MensajeError + Chr(13) + Err.Description & _
                                sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                "Sub GrabaCantidadesPreciosRegistradosEnSisSoatExoConvenio", "dllFactUcEstadoCuenta.cls")  'debb-02/05/2016
                  
5980      End If
5990      MsgBox ms_MensajeError
6000      oConexion.Close
6010      Set oConexion = Nothing
End Sub



Sub GenerarRecordsetProductos(ByRef mrs_FacturacionProductos As ADODB.Recordset)
    If mrs_FacturacionProductos.State = adStateOpen Then
       mrs_FacturacionProductos.Close
    End If
    With mrs_FacturacionProductos
          .Fields.Append "IdFacturacionProducto", adInteger
          .Fields.Append "IdProducto", adInteger
          .Fields.Append "FechaOrden", adDBTimeStamp, , adFldIsNullable
          .Fields.Append "IdPuntoCarga", adInteger
          .Fields.Append "NroDcto", adVarChar, 20, adFldIsNullable
          .Fields.Append "Codigo", adVarChar, 255, adFldIsNullable
          .Fields.Append "NombreProducto", adVarChar, 255, adFldIsNullable
          .Fields.Append "CantidadPagar", adInteger
          .Fields.Append "PrecioUnitario", adCurrency
          .Fields.Append "TotalPagar", adCurrency
          .Fields.Append "CantidadSIS", adCurrency
          .Fields.Append "PrecioSIS", adCurrency
          .Fields.Append "ImporteSIS", adCurrency
          .Fields.Append "CantidadSOAT", adCurrency
          .Fields.Append "PrecioSOAT", adCurrency
          .Fields.Append "ImporteSOAT", adCurrency
          .Fields.Append "EsConvenio", adVarChar, 2, adFldIsNullable
          .Fields.Append "CantidadConv", adCurrency
          .Fields.Append "PrecioConv", adCurrency
          .Fields.Append "ImporteConv", adCurrency
          .Fields.Append "ImporteEXO", adCurrency
          .Fields.Append "Cantidad", adCurrency
          .Fields.Append "TotalPorPagar", adCurrency
          .Fields.Append "CantidadDevuelta", adInteger
          .Fields.Append "IdEstadoFacturacion", adInteger
          .Fields.Append "DocReembolso", adVarChar, 20, adFldIsNullable
          .Fields.Append "IdServicioInternamiento", adInteger
          .Fields.Append "NombreServicio", adVarChar, 255
          .Fields.Append "IdAtencion", adInteger, , adFldIsNullable
          .Fields.Append "IdOrden", adInteger
          .Fields.Append "IdUsuarioAutorizaSeguro", adInteger, , adFldIsNullable
          .Fields.Append "FechaAutorizaSeguro", adDBTimeStamp, , adFldIsNullable
          .Fields.Append "IdUsuarioAutorizaPendiente", adInteger, , adFldIsNullable
          .Fields.Append "FechaAutorizaPendiente", adDBTimeStamp, , adFldIsNullable
          .Fields.Append "IdUsuarioAutorizaDevolucion", adInteger, , adFldIsNullable
          .Fields.Append "FechaAutorizaDevolucion", adDBTimeStamp, , adFldIsNullable
          .Fields.Append "IdComprobantePago", adInteger, , adFldIsNullable
          .Fields.Append "NroComprobante", adVarChar, 255
          .Fields.Append "IdComprobantePagoDevolucion", adInteger
          .Fields.Append "IdTipoFinanciamiento", adInteger
          .Fields.Append "IdFuenteFinanciamiento", adInteger, , adFldIsNullable
          .Fields.Append "Descripcion", adVarChar, 255, adFldIsNullable
          .Fields.Append "MovNumero", adVarChar, 9, adFldIsNullable
          .Fields.Append "MovTipo", adVarChar, 1, adFldIsNullable
          .Fields.Append "IdTipoConceptoFarmacia", adInteger
          .Fields.Append "idServicioDeEstancia", adInteger
          .Fields.Append "ServicioDeEstancia", adVarChar, 255, adFldIsNullable
          .Fields.Append "NroDocumento", adVarChar, 20, adFldIsNullable
          .Fields.Append "ImporteEnBoleta", adCurrency
          .Fields.Append "ComoSeTrabajaEnEstadoCuenta", adInteger
          .Fields.Append "IdOrdenPago", adInteger
          .Fields.Append "FechaDespacho", adDBTimeStamp, , adFldIsNullable
          .Fields.Append "CantidadEstancia", adCurrency
          .Fields.Append "PrecioEstancia", adCurrency
          .Fields.Append "EsPaquete", adBoolean
          .Fields.Append "Receta", adInteger
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
End Sub

Sub GrabaCantidadesPreciosEnElNuevoPlan(lnIdCuentaAtencion As Long, lnIdPlanNuevo As Long, _
                                        lnIdTipoFinanciamientoNuevo As Long, ml_IdUsuario As Long, _
                                        ml_IdAtencion As Long, mo_lnIdTablaLISTBARITEMS As Long, _
                                        mo_lcNombrePc As String, lcPlanActual As String, _
                                        lnIdTipoFinanciamientoActual As Long, lbNuevoPlanConPreciosAntPlan As Boolean, _
                                        ldFechaAfiliacionSIS As Date, lcSIScodigo As String, _
                                        lnIdSiaSis As Long, lcSISafiliacion As String)
          Dim oFacturacionServicioFinan As New FacturacionServicioFin
          Dim oDoFacturacionServicioFinanc As New DoFacturacionServicioFin
          Dim oDoFactOrdenServicioPagos As New DoFactOrdenServPagos
          Dim oFactServicioPagos As New FactOrdenServicioPagos
          Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
          Dim oFacturacionServicioPagos As New FacturacionServicioPagos
          '
          Dim oFactOrdenesBienes As New FactOrdenesBienes
          Dim oDoFactOrdenesBienes As New DoFactOrdenesBienes
          Dim oFacturacionBienesFinan As New FacturacionBienesFinanciam
          Dim oDoFacturacionBienesFinan As New DoFacturacionBienesFinanciam
          Dim oFacturacionBienesPagos As New FacturacionBienesPagos
          Dim oDoFacturacionBienesPagos As New DoFacturacionBienesPagos
          '
          Dim oAtenciones As New Atenciones, oDoAtencion As New DOAtencion
          Dim oAtencionesDatosAdicionales As New AtencionesDatosAdicionales, oDoAtencionDatosAdicionales As New DoAtencionDatosAdicionales
          '
          Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
          Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
          Dim oRsTmp1 As New Recordset
          Dim oRsTmp2 As New Recordset
          Dim oRsTmp3 As New Recordset
          Dim oRsOrdenesPago As New Recordset
          Dim oRsDevoluciones As New Recordset
          Dim oConexion As New ADODB.Connection
10        Dim lnIdOrdenPago As Long: Dim lnIdOrden As Long
20        Dim lbCreaCabeceraPagosParaSegurosQueNoCubre  As Boolean: Dim lbCreaDetallePagos As Boolean
30        Dim lbGeneraReciboPago As Boolean: Dim lbNuevo As Boolean
          Dim lcMovNumero As String
          Dim lcMovTipo As String
          Dim lnDetallePago As Integer
          Dim lcIdServiciosEstanciaEnParametros As String
          Dim lnPrecioUnitarioServicio As Double
          Dim lnDevoluciones As Long
          Dim lnPrecioUnitarioServicio1 As Double, lbEliminaOrden As Boolean

40        On Error GoTo ErrRecalculo
          '
50        lcIdServiciosEstanciaEnParametros = lcBuscaParametro.SeleccionaFilaParametro(202) & "."
60        lcIdServiciosEstanciaEnParametros = lcIdServiciosEstanciaEnParametros & lcBuscaParametro.SeleccionaFilaParametro(203) & "."
70        lcIdServiciosEstanciaEnParametros = lcIdServiciosEstanciaEnParametros & lcBuscaParametro.SeleccionaFilaParametro(204)
          
          '
80        With oRsOrdenesPago
90              .Fields.Append "IdOrden", adInteger
100             .Fields.Append "MovNumero", adVarChar, 9, adFldIsNullable
110             .CursorType = adOpenDynamic
120             .LockType = adLockOptimistic
130             .Open
140       End With
          '
150       oConexion.Open sighentidades.CadenaConexion
160       oConexion.CursorLocation = adUseClient
170       oConexion.BeginTrans
          '
180       Set oFacturacionServicioFinan.Conexion = oConexion
190       Set oFactServicioPagos.Conexion = oConexion
200       Set oFacturacionServicioPagos.Conexion = oConexion
          '
210       Set oFactOrdenesBienes.Conexion = oConexion
220       Set oFacturacionBienesPagos.Conexion = oConexion
230       Set oFacturacionBienesFinan.Conexion = oConexion
          '
240       Set oAtenciones.Conexion = oConexion
          Set oAtencionesDatosAdicionales.Conexion = oConexion
          '
250       lbGeneraReciboPago = ml_ReglasFacturacion.TiposFinanciamientoGeneraReciboPago(lnIdTipoFinanciamientoNuevo, oConexion)
          '************************* SERVICIOS  ******************************************
260       lbCreaCabeceraPagosParaSegurosQueNoCubre = lbGeneraReciboPago
          'Anula los Pagos Pendientes de Pagos -servicios
270       Set oRsTmp1 = mo_ReglasFacturacion.FactOrdenServicioPagosPorIdCuentaAtencion(lnIdCuentaAtencion)
280       oRsTmp1.Filter = "IdEstadoFacturacion<>1"
290       If oRsTmp1.RecordCount > 0 Then
300          oRsTmp1.MoveFirst
310          Do While Not oRsTmp1.EOF
320             lbNuevo = True
330             If oRsOrdenesPago.RecordCount > 0 Then
340                oRsOrdenesPago.MoveFirst
350                oRsOrdenesPago.Find "idOrden=" & oRsTmp1.Fields!IdOrden
360                If Not oRsOrdenesPago.EOF Then
370                   lbNuevo = False
380                End If
390             End If
400             If lbNuevo = True Then
410                oRsOrdenesPago.AddNew
420                oRsOrdenesPago.Fields!IdOrden = oRsTmp1.Fields!IdOrden
430                oRsOrdenesPago.Update
440             End If
450             oRsTmp1.MoveNext
460          Loop
470       End If
480       oRsTmp1.Filter = "IdEstadoFacturacion=1"
490       If oRsTmp1.RecordCount > 0 Then
500          oRsTmp1.MoveFirst
510          Do While Not oRsTmp1.EOF
                'es un PAGANTE que pasa a ser SIS, solo debe considerar CONSUMOS A PARTIR DE SU FECHA DE AFILIACION
                lbEliminaOrden = True
                If ldFechaAfiliacionSIS > 0 Then
                   If oRsTmp1!fechaDespacho < ldFechaAfiliacionSIS Then
                        lbEliminaOrden = False
                        lbNuevo = True
                        If oRsOrdenesPago.RecordCount > 0 Then
                            oRsOrdenesPago.MoveFirst
                            oRsOrdenesPago.Find "idOrden=" & oRsTmp1.Fields!IdOrden
                            If Not oRsOrdenesPago.EOF Then
                               lbNuevo = False
                            End If
                         End If
                         If lbNuevo = True Then
                            oRsOrdenesPago.AddNew
                            oRsOrdenesPago.Fields!IdOrden = oRsTmp1.Fields!IdOrden
                            oRsOrdenesPago.Update
                         End If
                   End If
                End If
                '
                If lbEliminaOrden = True Then
520                     lnIdOrdenPago = oRsTmp1.Fields!idOrdenPago
530                     oDoFactOrdenServicioPagos.idOrdenPago = oRsTmp1.Fields!idOrdenPago
540                     If Not oFactServicioPagos.SeleccionarPorId(oDoFactOrdenServicioPagos) Then
550                        ms_MensajeError = oFactServicioPagos.MensajeError: GoTo ErrRecalculo
560                     End If
570                     oDoFactOrdenServicioPagos.IdUsuarioAuditoria = ml_IdUsuario
580                     oDoFactOrdenServicioPagos.IdEstadoFacturacion = 9   'anulado
590                     oDoFactOrdenServicioPagos.IdOrden = 0
600                     If Not oFactServicioPagos.Modificar(oDoFactOrdenServicioPagos) Then
610                        ms_MensajeError = oFactServicioPagos.MensajeError: GoTo ErrRecalculo
620                     End If
                End If
630             oRsTmp1.MoveNext
640          Loop
650       End If
          'Elimina datos de FINANCIAMIENTOS - servicios
660       Set oRsTmp1 = mo_ReglasFacturacion.FactOrdenServicioFiltraPorIdCuenta(lnIdCuentaAtencion)
670       oRsTmp1.Filter = "idEstadoFacturacion=1"
680       If oRsTmp1.RecordCount > 0 Then
690          oRsTmp1.MoveFirst
700          Do While Not oRsTmp1.EOF
710             lnIdOrden = oRsTmp1.Fields!IdOrden
720             oDoFacturacionServicioFinanc.IdOrden = lnIdOrden
730             oDoFacturacionServicioFinanc.IdUsuarioAuditoria = ml_IdUsuario
740             If Not oFacturacionServicioFinan.Eliminar(oDoFacturacionServicioFinanc) Then
750                ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo ErrRecalculo
760             End If
770             Do While Not oRsTmp1.EOF And lnIdOrden = oRsTmp1.Fields!IdOrden
780                If lbGeneraReciboPago = False Then
790                   Set oRsTmp2 = mo_ReglasFacturacion.CatalogoServicioSeleccionarPorIdYtipoFinanciamiento(oRsTmp1.Fields!idProducto, lnIdTipoFinanciamientoNuevo)
800                   If oRsTmp2.RecordCount = 0 Then lbCreaCabeceraPagosParaSegurosQueNoCubre = True
810                End If
820                oRsTmp1.MoveNext
830                If oRsTmp1.EOF Then
840                   Exit Do
850                End If
860             Loop
870          Loop
880       End If
          'Agrega nuevos Pagos y Financiamientos -servicios
890       If oRsTmp1.RecordCount > 0 Then
900          oRsTmp1.MoveFirst
910          Do While Not oRsTmp1.EOF
920             lnIdOrden = oRsTmp1.Fields!IdOrden
930             lbNuevo = False
940             If oRsOrdenesPago.RecordCount > 0 Then
950                oRsOrdenesPago.MoveFirst
960                oRsOrdenesPago.Find "idOrden=" & oRsTmp1.Fields!IdOrden
970                If Not oRsOrdenesPago.EOF Then
980                   lbNuevo = True
990                End If
1000            End If
1010            If lbNuevo = True Then
                   'esa orden de pago ya fue pagada o anulada
1020               Do While Not oRsTmp1.EOF And lnIdOrden = oRsTmp1.Fields!IdOrden
1030                  oRsTmp1.MoveNext
1040                  If oRsTmp1.EOF Then
1050                     Exit Do
1060                  End If
1070               Loop
1080            Else
1090              If lbCreaCabeceraPagosParaSegurosQueNoCubre = True Then
1100                   oDoFactOrdenServicioPagos.fechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
                       'oDoFactOrdenServicioPagos.idComprobantePago
1110                   oDoFactOrdenServicioPagos.IdEstadoFacturacion = 1 'registrado
1120                   oDoFactOrdenServicioPagos.IdOrden = lnIdOrden
                       'oDoFactOrdenServicioPagos.idOrdenPago
1130                   oDoFactOrdenServicioPagos.idUsuario = ml_IdUsuario
1140                   oDoFactOrdenServicioPagos.IdUsuarioAuditoria = ml_IdUsuario
1150                   If Not oFactServicioPagos.Insertar(oDoFactOrdenServicioPagos) Then
1160                       ms_MensajeError = oFactServicioPagos.MensajeError: GoTo ErrRecalculo
1170                   End If
1180              End If
                  
                  
1190              lnDetallePago = 0
1200              Do While Not oRsTmp1.EOF And lnIdOrden = oRsTmp1.Fields!IdOrden
1210                 lbCreaDetallePagos = False
1220                 If lbGeneraReciboPago = False Then
1230                    lnPrecioUnitarioServicio = -999
1240                    Set oRsTmp2 = mo_ReglasFacturacion.CatalogoServicioSeleccionarPorIdYtipoFinanciamiento(oRsTmp1.Fields!idProducto, lnIdTipoFinanciamientoNuevo)
1250                    If oRsTmp2.RecordCount > 0 Then
1260                          lnPrecioUnitarioServicio = oRsTmp2.Fields!PrecioUnitario
1270                    Else
1280                    End If
1290                    If lnPrecioUnitarioServicio <> -999 Then
                           'estancia en Horas
1300                       If lnPrecioUnitarioServicio > 0 And oRsTmp1.Fields!idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaAdmisionHospitalizacion Then
1310                          lnPrecioUnitarioServicio = Round(lnPrecioUnitarioServicio / 24, 4)
1320                       End If
                           '
1330                       oDoFacturacionServicioFinanc.CantidadFinanciada = oRsTmp1.Fields!cantidad
1340                       oDoFacturacionServicioFinanc.FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
1350                       oDoFacturacionServicioFinanc.idProducto = oRsTmp1.Fields!idProducto
1360                       oDoFacturacionServicioFinanc.IdTipoFinanciamiento = lnIdTipoFinanciamientoNuevo
1370                       oDoFacturacionServicioFinanc.IdUsuarioAuditoria = ml_IdUsuario
1380                       oDoFacturacionServicioFinanc.IdUsuarioAutoriza = ml_IdUsuario
1390                       oDoFacturacionServicioFinanc.PrecioFinanciado = lnPrecioUnitarioServicio
1400                       oDoFacturacionServicioFinanc.TotalFinanciado = Round(oRsTmp1.Fields!cantidad * lnPrecioUnitarioServicio, 2)
1410                       oDoFacturacionServicioFinanc.idFuenteFinanciamiento = lnIdPlanNuevo
1420                       oDoFacturacionServicioFinanc.IdOrden = lnIdOrden
1430                       If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then
1440                          ms_MensajeError = oFacturacionServicioFinan.MensajeError: GoTo ErrRecalculo
1450                       End If
1460                    Else
1470                       lbCreaDetallePagos = True
1480                    End If
1490                 End If
1500                 If lbGeneraReciboPago = True Or lbCreaDetallePagos = True Then
1510                    If lnIdTipoFinanciamientoActual = 3 And lnIdTipoFinanciamientoNuevo = 1 And lbNuevoPlanConPreciosAntPlan = True Then  'soat/particular
                           '******solo cuando es SOAT y pasa hacia PARTICULAR, -->solo si tiene el Check de "Pasa de Soat hacia PARTICULAR con Precios SOAT"
                           '******toma precios SOAT
1520                       Set oRsTmp2 = mo_ReglasFacturacion.CatalogoServicioSeleccionarPorIdYtipoFinanciamiento(oRsTmp1.Fields!idProducto, lnIdTipoFinanciamientoActual)
1530                    Else
1540                       Set oRsTmp2 = mo_ReglasFacturacion.CatalogoServicioSeleccionarPorIdYtipoFinanciamiento(oRsTmp1.Fields!idProducto, lnIdTipoFinanciamientoNuevo)
1550                    End If
1560                    If oRsTmp2.RecordCount > 0 Then
1570                       lnPrecioUnitarioServicio1 = oRsTmp2.Fields!PrecioUnitario
                           'estancia en Horas
1580                       If lnPrecioUnitarioServicio1 > 0 And oRsTmp1.Fields!idPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaAdmisionHospitalizacion Then
1590                          lnPrecioUnitarioServicio1 = Round(oRsTmp2.Fields!PrecioUnitario / 24, 4)
1600                       End If
                           '
1610                       oDoFacturacionServicioPagos.cantidad = oRsTmp1.Fields!cantidad
1620                       oDoFacturacionServicioPagos.idProducto = oRsTmp1.Fields!idProducto
1630                       oDoFacturacionServicioPagos.IdUsuarioAuditoria = ml_IdUsuario
1640                       oDoFacturacionServicioPagos.Precio = lnPrecioUnitarioServicio1
1650                       oDoFacturacionServicioPagos.total = Round((oRsTmp1.Fields!cantidad * lnPrecioUnitarioServicio1), 2)
1660                       oDoFacturacionServicioPagos.idOrdenPago = oDoFactOrdenServicioPagos.idOrdenPago
1670                       If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
1680                          ms_MensajeError = oFacturacionServicioPagos.MensajeError: GoTo ErrRecalculo
1690                       End If
1700                       lnDetallePago = lnDetallePago + 1
1710                    End If
1720                 End If
1730                 oRsTmp1.MoveNext
1740                 If oRsTmp1.EOF Then
1750                    Exit Do
1760                 End If
1770              Loop
1780              If lbGeneraReciboPago = False And lbCreaCabeceraPagosParaSegurosQueNoCubre = True And lnDetallePago = 0 Then
1790                   If Not oFactServicioPagos.Eliminar(oDoFactOrdenServicioPagos) Then
1800                       ms_MensajeError = oFactServicioPagos.MensajeError: GoTo ErrRecalculo
1810                   End If
1820              End If
1830           End If
1840         Loop
1850      End If
          
          '************************* FARMACIA   ******************************************
          
1860      lbCreaCabeceraPagosParaSegurosQueNoCubre = lbGeneraReciboPago
          'Elimina datos de FINANCIAMIENTOS y PENDIENTES DE PAGO - farmacia
1870      Set oRsTmp1 = mo_ReglasFarmacia.farmMovimientoVentasSeleccionarPorIdCuentaAtencion(lnIdCuentaAtencion)
1880      oRsTmp1.Filter = "idEstadoMovimiento=1"
1890      If oRsTmp1.RecordCount > 0 Then
1900         oRsTmp1.MoveFirst
1910         Do While Not oRsTmp1.EOF
1920            oDoFactOrdenesBienes.MovNumero = oRsTmp1.Fields!MovNumero
1930            oDoFactOrdenesBienes.MovTipo = oRsTmp1.Fields!MovTipo
1940            If Not oFactOrdenesBienes.SeleccionarPorMovNumero(oDoFactOrdenesBienes) Then
1950               ms_MensajeError = oFactOrdenesBienes.MensajeError: GoTo ErrRecalculo
1960            End If
1970            If oDoFactOrdenesBienes.IdComprobantePago > 0 Then
1980                  lbNuevo = True
1990                  If oRsOrdenesPago.RecordCount > 0 Then
2000                     oRsOrdenesPago.MoveFirst
2010                     oRsOrdenesPago.Find "movNumero='" & oRsTmp1.Fields!MovNumero & "'"
2020                     If Not oRsOrdenesPago.EOF Then
2030                        lbNuevo = False
2040                     End If
2050                  End If
2060                  If lbNuevo = True Then
2070                     oRsOrdenesPago.AddNew
2080                     oRsOrdenesPago.Fields!MovNumero = oRsTmp1.Fields!MovNumero
2090                     oRsOrdenesPago.Update
2100                  End If
2110            Else
                    'es un PAGANTE que pasa a ser SIS, solo debe considerar CONSUMOS A PARTIR DE SU FECHA DE AFILIACION
                    lbEliminaOrden = True
                    If ldFechaAfiliacionSIS > 0 Then
                        If oRsTmp1!fechaCreacion < ldFechaAfiliacionSIS Then
                            lbEliminaOrden = False
                            lbNuevo = True
                            If oRsOrdenesPago.RecordCount > 0 Then
                                oRsOrdenesPago.MoveFirst
                                oRsOrdenesPago.Find "movNumero='" & oRsTmp1.Fields!MovNumero & "'"
                                If Not oRsOrdenesPago.EOF Then
                                    lbNuevo = False
                                End If
                            End If
                            If lbNuevo = True Then
                                oRsOrdenesPago.AddNew
                                oRsOrdenesPago.Fields!MovNumero = oRsTmp1.Fields!MovNumero
                                oRsOrdenesPago.Update
                            End If
                        End If
                    End If
                    '
                    If lbEliminaOrden = True Then
                        oDoFactOrdenesBienes.IdUsuarioAuditoria = ml_IdUsuario
                        oDoFactOrdenesBienes.MovNumero = ""
                        oDoFactOrdenesBienes.IdEstadoFacturacion = 9
                        If Not oFactOrdenesBienes.Modificar(oDoFactOrdenesBienes) Then
                            ms_MensajeError = oFactOrdenesBienes.MensajeError: GoTo ErrRecalculo
                        End If
                    End If
2180            End If
2190            oDoFacturacionBienesFinan.MovNumero = oRsTmp1.Fields!MovNumero
2200            oDoFacturacionBienesFinan.MovTipo = oRsTmp1.Fields!MovTipo
2210            oDoFacturacionBienesFinan.IdUsuarioAuditoria = ml_IdUsuario
2220            If Not oFacturacionBienesFinan.Eliminar(oDoFacturacionBienesFinan) Then
2230                 ms_MensajeError = oFacturacionBienesFinan.MensajeError: GoTo ErrRecalculo
2240            End If
2250            lcMovNumero = oRsTmp1.Fields!MovNumero
2260            lcMovTipo = oRsTmp1.Fields!MovTipo
2270            Do While Not oRsTmp1.EOF And lcMovNumero = oRsTmp1.Fields!MovNumero And lcMovTipo = oRsTmp1.Fields!MovTipo
2280               If lbGeneraReciboPago = False Then
2290                  Set oRsTmp2 = mo_ReglasFacturacion.CatalogoBienesSeleccionarPorIdYtipoFinanciamiento(oRsTmp1.Fields!idProducto, lnIdTipoFinanciamientoNuevo)
2300                  If oRsTmp2.RecordCount = 0 Then lbCreaCabeceraPagosParaSegurosQueNoCubre = True
2310               End If
2320               oRsTmp1.MoveNext
2330               If oRsTmp1.EOF Then
2340                  Exit Do
2350               End If
2360            Loop
2370         Loop
2380      End If
          'Agrega nuevos Pagos y Financiamientos -farmacia
2390      If oRsTmp1.RecordCount > 0 Then
2400         oRsTmp1.MoveFirst
2410         Do While Not oRsTmp1.EOF
2420            lcMovNumero = oRsTmp1.Fields!MovNumero
2430            lcMovTipo = oRsTmp1.Fields!MovTipo
2440            lbNuevo = False
2450            If oRsOrdenesPago.RecordCount > 0 Then
2460               oRsOrdenesPago.MoveFirst
2470               oRsOrdenesPago.Find "movNumero='" & oRsTmp1.Fields!MovNumero & "'"
2480               If Not oRsOrdenesPago.EOF Then
2490                  lbNuevo = True
2500               End If
2510            End If
2520            If lbNuevo = True Then
                  'esa orden de pago ya fue pagada o anulada
2530              Do While Not oRsTmp1.EOF And lcMovNumero = oRsTmp1.Fields!MovNumero And lcMovTipo = oRsTmp1.Fields!MovTipo
2540                 oRsTmp1.MoveNext
2550                 If oRsTmp1.EOF Then
2560                    Exit Do
2570                 End If
2580              Loop
2590            Else
2600              If lbCreaCabeceraPagosParaSegurosQueNoCubre = True Then
2610                    oDoFactOrdenesBienes.fechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
2620                    oDoFactOrdenesBienes.IdCuentaAtencion = lnIdCuentaAtencion
2630                    oDoFactOrdenesBienes.IdEstadoFacturacion = 1   'registrado
2640                    oDoFactOrdenesBienes.idPaciente = oRsTmp1.Fields!idPaciente
2650                    oDoFactOrdenesBienes.idPreVenta = 0
2660                    oDoFactOrdenesBienes.idPuntoCarga = 5
2670                    oDoFactOrdenesBienes.idUsuario = ml_IdUsuario
2680                    oDoFactOrdenesBienes.IdUsuarioAuditoria = ml_IdUsuario
2690                    oDoFactOrdenesBienes.MovNumero = oRsTmp1.Fields!MovNumero
2700                    oDoFactOrdenesBienes.MovTipo = oRsTmp1.Fields!MovTipo
2710                    If Not oFactOrdenesBienes.Insertar(oDoFactOrdenesBienes) Then
2720                        ms_MensajeError = oFactOrdenesBienes.MensajeError: GoTo ErrRecalculo
2730                    End If
2740              End If
2750              lnDetallePago = 0
2760              Do While Not oRsTmp1.EOF And lcMovNumero = oRsTmp1.Fields!MovNumero And lcMovTipo = oRsTmp1.Fields!MovTipo
                     'Devoluciones
2770                 lnDevoluciones = 0
2780                 Set oRsDevoluciones = mo_ReglasFacturacion.FacturacionBienesDevolucionesSeleccionarPorIdProducto(oRsTmp1.Fields!MovNumero, oRsTmp1.Fields!MovTipo, oRsTmp1.Fields!idProducto, oConexion)
2790                 If oRsDevoluciones.RecordCount > 0 Then
2800                    oRsDevoluciones.MoveFirst
2810                    Do While Not oRsDevoluciones.EOF
2820                       lnDevoluciones = lnDevoluciones + oRsDevoluciones.Fields!cantidadAdevolver
2830                       oRsDevoluciones.MoveNext
2840                    Loop
2850                 End If
2860                 oRsDevoluciones.Close
                     '
2870                 lbCreaDetallePagos = False
2880                 If lbGeneraReciboPago = False Then
2890                    Set oRsTmp2 = mo_ReglasFacturacion.CatalogoBienesSeleccionarPorIdYtipoFinanciamiento(oRsTmp1.Fields!idProducto, lnIdTipoFinanciamientoNuevo)
2900                    If oRsTmp2.RecordCount > 0 Then
2910                       oDoFacturacionBienesFinan.CantidadFinanciada = oRsTmp1.Fields!cantidad - lnDevoluciones
2920                       oDoFacturacionBienesFinan.FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
2930                       oDoFacturacionBienesFinan.idProducto = oRsTmp1.Fields!idProducto
2940                       oDoFacturacionBienesFinan.IdTipoFinanciamiento = lnIdTipoFinanciamientoNuevo
2950                       oDoFacturacionBienesFinan.IdUsuarioAuditoria = ml_IdUsuario
2960                       oDoFacturacionBienesFinan.IdUsuarioAutoriza = ml_IdUsuario
2970                       oDoFacturacionBienesFinan.PrecioFinanciado = oRsTmp2.Fields!PrecioUnitario
2980                       oDoFacturacionBienesFinan.TotalFinanciado = Round(oDoFacturacionBienesFinan.CantidadFinanciada * oRsTmp2.Fields!PrecioUnitario, 2)
2990                       oDoFacturacionBienesFinan.idFuenteFinanciamiento = lnIdPlanNuevo
3000                       oDoFacturacionBienesFinan.MovNumero = oRsTmp1.Fields!MovNumero
3010                       oDoFacturacionBienesFinan.MovTipo = oRsTmp1.Fields!MovTipo
3020                       If Not oFacturacionBienesFinan.Insertar(oDoFacturacionBienesFinan) Then
                              
3030                          ms_MensajeError = oFacturacionBienesFinan.MensajeError: GoTo ErrRecalculo
3040                       End If
3050                    Else
3060                       lbCreaDetallePagos = True
3070                    End If
3080                 End If
3090                 If lbGeneraReciboPago = True Or lbCreaDetallePagos = True Then
3100                    Set oRsTmp2 = mo_ReglasFacturacion.CatalogoBienesSeleccionarPorIdYtipoFinanciamiento(oRsTmp1.Fields!idProducto, lnIdTipoFinanciamientoNuevo)
3110                    If oRsTmp2.RecordCount > 0 Then
3120                       oDoFacturacionBienesPagos.cantidadPagar = oRsTmp1.Fields!cantidad - lnDevoluciones
3130                       oDoFacturacionBienesPagos.idProducto = oRsTmp1.Fields!idProducto
3140                       oDoFacturacionBienesPagos.IdUsuarioAuditoria = ml_IdUsuario
3150                       oDoFacturacionBienesPagos.PrecioVenta = oRsTmp2.Fields!PrecioUnitario
3160                       oDoFacturacionBienesPagos.TotalPagar = Round((oDoFacturacionBienesPagos.cantidadPagar * oRsTmp2.Fields!PrecioUnitario), 2)
3170                       oDoFacturacionBienesPagos.IdOrden = oDoFactOrdenesBienes.IdOrden
3180                       If Not oFacturacionBienesPagos.Insertar(oDoFacturacionBienesPagos) Then
3190                          ms_MensajeError = oFacturacionBienesPagos.MensajeError: GoTo ErrRecalculo
3200                       End If
3210                       lnDetallePago = lnDetallePago + 1
3220                    End If
3230                 End If
3240                 oRsTmp1.MoveNext
3250                 If oRsTmp1.EOF Then
3260                    Exit Do
3270                 End If
3280              Loop
3290            End If
3300            If lbGeneraReciboPago = False And lbCreaCabeceraPagosParaSegurosQueNoCubre = True And lnDetallePago = 0 Then
3310                  If Not oFactOrdenesBienes.Eliminar(oDoFactOrdenesBienes) Then
3320                     oDoFacturacionBienesPagos.IdOrden = oDoFactOrdenesBienes.IdOrden
3330                     oDoFacturacionBienesPagos.IdUsuarioAuditoria = oDoFactOrdenesBienes.IdUsuarioAuditoria
3340                     If Not oFacturacionBienesPagos.Eliminar(oDoFacturacionBienesPagos) Then
3350                        ms_MensajeError = oFacturacionBienesPagos.MensajeError: GoTo ErrRecalculo
3360                     End If
3370                     If Not oFactOrdenesBienes.Eliminar(oDoFactOrdenesBienes) Then
3380                         ms_MensajeError = oFactOrdenesBienes.MensajeError: GoTo ErrRecalculo
3390                     End If
3400                  End If
3410            End If
3420         Loop
3430      End If
          '
          
3440      oDoAtencion.IdAtencion = ml_IdAtencion
3450      If Not oAtenciones.SeleccionarPorId(oDoAtencion) Then
3460          ms_MensajeError = oAtenciones.MensajeError: GoTo ErrRecalculo
3470      End If
3480      oDoAtencion.IdFormaPago = lnIdTipoFinanciamientoNuevo
3490      oDoAtencion.idFuenteFinanciamiento = lnIdPlanNuevo
3500      oDoAtencion.IdUsuarioAuditoria = ml_IdUsuario
3510      If Not oAtenciones.Modificar(oDoAtencion) Then
3520          ms_MensajeError = oAtenciones.MensajeError: GoTo ErrRecalculo
3530      End If
          '
          If ldFechaAfiliacionSIS <> 0 Then
                oDoAtencionDatosAdicionales.IdAtencion = ml_IdAtencion
                If oAtencionesDatosAdicionales.SeleccionarPorId(oDoAtencionDatosAdicionales) Then
                    oDoAtencionDatosAdicionales.sisAfiliacion = lcSISafiliacion
                    oDoAtencionDatosAdicionales.SisCodigo = lcSIScodigo
                    oDoAtencionDatosAdicionales.idSiasis = lnIdSiaSis
                    oDoAtencionDatosAdicionales.IdUsuarioAuditoria = ml_IdUsuario
                    If oAtencionesDatosAdicionales.Modificar(oDoAtencionDatosAdicionales) Then
                    End If
                End If
          End If
          '
3540      Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_IdUsuario, "A", lnIdCuentaAtencion, "Recalculó Cta", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPlanActual)        'ListBarItems.idListItem
          
3550      oConexion.CommitTrans
3560      oConexion.Close
3570      Set oConexion = Nothing
          Set oFacturacionServicioFinan = Nothing
          Set oDoFacturacionServicioFinanc = Nothing
          Set oDoFactOrdenServicioPagos = Nothing
          Set oFactServicioPagos = Nothing
          Set oDoFacturacionServicioPagos = Nothing
          Set oFacturacionServicioPagos = Nothing
          '
          Set oFactOrdenesBienes = Nothing
          Set oDoFactOrdenesBienes = Nothing
          Set oFacturacionBienesFinan = Nothing
          Set oDoFacturacionBienesFinan = Nothing
          Set oFacturacionBienesPagos = Nothing
          Set oDoFacturacionBienesPagos = Nothing
          '
          Set oAtenciones = Nothing
          Set oDoAtencion = Nothing
          Set oAtencionesDatosAdicionales = Nothing
          Set oDoAtencionDatosAdicionales = Nothing
          '
          Set mo_ReglasFacturacion = Nothing
          Set mo_ReglasFarmacia = Nothing
          Set oRsTmp1 = Nothing
          Set oRsTmp2 = Nothing
          Set oRsTmp3 = Nothing
          Set oRsOrdenesPago = Nothing
          Set oRsDevoluciones = Nothing

3580      Exit Sub
ErrRecalculo:
3590      oConexion.RollbackTrans
3600      If Err.Description <> "" Then
3610          ms_MensajeError = ms_MensajeError + Chr(13) + Err.Description & _
                                sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                "Sub GrabaCantidadesPreciosEnElNuevoPlan", "dllFactUcEstadoCuenta.cls")  'debb-02/05/2016
                      
3620      End If
3630      MsgBox ms_MensajeError
3640      oConexion.Close
3650      Set oConexion = Nothing
          Set oFacturacionServicioFinan = Nothing
          Set oDoFacturacionServicioFinanc = Nothing
          Set oDoFactOrdenServicioPagos = Nothing
          Set oFactServicioPagos = Nothing
          Set oDoFacturacionServicioPagos = Nothing
          Set oFacturacionServicioPagos = Nothing
          '
          Set oFactOrdenesBienes = Nothing
          Set oDoFactOrdenesBienes = Nothing
          Set oFacturacionBienesFinan = Nothing
          Set oDoFacturacionBienesFinan = Nothing
          Set oFacturacionBienesPagos = Nothing
          Set oDoFacturacionBienesPagos = Nothing
          '
          Set oAtenciones = Nothing
          Set oDoAtencion = Nothing
          Set oAtencionesDatosAdicionales = Nothing
          Set oDoAtencionDatosAdicionales = Nothing
          '
          Set mo_ReglasFacturacion = Nothing
          Set mo_ReglasFarmacia = Nothing
          Set oRsTmp1 = Nothing
          Set oRsTmp2 = Nothing
          Set oRsTmp3 = Nothing
          Set oRsOrdenesPago = Nothing
          Set oRsDevoluciones = Nothing
          Exit Sub
3660      Resume
End Sub



