VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepMovimientoES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para el Reporte Movimiento de Entrada y Salida
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_IdAlmacen As Long
Dim ml_IdConcepto As Long
Dim ml_MovTipo As String    'E, S
Dim ml_IdEstado As Long   '0,1,   2=ambos
Dim ml_IdAlmacenOrigen As Long
Dim ml_IdAlmacenDestino As Long
Dim mo_ReporteUtil As New SIGHEntidades.ReporteUtil
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim ml_OrdenFiltro As String
Dim mda_FechaInicio As Date
Dim mda_FechaFin As Date
Dim ml_TextoDelFiltro  As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_idUsuario  As Long
Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property

Property Let Concepto(lValue As Long)
    ml_IdConcepto = lValue
End Property
Property Let MovTipo(lValue As String)
    ml_MovTipo = lValue
End Property

Property Let AlmacenOrigen(lValue As Long)
    ml_IdAlmacenOrigen = lValue
End Property
Property Let AlmacenDestino(lValue As Long)
    ml_IdAlmacenDestino = lValue
End Property
Property Let Estado(lValue As Long)
    ml_IdEstado = lValue
End Property
Property Let TextoDelFiltro(lValue As String)
    ml_TextoDelFiltro = lValue
End Property

Property Let FechaInicio(daValue As Date)
    mda_FechaInicio = daValue
End Property
Property Let FechaFin(daValue As Date)
    mda_FechaFin = daValue
End Property
Property Let IdAlmacen(lValue As Long)
    ml_IdAlmacen = lValue
End Property
Property Let OrdenFiltro(lValue As String)
    ml_OrdenFiltro = lValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New rMovimientoES
    oFormulario.idUsuario = ml_idUsuario
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub


Sub EjecutaFrm()
        Dim orProductosIngresados As New rProductosIngresados
        orProductosIngresados.NroReporte = 2
        orProductosIngresados.Show 1
        Set orProductosIngresados = Nothing
End Sub




Sub ReporteExoneraciones(lcFechaExonracionInicial As String, lcFechaExoneracionFinal As String, lnHwnd As Long)
    Dim mo_ReglasFarmacia As New ReglasFarmacia
    Dim mo_ReglasReportes As New ReglasReportes
    Dim oRsTmp1 As New Recordset
    Dim oRsReporte As New Recordset
    Dim ldFantencion As Date, lnHistoria As Long, lcPaciente As String, ldFexoneracion As Date
    Dim lcComprobante As String, lnMedicina As Double, lnMaterial As Double, lnCuenta As Long
    
    Set oRsTmp1 = mo_ReglasFarmacia.FacturacionBienesFinanciamientosXfechasExoneracion(CDate(lcFechaExonracionInicial), CDate(lcFechaExoneracionFinal))
    If oRsTmp1.RecordCount = 0 Then
        MsgBox "No existe información con esos datos", vbInformation, ""
    Else
        With oRsReporte
              .Fields.Append "Fmovimiento", adDate
              .Fields.Append "Historia", adVarChar, 10, adFldIsNullable
              .Fields.Append "Paciente", adVarChar, 50, adFldIsNullable
              .Fields.Append "Medicina", adDouble
              .Fields.Append "MaterialMedico", adDouble
              .Fields.Append "Total", adDouble
              .Fields.Append "Fexoneracion", adVarChar, 10, adFldIsNullable
              .Fields.Append "NroComprobante", adVarChar, 20, adFldIsNullable
              .LockType = adLockOptimistic
              .Open
        End With
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
           ldFantencion = oRsTmp1!fechaCreacion
           lnHistoria = oRsTmp1!NroHistoriaClinica
           lcPaciente = oRsTmp1!paciente
           ldFexoneracion = oRsTmp1!fechaAutoriza
           lcComprobante = oRsTmp1!comprobante
           lnMedicina = 0: lnMaterial = 0
           Do While Not oRsTmp1.EOF And ldFantencion = oRsTmp1!fechaCreacion And lcPaciente = oRsTmp1!paciente
              If oRsTmp1!TipoProducto = 0 Then
                 lnMedicina = lnMedicina + oRsTmp1!TotalFinanciado
              Else
                 lnMaterial = lnMaterial + oRsTmp1!TotalFinanciado
              End If
              oRsTmp1.MoveNext
              If oRsTmp1.EOF Then
                 Exit Do
              End If
           Loop
           oRsReporte.AddNew
           oRsReporte!fmovimiento = ldFantencion
           oRsReporte!historia = Trim(Str(lnHistoria))
           oRsReporte!paciente = Left(lcPaciente, 50)
           oRsReporte!medicina = lnMedicina
           oRsReporte!materialMedico = lnMaterial
           oRsReporte!total = lnMedicina + lnMaterial
           oRsReporte!fexoneracion = Format(ldFexoneracion, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
           oRsReporte!nroComprobante = lcComprobante
           oRsReporte.Update
        Loop
        oRsReporte.Sort = "fmovimiento,paciente"
        lcPaciente = "F. Exoneraciones " & lcFechaExonracionInicial & " al " & lcFechaExoneracionFinal
        mo_ReglasReportes.ExportarRecordSetAexcel oRsReporte, "EXONERACIONES", lcPaciente, "", _
                                                   lnHwnd, True, True
    End If
    Set mo_ReglasFarmacia = Nothing
    Set mo_ReglasReportes = Nothing
    Set oRsTmp1 = Nothing
    Set oRsReporte = Nothing

End Sub




Sub ReporteCreditosPendientes(lcFechaMovimInicial As String, lcFechaMovimFinal As String, lnHwnd As Long)
    Dim mo_ReglasFarmacia As New ReglasFarmacia
    Dim mo_ReglasReportes As New ReglasReportes
    Dim mo_ReglasFacturacion As New ReglasFacturacion
    Dim oRsTmp1 As New Recordset
    Dim oRsReporte As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oConexion As New Connection
    Dim ldFantencion As Date, lnHistoria As Long, lcPaciente As String, ldFexoneracion As Date
    Dim lcComprobante As String, lnMedicina As Double, lnMaterial As Double, lcFuente As String, lnCuenta As Long
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    '
    With oRsReporte
          .Fields.Append "Historia", adVarChar, 10, adFldIsNullable
          .Fields.Append "Paciente", adVarChar, 50, adFldIsNullable
          .Fields.Append "Medicina", adDouble
          .Fields.Append "MaterialMedico", adDouble
          .Fields.Append "Total", adDouble
          .Fields.Append "Fmovimiento", adDate
          .Fields.Append "Plan", adVarChar, 20, adFldIsNullable
          .Fields.Append "Cuenta", adInteger
          .LockType = adLockOptimistic
          .Open
    End With
    'procesa Pacientes con CREDITO solo con SEGUROS(menos SIS)
    Set oRsTmp1 = mo_ReglasFarmacia.farmMovimientoVentasPacientesSegurosConCredito(CDate(lcFechaMovimInicial), _
                                                                                    CDate(lcFechaMovimFinal))
    If oRsTmp1.RecordCount > 0 Then
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
            Set oRsTmp2 = mo_ReglasFacturacion.FacturacionReembolsosXcuenta(oRsTmp1!idCuentaAtencion, oConexion)
            If oRsTmp2.RecordCount > 0 Then
                oRsTmp1.MoveNext
            Else
                ldFantencion = oRsTmp1!fechaCreacion
                lnHistoria = oRsTmp1!NroHistoriaClinica
                lcPaciente = oRsTmp1!paciente
                lcFuente = oRsTmp1!dfuente
                lnCuenta = oRsTmp1!idCuentaAtencion
                lnMedicina = 0: lnMaterial = 0
                Do While Not oRsTmp1.EOF And ldFantencion = oRsTmp1!fechaCreacion And lcPaciente = oRsTmp1!paciente
                  If oRsTmp1!TipoProducto = 0 Then
                     lnMedicina = lnMedicina + oRsTmp1!TotalFinanciado
                  Else
                     lnMaterial = lnMaterial + oRsTmp1!TotalFinanciado
                  End If
                  oRsTmp1.MoveNext
                  If oRsTmp1.EOF Then
                     Exit Do
                  End If
                Loop
                oRsReporte.AddNew
                oRsReporte!historia = Trim(Str(lnHistoria))
                oRsReporte!paciente = Left(lcPaciente, 50)
                oRsReporte!medicina = lnMedicina
                oRsReporte!materialMedico = lnMaterial
                oRsReporte!total = lnMedicina + lnMaterial
                oRsReporte!fmovimiento = ldFantencion
                oRsReporte!plan = Left(lcFuente, 20)
                oRsReporte!cuenta = lnCuenta
                oRsReporte.Update
            End If
            oRsTmp2.Close
        Loop
    End If
    oRsTmp1.Close
    'procesa Pacientes con CREDITO solo PAGANTES(menos PREVENTAS)
    Set oRsTmp1 = mo_ReglasFarmacia.farmMovimientoVentasPacientesPagantesConCredito(CDate(lcFechaMovimInicial), _
                                                                                   CDate(lcFechaMovimFinal), _
                                                                                    0, 0)
    oRsTmp1.Filter = "FechaCobranza=NULL"
    If oRsTmp1.RecordCount > 0 Then
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
            ldFantencion = oRsTmp1!fechaCreacion
            lnHistoria = oRsTmp1!NroHistoriaClinica
            lcPaciente = oRsTmp1!paciente
            lcFuente = oRsTmp1!dfuente
            lnMedicina = 0: lnMaterial = 0
            lnCuenta = oRsTmp1!idCuentaAtencion
            Do While Not oRsTmp1.EOF And ldFantencion = oRsTmp1!fechaCreacion And lcPaciente = oRsTmp1!paciente
              If oRsTmp1!TipoProducto = 0 Then
                 lnMedicina = lnMedicina + oRsTmp1!TotalPagar
              Else
                 lnMaterial = lnMaterial + oRsTmp1!TotalPagar
              End If
              oRsTmp1.MoveNext
              If oRsTmp1.EOF Then
                 Exit Do
              End If
            Loop
            oRsReporte.AddNew
            oRsReporte!historia = Trim(Str(lnHistoria))
            oRsReporte!paciente = Left(lcPaciente, 50)
            oRsReporte!medicina = lnMedicina
            oRsReporte!materialMedico = lnMaterial
            oRsReporte!total = lnMedicina + lnMaterial
            oRsReporte!fmovimiento = ldFantencion
            oRsReporte!plan = Left(lcFuente, 20)
            oRsReporte!cuenta = lnCuenta
            oRsReporte.Update
        Loop
    End If
    oRsTmp1.Close
    '
    If oRsReporte.RecordCount = 0 Then
        MsgBox "No existe información con esos datos", vbInformation, ""
    Else
        oRsReporte.Sort = "fmovimiento,paciente"
        lcPaciente = "F. movimiento: " & lcFechaMovimInicial & " al " & lcFechaMovimFinal
        mo_ReglasReportes.ExportarRecordSetAexcel oRsReporte, "CREDITOS PENDIENTES DE PAGOS", lcPaciente, "", _
                                                   lnHwnd, True, True
    End If
    
    oConexion.Close
    Set mo_ReglasFarmacia = Nothing
    Set mo_ReglasReportes = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set oRsTmp1 = Nothing
    Set oRsReporte = Nothing
    Set oRsTmp2 = Nothing
    Set oConexion = Nothing

End Sub



Sub ReporteCreditosCancelados(lcFechaMovimInicial As String, lcFechaMovimFinal As String, lnHwnd As Long, _
                                                                     lbMovimientoEnElrangoFechas As Boolean)
    Dim mo_ReglasFarmacia As New ReglasFarmacia
    Dim mo_ReglasReportes As New ReglasReportes
    Dim mo_ReglasFacturacion As New ReglasFacturacion
    Dim oRsTmp1 As New Recordset
    Dim oRsReporte As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oConexion As New Connection
    Dim ldFantencion As Date, lnHistoria As Long, lcPaciente As String, ldFboleta As Date, lnCuenta As Long
    Dim lcComprobante As String, lnMedicina As Double, lnMaterial As Double, lcFuente As String, lbSigue As Boolean
    Dim lnExoneracion As Double
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    '
    With oRsReporte
          .Fields.Append "Historia", adVarChar, 10, adFldIsNullable
          .Fields.Append "Paciente", adVarChar, 40, adFldIsNullable
          .Fields.Append "Medicina", adDouble
          .Fields.Append "MaterialMedico", adDouble
          .Fields.Append "Total", adDouble
          .Fields.Append "Fmovimiento", adDate
          .Fields.Append "Fcancelacion", adVarChar, 10, adFldIsNullable
          .Fields.Append "Plan", adVarChar, 15, adFldIsNullable
          .Fields.Append "NroComprobante", adVarChar, 20, adFldIsNullable
          .Fields.Append "Cuenta", adInteger
          .LockType = adLockOptimistic
          .Open
    End With
    'procesa Pacientes con CREDITO solo con SEGUROS(menos SIS)
    Set oRsTmp1 = mo_ReglasFarmacia.farmMovimientoVentasPacientesSegurosConCredito(CDate(lcFechaMovimInicial), _
                                                                                    CDate(lcFechaMovimFinal))
    If oRsTmp1.RecordCount > 0 Then
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
            lbSigue = True
            Set oRsTmp2 = mo_ReglasFacturacion.FacturacionReembolsosXcuenta(oRsTmp1!idCuentaAtencion, oConexion)
            If oRsTmp2.RecordCount = 0 Then
               lbSigue = False
            Else
               If oRsTmp2!fechaCobranza >= CDate(lcFechaMovimInicial) And oRsTmp2!fechaCobranza <= CDate(lcFechaMovimFinal) Then
                  If lbMovimientoEnElrangoFechas = False Then
                     lbSigue = False
                  End If
               Else
                  If lbMovimientoEnElrangoFechas = True Then
                     lbSigue = False
                  End If
               End If
               ldFboleta = oRsTmp2!fechaCobranza
               lcComprobante = oRsTmp2!nroSerie & "-" & oRsTmp2!nroDocumento
            End If
            oRsTmp2.Close
            If lbSigue = False Then
                oRsTmp1.MoveNext
            Else
                ldFantencion = oRsTmp1!fechaCreacion
                lnHistoria = oRsTmp1!NroHistoriaClinica
                lcPaciente = oRsTmp1!paciente
                lcFuente = oRsTmp1!dfuente
                lnCuenta = oRsTmp1!idCuentaAtencion
                lnMedicina = 0: lnMaterial = 0
                Do While Not oRsTmp1.EOF And ldFantencion = oRsTmp1!fechaCreacion And lcPaciente = oRsTmp1!paciente
                  If oRsTmp1!TipoProducto = 0 Then
                     lnMedicina = lnMedicina + oRsTmp1!TotalFinanciado
                  Else
                     lnMaterial = lnMaterial + oRsTmp1!TotalFinanciado
                  End If
                  oRsTmp1.MoveNext
                  If oRsTmp1.EOF Then
                     Exit Do
                  End If
                Loop
                oRsReporte.AddNew
                oRsReporte!historia = Trim(Str(lnHistoria))
                oRsReporte!paciente = Left(lcPaciente, 50)
                oRsReporte!medicina = lnMedicina
                oRsReporte!materialMedico = lnMaterial
                oRsReporte!total = lnMedicina + lnMaterial
                oRsReporte!fmovimiento = ldFantencion
                oRsReporte!plan = Left(lcFuente, 15)
                oRsReporte!fcancelacion = Format(ldFboleta, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
                oRsReporte!nroComprobante = lcComprobante
                oRsReporte!cuenta = lnCuenta
                oRsReporte.Update
            End If
        Loop
    End If
    oRsTmp1.Close
    'procesa Pacientes con CREDITO solo PAGANTES(menos PREVENTAS)
    Set oRsTmp1 = mo_ReglasFarmacia.farmMovimientoVentasPacientesPagantesConCredito(CDate(lcFechaMovimInicial), _
                                                                                   CDate(lcFechaMovimFinal), _
                                                                                    0, 0)
    oRsTmp1.Filter = "FechaCobranza<>NULL"
    If oRsTmp1.RecordCount > 0 Then
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF

            lbSigue = True
            If oRsTmp1!fechaCobranza >= CDate(lcFechaMovimInicial) And oRsTmp1!fechaCobranza <= CDate(lcFechaMovimFinal) Then
               If lbMovimientoEnElrangoFechas = False Then
                  lbSigue = False
               End If
            Else
               If lbMovimientoEnElrangoFechas = True Then
                  lbSigue = False
               End If
            End If
            If lbSigue = False Then
                oRsTmp1.MoveNext
            Else
                ldFantencion = oRsTmp1!fechaCreacion
                lnHistoria = oRsTmp1!NroHistoriaClinica
                lcPaciente = oRsTmp1!paciente
                lcFuente = oRsTmp1!dfuente
                ldFboleta = oRsTmp1!fechaCobranza
                lcComprobante = oRsTmp1!comprobante
                lnCuenta = oRsTmp1!idCuentaAtencion
                lnMedicina = 0: lnMaterial = 0
                Do While Not oRsTmp1.EOF And ldFantencion = oRsTmp1!fechaCreacion And lcPaciente = oRsTmp1!paciente
                  lnExoneracion = 0
                  If oRsTmp1!exoneraciones > 0 Then
                     Set oRsTmp2 = mo_ReglasFarmacia.FacturacionBienesFinancSeleccionarPorProducto(oConexion, _
                                                        oRsTmp1!movNumero, oRsTmp1!MovTipo, oRsTmp1!idProducto)
                     oRsTmp2.Filter = "idTipoFinanciamiento=9"
                     If oRsTmp2.RecordCount > 0 Then
                        oRsTmp2.MoveFirst
                        Do While Not oRsTmp2.EOF
                           lnExoneracion = lnExoneracion + oRsTmp2!TotalFinanciado
                           oRsTmp2.MoveNext
                        Loop
                     End If
                     oRsTmp2.Close
                  End If
                  If oRsTmp1!TipoProducto = 0 Then
                     lnMedicina = lnMedicina + oRsTmp1!TotalPagar - lnExoneracion
                  Else
                     lnMaterial = lnMaterial + oRsTmp1!TotalPagar - lnExoneracion
                  End If
                  oRsTmp1.MoveNext
                  If oRsTmp1.EOF Then
                     Exit Do
                  End If
                Loop
                oRsReporte.AddNew
                oRsReporte!historia = Trim(Str(lnHistoria))
                oRsReporte!paciente = Left(lcPaciente, 50)
                oRsReporte!medicina = lnMedicina
                oRsReporte!materialMedico = lnMaterial
                oRsReporte!total = lnMedicina + lnMaterial
                oRsReporte!fmovimiento = ldFantencion
                oRsReporte!plan = Left(lcFuente, 15)
                oRsReporte!fcancelacion = Format(ldFboleta, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
                oRsReporte!nroComprobante = lcComprobante
                oRsReporte!cuenta = lnCuenta
                oRsReporte.Update
            End If
        Loop
    End If
    oRsTmp1.Close
    '
    If oRsReporte.RecordCount = 0 Then
        MsgBox "No existe información con esos datos", vbInformation, ""
    Else
        oRsReporte.Sort = "fmovimiento,paciente"
        lcPaciente = "F. movimiento: " & lcFechaMovimInicial & " al " & lcFechaMovimFinal
        mo_ReglasReportes.ExportarRecordSetAexcel oRsReporte, "CREDITOS OTORGADOS Y CANCELADOS " & _
                                                              IIf(lbMovimientoEnElrangoFechas = True, _
                                                              "EN EL RANGO DE FECHAS", "FUERA DEL RANGO DE FECHAS"), _
                                                              lcPaciente, "", lnHwnd, True, True
    End If
    
    oConexion.Close
    Set mo_ReglasFarmacia = Nothing
    Set mo_ReglasReportes = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set oRsTmp1 = Nothing
    Set oRsReporte = Nothing
    Set oRsTmp2 = Nothing
    Set oConexion = Nothing

End Sub





