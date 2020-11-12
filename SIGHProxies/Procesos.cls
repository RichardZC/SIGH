VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Procesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para varios procesos
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_lcNombrePc As String
Dim ml_idUsuario As Long
Dim mo_lnIdTablaLISTBARITEMS  As Long
Dim ms_MensajeError As String
Dim lcSql As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lcParaMsgbox As String
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim oRsDxCpt As New Recordset
Dim ml_ValorActual As String
'GLCC 02/11/20 CAMBIO36 INICIO
'ANTERIOR: Const wxNueve As String = "9"
'GLCC 02/11/20 CAMBIO36 FIN
Const wxObservacionCrearPaciente As String = "Creado desde Citas Web"
Property Get ValorActual() As String
   ValorActual = ml_ValorActual
End Property

Property Let ValorActual(lValue As String)
   ml_ValorActual = lValue
End Property
Property Let MensajeError(lValue As String)
    ms_MensajeError = lValue
End Property

Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property







Public Sub CrearTablaTemporalDxCPT() 'Frank 12092014
    If oRsDxCpt.State = 1 Then
       Set oRsDxCpt = Nothing
    End If
    With oRsDxCpt
        .Fields.Append "IdDxCpt", adInteger
        .Fields.Append "Dx", adVarChar, 255, adFldIsNullable
        .Fields.Append "TDx", adVarChar, 255, adFldIsNullable
        .Fields.Append "CDx", adVarChar, 255, adFldIsNullable
        .Fields.Append "LDx", adVarChar, 255, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Public Sub LimpiarTablaDxCpt() 'Frank 12092014
    With oRsDxCpt
        If .RecordCount > 0 Then
           .MoveFirst
           Do While Not .EOF
              .Delete
              .Update
              .MoveNext
           Loop
        End If
    End With
End Sub




Function DevuelveDESDEdeTablaEdad000(lnTipoEdad As Long, lcRangoTabla As String) As Integer
    Dim lcBuscarTxt As String, lnBuscarLen As Integer, lnPosic As Integer
    Dim lcRango As String, lnFor As Integer
    lcBuscarTxt = "(" & Trim(str(lnTipoEdad)) & "."
    lnBuscarLen = Len(lcBuscarTxt)
    lnPosic = InStr(lcRangoTabla, lcBuscarTxt)
    lcRango = ""
    For lnFor = lnPosic + lnBuscarLen To Len(lcRangoTabla)
        If Mid(lcRangoTabla, lnFor, 1) = "-" Then
           Exit For
        End If
        lcRango = lcRango + Mid(lcRangoTabla, lnFor, 1)
    Next
    DevuelveDESDEdeTablaEdad000 = Val(lcRango)
End Function

Function DevuelveHASTAdeTablaEdad000(lnTipoEdad As Long, lcRangoTabla As String) As Integer
    Dim lcBuscarTxt As String, lnBuscarLen As Integer, lnPosic As Integer
    Dim lcRango As String, lnFor As Integer
    lcBuscarTxt = "-" & Trim(str(lnTipoEdad)) & "."
    lnBuscarLen = Len(lcBuscarTxt)
    lnPosic = InStr(lcRangoTabla, lcBuscarTxt)
    lcRango = ""
    For lnFor = lnPosic + lnBuscarLen To Len(lcRangoTabla)
        If Mid(lcRangoTabla, lnFor, 1) = ")" Then
           Exit For
        End If
        lcRango = lcRango + Mid(lcRangoTabla, lnFor, 1)
    Next
    DevuelveHASTAdeTablaEdad000 = Val(lcRango)
End Function






'********************************************************* SEM *****************************
Sub ExportaSEMversion2009()
'  If (txtFechaInicio.Text = "" Or Not IsDate(txtFechaInicio.Text)) Or (txtFechaFin.Text = "" Or Not IsDate(txtFechaFin.Text)) Or ssoEmergencias.Value = False And ssoHospitalizacion.Value = False Or txtNCorrelativo.Text = "" Then
'    MsgBox "Debe completar los datos requeridos", vbInformation, Me.Caption
'    txtFechaInicio.SetFocus
'    Exit Sub
'  End If
'
'  If MsgBox("Esta seguro de exportar datos?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
'    Me.MousePointer = 11
'    Frame4.Enabled = False
'    btnAceptar.Enabled = False
'
'    'On Error GoTo ErrorProceso
'    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
'    Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
'    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
'    Dim lcBuscaParametro As New SIGHDatos.Parametros
'    Dim rsReporte As New Recordset
'    Dim oRsTmp As New Recordset
'    Dim oRsFox As New Recordset
'    Dim oConexionFox As New Connection
'    Dim lcDx1 As String, lcDx2 As String
'    Dim lnNumMov As Long, lcSql As String, lnEstancia As Integer
'    Dim lcFec1 As String, lcFec2 As String
'    Dim oConexion As New Connection
'        '
'    oConexion.CommandTimeout = 300
'    oConexion.Open SIGHEntidades.CadenaConexion
'        '
'    oConexionFox.CommandTimeout = 300
'    oConexionFox.Open "DSN=his"
'
'    lcSql = "SELECT *  FROM " & ml_Tabla & " WHERE month(fecIng)>=" & Month(CDate(txtFechaInicio.Text)) & " and month(fecIng)<=" & Month(CDate(txtFechaFin.Text))
'    oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
'    If oRsFox.RecordCount > 0 Then
'      oRsFox.MoveFirst
'      Do While Not oRsFox.EOF
'        If oRsFox.Fields!fecIng >= CDate(txtFechaInicio.Text) And oRsFox.Fields!fecIng <= CDate(txtFechaFin.Text) Then
'          oRsFox.Delete
'          oRsFox.Update
'        End If
'        oRsFox.MoveNext
'      Loop
'    End If
'    oRsFox.Close
'
'    oRsFox.Open "SELECT * FROM " & ml_Tabla & " ORDER BY numMov", oConexionFox, adOpenKeyset, adLockOptimistic
'    If oRsFox.RecordCount = 0 Then
'      'txtNCorrelativo.Text = "1"
'    Else
'      oRsFox.MoveLast
'      txtNCorrelativo.Text = oRsFox.Fields!numMov + 1
'    End If
'    DoEvents
'
'    Set rsReporte = mo_ReglasAdmision.AtencionesSeleccionarHospEmergPorFechasDeEgresoMedico(CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text), "FECHAEGRESO")
'    rsReporte.Filter = "codigoCondicionAltaSEM<>null and idTipoServicio=" & ml_Valor
'    If rsReporte.RecordCount = 0 Then
'      MsgBox "No existe Información  con esos Datos", vbInformation, Me.Caption
'    Else
'      lnNumMov = Val(txtNCorrelativo.Text)
'      rsReporte.MoveFirst
'      Do While Not rsReporte.EOF
'        'solo Hospitalizados
'        lcDx1 = "": lcDx2 = ""
'        Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(rsReporte.Fields!idAtencion, 3, oConexion)
'        oRsTmp.Filter = "IdTipoDiagnostico=301"   'vivos-princial
'        If oRsTmp.RecordCount > 0 Then
'          '*****vivos
'          oRsTmp.MoveFirst
'          lcDx1 = Left(oRsTmp.Fields!CodigoCIE2004, 5)
'          oRsTmp.Filter = "IdTipoDiagnostico=302"
'          If oRsTmp.RecordCount > 0 Then
'            oRsTmp.MoveFirst
'            lcDx2 = Left(oRsTmp.Fields!CodigoCIE2004, 5)
'          End If
'        Else
'          '******fallecidos
'          Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(rsReporte.Fields!idAtencion, 4, oConexion)
'          If oRsTmp.RecordCount > 0 Then
'                oRsTmp.Filter = "IdTipoDiagnostico=303"    'Final
'                If oRsTmp.RecordCount > 0 Then
'                  oRsTmp.MoveFirst
'                  lcDx1 = Left(oRsTmp.Fields!CodigoCIE2004, 5)
'                End If
'                oRsTmp.Filter = "IdTipoDiagnostico=304 or IdTipoDiagnostico=305"    'fallecidos - Final
'                If oRsTmp.RecordCount > 0 Then
'                  oRsTmp.MoveFirst
'                  lcDx2 = Left(oRsTmp.Fields!CodigoCIE2004, 5)
'                End If
'          End If
'        End If
'
'        '
'        lnEstancia = DiasDeEstanciaEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, rsReporte!FechaEgreso, rsReporte!horaEgreso)
'        lnEstancia = lcBuscaParametro.DiasDelPacienteEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, rsReporte!FechaEgreso, rsReporte!horaEgreso)
'        If lnEstancia = 0 Then lnEstancia = 1
'        '
'        oRsFox.AddNew
'        oRsFox.Fields!numMov = lnNumMov
'        oRsFox.Fields!cod_disa = lcBuscaParametro.SeleccionaFilaParametro(239)
'        oRsFox.Fields!cod_red = lcBuscaParametro.SeleccionaFilaParametro(240)
'        oRsFox.Fields!cod_mred = lcBuscaParametro.SeleccionaFilaParametro(241)
'        oRsFox.Fields!cod_est = Trim(Str(Val(lcBuscaParametro.SeleccionaFilaParametro(208))))
'        If Not IsNull(rsReporte.Fields!NroHistoriaClinica) Then
'          oRsFox.Fields!numhc = Left(Trim(Str(rsReporte.Fields!NroHistoriaClinica)), 7)
'        Else
'          oRsFox.Fields!numhc = ""
'        End If
'        oRsFox.Fields!apenom = Left(Trim(rsReporte.Fields!ApellidoPaterno) & " " & Trim(rsReporte.Fields!ApellidoMaterno) & " " & Trim(rsReporte.Fields!PrimerNombre) & " " & IIf(IsNull(rsReporte.Fields!SegundoNombre), "", Trim(rsReporte.Fields!SegundoNombre)), 50)
'        oRsFox.Fields!Sexo = Trim(Str(rsReporte.Fields!idTipoSexo))
'        oRsFox.Fields!Edad = rsReporte.Fields!Edad
'        oRsFox.Fields!TipoEdad = rsReporte.Fields!tipoEdadSEM
'        oRsFox.Fields!ubigeo = lcBuscaParametro.SeleccionaFilaParametro(242)
'        oRsFox.Fields!fecIng = rsReporte.Fields!FechaIngreso  'Format(rsReporte.Fields!FechaIngreso, "mm/dd/yyyy")
'        oRsFox.Fields!fecegr = rsReporte.Fields!FechaEgreso    'Format(rsReporte.Fields!FechaEgreso, "mm/dd/yyyy")
'        oRsFox.Fields!totalest = lnEstancia
'        oRsFox.Fields!Servicio = IIf(IsNull(rsReporte.Fields!codigoServicioSEM), "xx", rsReporte.Fields!codigoServicioSEM)
'        oRsFox.Fields!condicion = rsReporte.Fields!codigoCondicionAltaSEM
'        oRsFox.Fields!seguro = rsReporte.Fields!CodigoFuenteFinanciamientoSEM
'        oRsFox.Fields!coddiag = lcDx1
'        oRsFox.Fields!coddiag2 = lcDx2
'        oRsFox.Fields!ubicacionh = IIf(IsNull(rsReporte.Fields!ubicacionSEM), "xxxxx", rsReporte.Fields!ubicacionSEM)
'        oRsFox.Fields!medicoh = IIf(IsNull(rsReporte.Fields!Colegiatura), "xxxxx", Left(Trim(rsReporte.Fields!Colegiatura), 5))
'        oRsFox.Fields!estado = 1
'        oRsFox.Fields!fechareg = Format(Date, "mm/dd/yyyy")
'        oRsFox.Update
'        lnNumMov = lnNumMov + 1
'        rsReporte.MoveNext
'      Loop
'      Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "A", 0, "Atenciones", oConexion, 500 + 140, mo_lcNombrePc, "Exportar al SEM: " & txtFechaInicio.Text & " al " & txtFechaFin.Text)       '500+ ListBarReporte.idReporte
'    End If
'    oRsFox.Close
'    oConexionFox.Close
'    oConexion.Close
'    Set oConexion = Nothing
'    Set oConexionFox = Nothing
'    Set oRsFox = Nothing
'    MsgBox "Exportación de datos finalizado correctamente" & Chr(13) & "Valor del Último NumMov: " & lnNumMov - 1, vbInformation, Me.Caption
'    Frame4.Enabled = True
'    btnAceptar.Enabled = True
'    txtFechaInicio.SetFocus
'    Me.MousePointer = 1
'    'Me.Visible = False
'    Exit Sub
'  End If
'  Exit Sub
'ErrorProceso:
'  MsgBox Err.Description, vbInformation, Me.Caption
'  'Resume
End Sub
Sub ExportaSEM(lcFechaInicio As String, lcFechaFinal As String, lcCorrelativo As String, ml_Valor As Integer, ml_Tabla As String)
    On Error GoTo ErrorProceso
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_AdminAdmision As New ReglasAdmision
    Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_ReglasServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
    Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
    
    Dim rsReporte As New Recordset
    Dim oRsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oRsFox As New Recordset
    Dim oConexionFox As New Connection
    Dim lcDx1 As String, lcDx2 As String
    Dim lnNumMov As Long, lcSql As String, lnEstancia As Integer
    Dim lcFec1 As String, lcFec2 As String
    Dim lcMotivo As String, lcDestinoAtencion As String, lcReferido As String, lcUsuario As String
    Dim lcDireccionDomicilio As String, lcNombreAcompaniante As String
    Dim lnEsObservacionEmergencia As Integer, lcServicio As String, lcMedico As String, lcServicioObs As String
    Dim lcMedicoObs As String, ldFechaIngObs As Date, lcHoraIngObs As String, lcCamaObs As String, ldFechaSalObs As Date
    Dim lcHoraSalObs As String, lcDxObs As String, lcHoraEstanciaMax As String, lnTotal As Long
    Dim oConexion As New Connection
    lcHoraEstanciaMax = lcBuscaParametro.SeleccionaFilaParametro(201)
        '
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
        '
    oConexionFox.CommandTimeout = 300
    lcParaMsgbox = " (Eliminar Datos)"
    If ml_Valor = 2 Then
        '*********************************Emergencia
        
        oConexionFox.Open "DSN=SEMNUEVO"
        ml_Tabla = "emergencias"
        mo_ReglasComunes.AbreTablaFox oRsFox, oConexionFox, ml_Tabla, "", ""
        If oRsFox.RecordCount > 0 Then
          oRsFox.MoveFirst
          Do While Not oRsFox.EOF
            If oRsFox.Fields!em_fecha >= CDate(lcFechaInicio) And oRsFox.Fields!em_fecha <= CDate(lcFechaFinal) Then
              oRsFox.Delete
              oRsFox.Update
            End If
            oRsFox.MoveNext
          Loop
        End If
        oRsFox.Close
        '
        mo_ReglasComunes.AbreTablaFox oRsFox, oConexionFox, ml_Tabla, "", "em_codigo"
        If oRsFox.RecordCount = 0 Then
          'lcCorrelativo = "1"
        Else
          oRsFox.MoveLast
          lcCorrelativo = oRsFox.Fields!em_codigo + 1
        End If
    Else
        '*********************************Hospitalizacion
        oConexionFox.Open "DSN=SEM"
        mo_ReglasComunes.AbreTablaFox oRsFox, oConexionFox, ml_Tabla, "month(fecIng)>=" & Month(CDate(lcFechaInicio)) & " and month(fecIng)<=" & Month(CDate(lcFechaFinal)), ""
        If oRsFox.RecordCount > 0 Then
          oRsFox.MoveFirst
          Do While Not oRsFox.EOF
            If oRsFox.Fields!fecIng >= CDate(lcFechaInicio) And oRsFox.Fields!fecIng <= CDate(lcFechaFinal) Then
              oRsFox.Delete
              oRsFox.Update
            End If
            oRsFox.MoveNext
          Loop
        End If
        oRsFox.Close
        '
        mo_ReglasComunes.AbreTablaFox oRsFox, oConexionFox, ml_Tabla, "", "numMov"
        If oRsFox.RecordCount = 0 Then
          'lcCorrelativo = "1"
        Else
          oRsFox.MoveLast
          lcCorrelativo = oRsFox.Fields!numMov + 1
        End If
    End If
 '   DoEvents
    '
    lcUsuario = " "
    Set oRsTmp1 = mo_ReglasComunes.EmpleadosSeleccionarPorIdEmpleado(ml_idUsuario)
    If oRsTmp1.RecordCount > 0 Then
       If Not IsNull(oRsTmp1.Fields!Usuario) Then
          lcUsuario = oRsTmp1.Fields!Usuario
       End If
    End If
    oRsTmp1.Close
    '
    Set rsReporte = mo_ReglasAdmision.AtencionesSeleccionarHospEmergPorFechasDeEgresoMedico(CDate(lcFechaInicio), CDate(lcFechaFinal), "FECHAEGRESO")
    rsReporte.Filter = "codigoCondicionAltaSEM<>null and idTipoServicio=" & ml_Valor
    lnTotal = rsReporte.RecordCount
    If lnTotal = 0 Then
      MsgBox "No existe Información  con esos Datos", vbInformation, "Mensaje"
    Else
      
      
      lnNumMov = Val(lcCorrelativo)
      rsReporte.MoveFirst
      Do While Not rsReporte.EOF
If rsReporte.Fields!idCuentaAtencion = 50829 Then
lcDx1 = ""
End If
        lcParaMsgbox = " -->Cuenta: " & rsReporte.Fields!idCuentaAtencion
        'solo Hospitalizados
        lcDx1 = "": lcDx2 = ""
        Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(rsReporte.Fields!idAtencion, 3, oConexion)
        oRsTmp.Filter = "IdTipoDiagnostico=301"   'vivos-princial
        If oRsTmp.RecordCount > 0 Then
          '*****vivos
          oRsTmp.MoveFirst
          lcDx1 = Left(oRsTmp.Fields!CodigoCie2004, 5)
          oRsTmp.Filter = "IdTipoDiagnostico=302"
          If oRsTmp.RecordCount > 0 Then
            oRsTmp.MoveFirst
            lcDx2 = Left(oRsTmp.Fields!CodigoCie2004, 5)
          End If
        Else
          '******fallecidos
          Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(rsReporte.Fields!idAtencion, 4, oConexion)
          If oRsTmp.RecordCount > 0 Then
                oRsTmp.Filter = "IdTipoDiagnostico=303"    'Final
                If oRsTmp.RecordCount > 0 Then
                  oRsTmp.MoveFirst
                  lcDx1 = Left(oRsTmp.Fields!CodigoCie2004, 5)
                End If
                oRsTmp.Filter = "IdTipoDiagnostico=304 or IdTipoDiagnostico=305"    'fallecidos - Final
                If oRsTmp.RecordCount > 0 Then
                  oRsTmp.MoveFirst
                  lcDx2 = Left(oRsTmp.Fields!CodigoCie2004, 5)
                End If
          End If
        End If
    
        'oRsTmp.Close
        '
        lcMotivo = " "
        Set oRsTmp1 = mo_AdminAdmision.AtencionesEmergenciaXidAtencion(rsReporte.Fields!idAtencion, oConexion)
        If oRsTmp1.RecordCount > 0 Then
           If Not IsNull(oRsTmp1.Fields!motivoSEM) Then
              lcMotivo = oRsTmp1.Fields!motivoSEM
           End If
        End If
        oRsTmp1.Close
        '
        lcDireccionDomicilio = "": lcNombreAcompaniante = ""
        Set oRsTmp1 = mo_AdminAdmision.atencionesDatosAdicionalesXfiltro("dbo.AtencionesDatosAdicionales.idAtencion=" & rsReporte.Fields!idAtencion, oConexion)
        If oRsTmp1.RecordCount > 0 Then
           If Not IsNull(oRsTmp1.Fields!DireccionDomicilio) Then
              lcDireccionDomicilio = oRsTmp1.Fields!DireccionDomicilio
           End If
           If Not IsNull(oRsTmp1.Fields!NombreAcompaniante) Then
              lcNombreAcompaniante = oRsTmp1.Fields!NombreAcompaniante
           End If
        End If
        oRsTmp1.Close
        '
        lcDestinoAtencion = " "
        If Not IsNull(rsReporte.Fields!IdDestinoAtencion) Then
            Set oRsTmp1 = mo_ReglasAdmision.TiposDestinoAtencionXidDestinoAtencion(rsReporte.Fields!IdDestinoAtencion, oConexion)
            If oRsTmp1.RecordCount > 0 Then
               If IsNull(oRsTmp1.Fields!destinoSEM) Then
                    MsgBox "Existen Servicios de Emergencia sin CODIGO SEM", vbInformation, "Mensaje"  'debb 05/12/2012
                    Exit Sub
               Else
                    lcDestinoAtencion = oRsTmp1.Fields!destinoSEM
               End If
            End If
            oRsTmp1.Close
        End If
        '
        lcReferido = " "
        If Not IsNull(rsReporte.Fields!idEstablecimientoDestino) Then
            Set oRsTmp1 = mo_ReglasComunes.EstablecimientosSeleccionarPorFiltro("idEstablecimiento=" & rsReporte.Fields!idEstablecimientoDestino, oConexion)
            If oRsTmp1.RecordCount > 0 Then
               lcReferido = Left(Trim(str(Val(oRsTmp1.Fields!codigo))), 4)
            End If
            oRsTmp1.Close
        End If
        '
        lnEstancia = lcBuscaParametro.DiasDelPacienteEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, rsReporte!FechaEgreso, rsReporte!horaEgreso, lcHoraEstanciaMax)
        If lnEstancia = 0 Then lnEstancia = 1
        '
        If ml_Valor = 2 Then
            '*********************************Emergencia
            lcServicio = IIf(IsNull(rsReporte.Fields!codigoServicioSem), ".", rsReporte.Fields!codigoServicioSem)
            lcMedico = IIf(IsNull(rsReporte.Fields!Colegiatura), ".", Left(rsReporte.Fields!Colegiatura, 5))
            lnEsObservacionEmergencia = 0
            If Not IsNull(rsReporte.Fields!EsObservacionEmergencia) Then
               If rsReporte.Fields!EsObservacionEmergencia = True Then
                    lnEsObservacionEmergencia = 1
                    Set oRsTmp1 = mo_ReglasAdmision.AtencionesEstanciaHospitalariaXidAtencionIdServicioIngreso(rsReporte.Fields!IdServicioIngreso, rsReporte.Fields!idAtencion, oConexion)
                    If oRsTmp1.RecordCount > 0 Then
                       If oRsTmp1.Fields!EsObservacionEmergencia = True Then
                          '***El Servicio de Ingreso y el Servicio de Alta son OBSERVACIONES DE EMERGENCIA
                          '***Solo se muestra datos del Servicio de Alta ya no es necesario mostrar los datos
                          '***de OBSERVACIONES
                          lnEsObservacionEmergencia = 0
                       Else
                          '***El servicio de Ingreso es un CONSULTORIO y el servicio de alta es OBSERVACION
                          '***Hubo transferencia desde CONSULTORIO hacia OBSERVACION
                          '***Muestra datos del CONSULTORIO y datos DE OBSERVACION
                          lcServicioObs = lcServicio
                          lcMedicoObs = lcMedico
                          lcDxObs = lcDx1
                          ldFechaIngObs = oRsTmp1.Fields!fechaDesocupacion
                          lcHoraIngObs = oRsTmp1.Fields!horaDesocupacion
                          'debb 05/12/2012
                          If IsNull(oRsTmp1.Fields!cama) Then  'debb 05/12/2012
                             MsgBox "Para la cuenta. " & rsReporte.Fields!idCuentaAtencion & " no se registró cama, siendo un SERVICIO DE OBSERVACION", vbInformation, "Mensaje"
                             Exit Sub
                          Else
                             lcCamaObs = oRsTmp1.Fields!cama
                          End If
                          'debb 05/12/2012
                          ldFechaSalObs = rsReporte.Fields!FechaEgreso
                          lcHoraSalObs = rsReporte.Fields!horaEgreso
                          lcServicio = IIf(IsNull(oRsTmp1.Fields!codigoServicioSem), ".", oRsTmp1.Fields!codigoServicioSem)
                          lcMedico = IIf(IsNull(oRsTmp1.Fields!Colegiatura), ".", Left(oRsTmp1.Fields!Colegiatura, 5))
                       End If
                    End If
                    oRsTmp1.Close
               End If
            End If
            oRsFox.AddNew
            oRsFox.Fields!em_codigo = lnNumMov
            oRsFox.Fields!em_codest = Trim(str(Val(lcBuscaParametro.SeleccionaFilaParametro(208))))
            oRsFox.Fields!em_disa = lcBuscaParametro.SeleccionaFilaParametro(239)
            oRsFox.Fields!em_red = lcBuscaParametro.SeleccionaFilaParametro(240)
            oRsFox.Fields!em_mred = lcBuscaParametro.SeleccionaFilaParametro(241)
            If Not IsNull(rsReporte.Fields!NroHistoriaClinica) Then
               oRsFox.Fields!em_registro = Trim(str(rsReporte.Fields!NroHistoriaClinica))
            Else
               oRsFox.Fields!em_registro = "."
            End If
            oRsFox.Fields!em_fecha = rsReporte.Fields!FechaIngreso  'Format(rsReporte.Fields!FechaIngreso, "mm/dd/yyyy")
            If IsNull(rsReporte.Fields!idTipoGravedad) Then
               oRsFox.Fields!em_conding = 1
            Else
               oRsFox.Fields!em_conding = IIf(rsReporte.Fields!idTipoGravedad = 5, 2, 1)
            End If
            oRsFox.Fields!em_apenom = Left(Trim(rsReporte.Fields!ApellidoPaterno) & " " & Trim(rsReporte.Fields!ApellidoMaterno) & " " & Trim(rsReporte.Fields!PrimerNombre) & " " & IIf(IsNull(rsReporte.Fields!SegundoNombre), "", Trim(rsReporte.Fields!SegundoNombre)), 50)
            oRsFox.Fields!em_sexo = Trim(str(rsReporte.Fields!idTipoSexo))
            oRsFox.Fields!em_edad = rsReporte.Fields!Edad
            oRsFox.Fields!em_tedad = rsReporte.Fields!tipoEdadSEM
            oRsFox.Fields!em_tipopac = rsReporte.Fields!CodigoFuenteFinanciamientoSEM
            oRsFox.Fields!em_direccion = Left(lcDireccionDomicilio, 50)
            If Not IsNull(rsReporte.Fields!IdDistritoProcedencia) Then
               oRsFox.Fields!em_procedencia = Right("0" & Trim(str(rsReporte.Fields!IdDistritoProcedencia)), 6)
            Else
               oRsFox.Fields!em_procedencia = " "
            End If
            If Not IsNull(rsReporte.Fields!IdDistritoDomicilio) Then
               oRsFox.Fields!em_residencia = Right("0" & Trim(str(rsReporte.Fields!IdDistritoDomicilio)), 6)
            Else
               oRsFox.Fields!em_residencia = " "
            End If
            oRsFox.Fields!em_acompa = Left(lcNombreAcompaniante, 40)
            oRsFox.Fields!em_horate = rsReporte.Fields!HoraIngreso
            oRsFox.Fields!em_diagp = lcDx1
            oRsFox.Fields!em_diags = lcDx2
            oRsFox.Fields!em_medico = lcMedico
            oRsFox.Fields!em_motivo = lcMotivo
            oRsFox.Fields!em_servicio = lcServicio
            oRsFox.Fields!em_observ = lnEsObservacionEmergencia
            If lnEsObservacionEmergencia = 1 Then
                oRsFox.Fields!em_obshin = lcHoraIngObs
                oRsFox.Fields!em_obshsa = lcHoraSalObs
                oRsFox.Fields!em_ofechin = ldFechaIngObs
                oRsFox.Fields!em_ofechsa = ldFechaSalObs
                oRsFox.Fields!em_omedico = lcMedicoObs
                oRsFox.Fields!em_ocama = lcCamaObs
                oRsFox.Fields!em_ocdx = lcDxObs
            End If
            oRsFox.Fields!em_destino = lcDestinoAtencion
            oRsFox.Fields!em_referido = lcReferido
            oRsFox.Fields!em_condeg = rsReporte.Fields!codigoCondicionAltaSEM
            oRsFox.Fields!em_fechareg = Format(Date, "mm/dd/yyyy")
            oRsFox.Fields!em_usuario = lcUsuario
            oRsFox.Update
        Else
            '*********************************Hospitalizacion
            oRsFox.AddNew
            oRsFox.Fields!numMov = lnNumMov
            oRsFox.Fields!cod_disa = lcBuscaParametro.SeleccionaFilaParametro(239)
            oRsFox.Fields!cod_red = lcBuscaParametro.SeleccionaFilaParametro(240)
            oRsFox.Fields!cod_mred = lcBuscaParametro.SeleccionaFilaParametro(241)
            oRsFox.Fields!cod_est = Trim(str(Val(lcBuscaParametro.SeleccionaFilaParametro(208))))
            If Not IsNull(rsReporte.Fields!NroHistoriaClinica) Then
              oRsFox.Fields!numHc = Left(Trim(str(rsReporte.Fields!NroHistoriaClinica)), 7)
            Else
              oRsFox.Fields!numHc = ""
            End If
            oRsFox.Fields!apenom = Left(Trim(rsReporte.Fields!ApellidoPaterno) & " " & Trim(rsReporte.Fields!ApellidoMaterno) & " " & Trim(rsReporte.Fields!PrimerNombre) & " " & IIf(IsNull(rsReporte.Fields!SegundoNombre), "", Trim(rsReporte.Fields!SegundoNombre)), 50)
            oRsFox.Fields!Sexo = Trim(str(rsReporte.Fields!idTipoSexo))
            oRsFox.Fields!Edad = rsReporte.Fields!Edad
            oRsFox.Fields!TipoEdad = rsReporte.Fields!tipoEdadSEM
            oRsFox.Fields!Ubigeo = lcBuscaParametro.SeleccionaFilaParametro(242)
            oRsFox.Fields!fecIng = rsReporte.Fields!FechaIngreso  'Format(rsReporte.Fields!FechaIngreso, "mm/dd/yyyy")
            oRsFox.Fields!fecegr = rsReporte.Fields!FechaEgreso    'Format(rsReporte.Fields!FechaEgreso, "mm/dd/yyyy")
            oRsFox.Fields!totalest = lnEstancia
            oRsFox.Fields!Servicio = IIf(IsNull(rsReporte.Fields!codigoServicioSem), "xx", rsReporte.Fields!codigoServicioSem)
            oRsFox.Fields!condicion = rsReporte.Fields!codigoCondicionAltaSEM
            oRsFox.Fields!seguro = rsReporte.Fields!CodigoFuenteFinanciamientoSEM
            oRsFox.Fields!coddiag = lcDx1
            oRsFox.Fields!coddiag2 = lcDx2
            oRsFox.Fields!ubicacionh = IIf(IsNull(rsReporte.Fields!ubicacionSEM), "xxxxx", rsReporte.Fields!ubicacionSEM)
            oRsFox.Fields!medicoh = IIf(IsNull(rsReporte.Fields!Colegiatura), "xxxxx", Left(Trim(rsReporte.Fields!Colegiatura), 5))
            oRsFox.Fields!Estado = 1
            oRsFox.Fields!FechaReg = Format(Date, "mm/dd/yyyy")
            oRsFox.Update
        End If
        lnNumMov = lnNumMov + 1
        rsReporte.MoveNext
      Loop
      Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "A", 0, "Atenciones", oConexion, 500 + 140, mo_lcNombrePc, "Exportar al SEM: " & lcFechaInicio & " al " & lcFechaFinal)       '500+ ListBarReporte.idReporte
    End If
    oRsFox.Close
    oConexionFox.Close
    oConexion.Close
    Set oConexion = Nothing
    Set oConexionFox = Nothing
    Set oRsFox = Nothing
    Exit Sub
ErrorProceso:
    MsgBox Err.Description & Chr(13) & lcParaMsgbox    'debb 05/12/2012
    Resume
End Sub









Private Function isValidField(obj_Field As ADODB.Field) As Boolean
    With obj_Field

        On Error GoTo error_handler
        Select Case obj_Field.Type
            Case adBinary, adIDispatch, adIUnknown, adUserDefined
                isValidField = False
            ' -- Campo válido
            Case Else
                isValidField = True
        End Select
    End With
Exit Function
error_handler:
End Function

'debb-29/04/2016
Sub GeneraTXT_una_linea(oRsTmp1 As Recordset, lcArchivoTXT As String, lnBarraTotal As Long, lnBarraNumero As Integer, _
                        lcSeparadorCampo As String, lcSeparadorSaltoLinea As String)
        On Error GoTo ErrGenTXT
        Dim lniFreeFile As Integer, lnUltimoCampo As Integer, lnCont As Integer, lcLinea As String
        Dim lcCampoValor As String, lnContador As Long
        
        
        lniFreeFile = FreeFile
        Open lcArchivoTXT For Output As #lniFreeFile
        lnUltimoCampo = oRsTmp1.Fields.Count
        '
        lnContador = 1
        '
        oRsTmp1.MoveFirst
        lcLinea = ""
        Do While Not oRsTmp1.EOF
           DoEvents
           lnContador = lnContador + 1
           '
           
           For lnCont = 0 To lnUltimoCampo - 1
               lcCampoValor = ""
               If Not IsNull(oRsTmp1.Fields(lnCont).Value) Then
                    Select Case oRsTmp1.Fields(lnCont).Type
                    Case adDate
                         lcCampoValor = lcCampoValor & oRsTmp1.Fields(lnCont)
                    Case adChar, adVarChar
                         lcCampoValor = Trim(oRsTmp1.Fields(lnCont))
                    Case adDouble, adInteger
                         lcCampoValor = Trim(str((oRsTmp1.Fields(lnCont))))
                    Case Else
                         lcCampoValor = lcCampoValor & oRsTmp1.Fields(lnCont)
                    End Select
               End If
               If lnUltimoCampo - 1 = lnCont Then
                  lcLinea = lcLinea & lcCampoValor & lcSeparadorSaltoLinea
               Else
                  lcLinea = lcLinea & lcCampoValor & lcSeparadorCampo
               End If
           Next
           
           oRsTmp1.MoveNext
        Loop
        lcLinea = Left(lcLinea, Len(lcLinea) - 1)
        Print #lniFreeFile, lcLinea
        Close #lniFreeFile
        Exit Sub
ErrGenTXT:
        MsgBox Err.Description
       ' Resume
End Sub


'debb-22/04/2016
Sub GeneraTXT(oRsTmp1 As Recordset, lcArchivoTXT As String, lnBarraTotal As Long, lnBarraNumero As Integer, _
                                                lcSeparadorCampo As String, lcSeparadorSaltoLinea As String)
        On Error GoTo ErrGenTXT
        Dim lniFreeFile As Integer, lnUltimoCampo As Integer, lnCont As Integer, lcLinea As String
        Dim lcCampoValor As String, lnContador As Long, lcDNI As String
        'Const lcSeparadorCampo As String = "|":  Dim lcSeparadorSaltoLinea As String
        Dim lcDisa As String, lcFormato As String, lcNumero As String
        Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
        Dim oConexionExterna As New Connection
        Dim oRsAfiliadosSIS As New Recordset
        If Right(lcArchivoTXT, 13) = "Pacientes.txt" Then
            oConexionExterna.CommandTimeout = 300
            oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
            oConexionExterna.CursorLocation = adUseClient
        End If
        
        'lcSeparadorSaltoLinea = Chr(10)
        lniFreeFile = FreeFile
        Open lcArchivoTXT For Output As #lniFreeFile
        lnUltimoCampo = oRsTmp1.Fields.Count
        '
        lnContador = 1
        '
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
           '
           lcLinea = ""
           lcDNI = ""
           For lnCont = 0 To lnUltimoCampo - 1
               If UCase(oRsTmp1.Fields(lnCont).Name) = "CABNROENVIOALSIS" Then
                  If IsNull(oRsTmp1.Fields(lnCont).Value) Then
                     oRsTmp1.Fields(lnCont).Value = 1
                  Else
                     oRsTmp1.Fields(lnCont).Value = Trim(str(Val(oRsTmp1.Fields(lnCont).Value) + 1))
                  End If
                  oRsTmp1.Update
               End If
               lcCampoValor = ""
               
               If Not IsNull(oRsTmp1.Fields(lnCont).Value) Then
                    Select Case oRsTmp1.Fields(lnCont).Type
                    Case adDate
                         lcCampoValor = lcCampoValor & oRsTmp1.Fields(lnCont)
                    Case adChar, adVarChar
                         lcCampoValor = Trim(oRsTmp1.Fields(lnCont))
                         If UCase(oRsTmp1.Fields(lnCont).Name) = "DXCODIGO" Or UCase(oRsTmp1.Fields(lnCont).Name) = "CODIGO" And oRsTmp1.Fields(lnCont).DefinedSize = 5 Then
                             lcCampoValor = sighentidades.DevuelveCodigoDxSinPUNTO(Left(oRsTmp1.Fields(lnCont).Value, 5))
                         End If
                         If lcCampoValor = lcSeparadorCampo Then
                            lcCampoValor = "_"
                         End If
                    Case adDouble, adInteger
                         lcCampoValor = Trim(str((oRsTmp1.Fields(lnCont))))
                    Case Else
                         lcCampoValor = lcCampoValor & oRsTmp1.Fields(lnCont)
                    End Select
               End If
               If (lnUltimoCampo - 1) = lnCont Then
                  If lcSeparadorSaltoLinea = "" Then
                     lcLinea = lcLinea & lcCampoValor
                  Else
                     lcLinea = lcLinea & lcCampoValor & lcSeparadorCampo
                  End If
               Else
                  lcLinea = lcLinea & lcCampoValor & lcSeparadorCampo
               End If
               If Right(lcArchivoTXT, 13) = "Pacientes.txt" And UCase(oRsTmp1.Fields(lnCont).Name) = "NRODOCUMENTO" Then
                  lcDNI = lcCampoValor
               End If
           Next
           '
           If Right(lcArchivoTXT, 13) = "Pacientes.txt" Then
                lcDisa = "": lcFormato = "": lcNumero = ""
                If Len(lcDNI) = 8 Then
                     Set oRsAfiliadosSIS = mo_ReglasSISgalenhos.SisFiliacionesXdni(lcDNI, oConexionExterna)
                     If oRsAfiliadosSIS.RecordCount > 0 Then
                        lcDisa = Trim(oRsAfiliadosSIS.Fields!AfiliacionDisa)
                        lcFormato = Trim(oRsAfiliadosSIS.Fields!AfiliacionTipoFormato)
                        lcNumero = Trim(oRsAfiliadosSIS.Fields!AfiliacionNroFormato)
                     End If
                     oRsAfiliadosSIS.Close
                End If
                lcLinea = lcLinea & lcDisa & lcSeparadorCampo & lcFormato & lcSeparadorCampo & lcNumero & lcSeparadorCampo
           End If
           Print #lniFreeFile, lcLinea
           oRsTmp1.MoveNext
        Loop
        Close #lniFreeFile
        Set oConexionExterna = Nothing
        Set oRsAfiliadosSIS = Nothing
        Exit Sub
ErrGenTXT:
        MsgBox Err.Description
        'Resume
End Sub




Sub CreandoRutaEnDisco(lcRutaTemporal As String)
    Dim lcSql As String
    On Error Resume Next
    lcSql = "c:"
    sighentidades.ejecutarComando lcSql
    lcSql = "cd \"
    sighentidades.ejecutarComando lcSql
    lcSql = "md " & lcRutaTemporal
    sighentidades.ejecutarComando lcSql
    lcSql = "cd " & lcRutaTemporal
    sighentidades.ejecutarComando lcSql
    lcSql = "del *.*"

End Sub
Function LeerArchivoTexto(nombreFichero As String) As String
    Dim numlib As Integer, isOpen As Boolean
    On Error GoTo Manejador_Error
    ' Obtengo el siguiente número libre de archivo
    numlib = FreeFile()
    Open nombreFichero For Input As #numlib
    ' Se ha abierto el fichero sin problemas
    isOpen = True
    ' Leo todo el contenido en una única operación
    LeerArchivoTexto = Input(LOF(numlib), numlib)
    ' Cierro el archivo
Manejador_Error:
    If isOpen Then Close #numlib
    If Err Then Err.Raise Err.Number, , Err.Description
End Function


Public Function EstadoDeArchivo(ByVal Archivo As String) As Boolean
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(Archivo)) Then
    EstadoDeArchivo = True
    Else
    EstadoDeArchivo = False
    End If
End Function


'Actualiza valores y Devuelve percentil de la ATENCION ACTUAL DEL PACIENTE
'debb-2/3/2015
Sub CalculaPercentiles(lnPesoKg As Double, lnTallaCM As Long, ml_EdadEnMeses As Long, _
                       ml_idTipoSexo As Long, lnEdadEnAniosEnAtencion As Integer, _
                       ByRef lnPercentilPE As Long, ByRef lnPercentilTE As Long, ByRef lnPercentilPT As Long, ByRef lnPercentilIMC As Double, _
                       ByRef lnPercentilPE_Z As Double, ByRef lnPercentilTE_Z As Double, ByRef lnPercentilPT_Z As Double, ByRef lnPercentilIMC_Z As Double)
    Const lnPercentilNull As Long = 0
    lnPercentilPE = lnPercentilNull: lnPercentilTE = lnPercentilNull: lnPercentilPT = lnPercentilNull: lnPercentilIMC = lnPercentilNull
    lnPercentilPE_Z = lnPercentilNull: lnPercentilTE_Z = lnPercentilNull: lnPercentilPT_Z = lnPercentilNull: lnPercentilIMC_Z = lnPercentilNull
    If lnPesoKg > 0 And lnTallaCM > 0 Then
       On Error Resume Next
       Dim EXL As Excel.Application
       Set EXL = New Excel.Application
       Dim W As Excel.Workbook
       Dim lnEdadUSAmaxima
       lnEdadUSAmaxima = 5
       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
           Set W = EXL.Workbooks.Open(App.Path & "\Plantillas\cred.xls")       'usa
       Else
           Set W = EXL.Workbooks.Open(App.Path & "\Plantillas\cred who.xls")    'oms
       End If
       Dim s As Excel.Worksheet
       Dim lnEdadEnMesesMasPuntoCinco As Double, lnMinimo As Double, lnMaximo As Double, lnIMC As Double
       Dim lnTallaEnCmMasPuntoCinco As Double
       lnEdadEnMesesMasPuntoCinco = ml_EdadEnMeses + 0.5
       lnTallaEnCmMasPuntoCinco = lnTallaCM + 0.5
       'Peso Edad
       Set s = W.Sheets("P-E")
       
       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(243, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilPE = lnPercentilNull
          s.Cells(246, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnEdadEnMesesMasPuntoCinco
          s.Cells(247, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnPesoKg
          lnPercentilPE = s.Cells(255, IIf(ml_idTipoSexo = 1, 3, 19)).Value
          lnPercentilPE_Z = s.Cells(254, IIf(ml_idTipoSexo = 1, 3, 19)).Value
       Else
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(1858, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilPE = lnPercentilNull
          s.Cells(1861, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnEdadEnMesesMasPuntoCinco
          s.Cells(1862, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnPesoKg
          lnPercentilPE = s.Cells(1870, IIf(ml_idTipoSexo = 1, 4, 20)).Value
          lnPercentilPE_Z = s.Cells(1869, IIf(ml_idTipoSexo = 1, 4, 20)).Value
       End If





       'Talla Edad
       Set s = W.Sheets("T-E")
       'If lnEdadEnAniosEnAtencion <= lnEdadUSAmaxima Then
       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(243, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilTE = lnPercentilNull
          s.Cells(246, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnEdadEnMesesMasPuntoCinco
          s.Cells(247, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnTallaCM
          lnPercentilTE = s.Cells(255, IIf(ml_idTipoSexo = 1, 3, 19)).Value
          lnPercentilTE_Z = s.Cells(254, IIf(ml_idTipoSexo = 1, 3, 19)).Value
       Else
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(1858, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilTE = lnPercentilNull
          s.Cells(1861, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnEdadEnMesesMasPuntoCinco
          s.Cells(1862, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnTallaCM
          lnPercentilTE = s.Cells(1870, IIf(ml_idTipoSexo = 1, 4, 20)).Value
          lnPercentilTE_Z = s.Cells(1869, IIf(ml_idTipoSexo = 1, 4, 20)).Value
       End If
       'Peso Talla
       Set s = W.Sheets("P-T")
       'If lnEdadEnAniosEnAtencion <= lnEdadUSAmaxima Then
       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(61, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilPT = lnPercentilNull
          s.Cells(64, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnTallaEnCmMasPuntoCinco
          s.Cells(65, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnPesoKg
          lnPercentilPT = s.Cells(73, IIf(ml_idTipoSexo = 1, 3, 19)).Value
          lnPercentilPT_Z = s.Cells(72, IIf(ml_idTipoSexo = 1, 3, 19)).Value
       Else
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(652, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilPT = lnPercentilNull
          s.Cells(655, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnTallaEnCmMasPuntoCinco
          s.Cells(656, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnPesoKg
          lnPercentilPT = s.Cells(664, IIf(ml_idTipoSexo = 1, 4, 20)).Value
          lnPercentilPT_Z = s.Cells(663, IIf(ml_idTipoSexo = 1, 4, 20)).Value
       End If

       'Edad IMC
       lnIMC = Round((lnPesoKg / (lnTallaCM * lnTallaCM)), 0)
       Set s = W.Sheets("E-IMC")
       'If lnEdadEnAniosEnAtencion <= lnEdadUSAmaxima Then
       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 19)).Value
          lnMaximo = s.Cells(220, IIf(ml_idTipoSexo = 1, 2, 19)).Value
          lnPercentilIMC = lnPercentilNull
          s.Cells(223, IIf(ml_idTipoSexo = 1, 4, 21)).Value = lnEdadEnMesesMasPuntoCinco
          s.Cells(224, IIf(ml_idTipoSexo = 1, 4, 21)).Value = lnIMC
          lnPercentilIMC = s.Cells(232, IIf(ml_idTipoSexo = 1, 3, 20)).Value
          lnPercentilIMC_Z = s.Cells(231, IIf(ml_idTipoSexo = 1, 3, 20)).Value
       Else
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(1858, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilIMC = lnPercentilNull
          s.Cells(1861, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnEdadEnMesesMasPuntoCinco
          s.Cells(1862, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnIMC
          lnPercentilIMC = s.Cells(1870, IIf(ml_idTipoSexo = 1, 4, 20)).Value
          lnPercentilIMC_Z = s.Cells(1869, IIf(ml_idTipoSexo = 1, 4, 20)).Value
       End If
       W.Close False
       Set s = Nothing
       Set W = Nothing
       Set EXL = Nothing
    End If
End Sub


'debb-15/06/2016
Sub ExportaSEM2016(lcFechaInicio As String, lcFechaFinal As String, lcCorrelativo As String, _
                    ml_Valor As Integer, ml_Tabla As String, lbMuestraUsuario As Boolean, lbConsideraDxIngresos As Boolean)
    On Error GoTo ErrorProceso
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_AdminAdmision As New ReglasAdmision
    Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_ReglasServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
    Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
    Dim mo_ReglasHoteleria As New SIGHNegocios.ReglasHoteleria
    Dim oRsTmp2 As New Recordset
    Dim lcMensajes99 As String
    Dim rsReporte As New Recordset
    Dim oRsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oRsFox As New Recordset
    Dim rsServicio As New Recordset
    Dim oRsErrores As New Recordset
    Dim lbPrimerError As Boolean, lbNoExisteH As Boolean
    Dim oConexionFox As New Connection
    Dim oConexion As New Connection
    Dim lcDx1 As String, lcDx2 As String
    Dim lnNumMov As Long, lcSql As String, lnEstancia As Integer
    Dim lcFec1 As String, lcFec2 As String
    Dim lcMotivo As String, lcDestinoAtencion As String, lcReferido As String, lcUsuario As String
    Dim lcDireccionDomicilio As String, lcNombreAcompaniante As String
    Dim lnEsObservacionEmergencia As Integer, lcServicio As String, lcMedico As String, lcServicioObs As String
    Dim lcMedicoObs As String, ldFechaIngObs As Date, lcHoraIngObs As String, lcCamaObs As String, ldFechaSalObs As Date
    Dim lcHoraSalObs As String, lcDxObs As String, lcHoraEstanciaMax As String, lnTotal As Long
    Dim lcRenaes As String, lcUbigeoEESS As String, lcDisa As String, lcRed As String, lcMred As String
    Dim lcCodSal As String, lnRNmuertos As Long, lnRNvivos As Long, ldFechaNacimiento As Date, lcProfParto As String
    Dim lcCpt1 As String, lcCpt2 As String, lcCpt3 As String, lcCpt4 As String, lcDxTipo1 As String, lcDxTipo2 As String
    Dim lcDNIacompainante As String, lcNroEmergencia As String, lcUsuario9 As String, ldUsuarioFecha9 As Date
    Dim lcDx3 As String, lcDx4 As String, lcDxTipo3 As String, lcDxTipo4 As String, lcDiagn As String

    Const lcEtniaDefault As String = "80"

    If ml_Valor = 2 Then
       FileCopy App.Path + "\archivos\sem_em1.dbf", App.Path + "\archivos\sem_emer.dbf"
    Else
       FileCopy App.Path + "\archivos\sem_eg1.dbf", App.Path + "\archivos\sem_egre.dbf"
    End If
    
    lcRenaes = Right("00000000" & lcBuscaParametro.SeleccionaFilaParametro(208), 10)
    lcUbigeoEESS = lcBuscaParametro.SeleccionaFilaParametro(242)
    lcHoraEstanciaMax = lcBuscaParametro.SeleccionaFilaParametro(201)
    lcDisa = lcBuscaParametro.SeleccionaFilaParametro(239)
    lcRed = lcBuscaParametro.SeleccionaFilaParametro(240)
    lcMred = lcBuscaParametro.SeleccionaFilaParametro(241)
        '
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
        '
    oConexionFox.CommandTimeout = 300
    lcParaMsgbox = " (Eliminar Datos)"
    
    lcMensajes99 = ""
    
    If ml_Valor = 2 Then
        '*********************************Emergencia
        
        oConexionFox.Open "DSN=SEM"
        ml_Tabla = "sem_emer"
        mo_ReglasComunes.AbreTablaFox oRsFox, oConexionFox, ml_Tabla, "", ""
        If oRsFox.RecordCount > 0 Then
          oRsFox.MoveFirst
          Do While Not oRsFox.EOF
            If oRsFox.Fields!fecate >= CDate(lcFechaInicio) And oRsFox.Fields!fecate <= CDate(lcFechaFinal) Then
              oRsFox.Delete
              oRsFox.Update
            End If
            oRsFox.MoveNext
          Loop
        End If
        oRsFox.Close
        '
        mo_ReglasComunes.AbreTablaFox oRsFox, oConexionFox, ml_Tabla, "", ""
    Else
        '*********************************Hospitalizacion
        oConexionFox.Open "DSN=SEM"
        ml_Tabla = "sem_egre"
        mo_ReglasComunes.AbreTablaFox oRsFox, oConexionFox, ml_Tabla, "", ""
        If oRsFox.RecordCount > 0 Then
          oRsFox.MoveFirst
          Do While Not oRsFox.EOF
            If oRsFox.Fields!fecIng >= CDate(lcFechaInicio) And oRsFox.Fields!fecIng <= CDate(lcFechaFinal) Then
              oRsFox.Delete
              oRsFox.Update
            End If
            oRsFox.MoveNext
          Loop
        End If
        oRsFox.Close
        '
        mo_ReglasComunes.AbreTablaFox oRsFox, oConexionFox, ml_Tabla, "", ""
    End If
 '   DoEvents
    '
    lcUsuario = " "
    Set oRsTmp1 = mo_ReglasComunes.EmpleadosSeleccionarPorIdEmpleado(ml_idUsuario)
    If oRsTmp1.RecordCount > 0 Then
       If Not IsNull(oRsTmp1.Fields!Usuario) Then
          lcUsuario = oRsTmp1.Fields!Usuario
       End If
    End If
    oRsTmp1.Close
    '
    Set rsReporte = mo_ReglasAdmision.AtencionesSeleccionarHospEmergPorFechasDeEgresoMedico(CDate(lcFechaInicio), CDate(lcFechaFinal), "FECHAEGRESO")
    rsReporte.Filter = "codigoCondicionAltaSEM<>null and idTipoServicio=" & ml_Valor
    lnTotal = rsReporte.RecordCount
    
    If lnTotal = 0 Then
      MsgBox "No existe Información  con esos Datos", vbInformation, "Mensaje"
    Else
      '
      lbPrimerError = True
      With oRsErrores
            .Fields.Append "Historia", adInteger
            .Fields.Append "FechaIng", adDate
            .Fields.Append "Error", adVarChar, 50, adFldIsNullable
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .Open
      End With
      '
      lnNumMov = Val(lcCorrelativo)
      rsReporte.MoveFirst
      Do While Not rsReporte.EOF
If rsReporte!idCuentaAtencion = 204772 Then
lcParaMsgbox = ""
End If
        '
        If rsReporte!IdEtnia <> lcEtniaDefault Then
           lcMensajes99 = lcMensajes99 & " HC=" & Trim(str(rsReporte.Fields!NroHistoriaClinica)) & " tiene ETNIA diferente a: " & lcEtniaDefault & Chr(13)
        End If
        If Len(Trim(rsReporte.Fields!CodigoFuenteFinanciamientoSEM)) <> 2 Then
           lcMensajes99 = lcMensajes99 & " N° Cuenta=" & Trim(str(rsReporte.Fields!idCuentaAtencion)) & " tiene FUENTE FINANCIAMIENTO diferente a 2 dígitos, vefiricar FUENTE FINANCIAMIENTO en CUENTA y luego en 'Fact-Config->Fuente Financiamiento->código sistema SEM'" & Chr(13)
        End If
        '
        If lbPrimerError = True Then
           lbPrimerError = False
           oRsErrores.AddNew
           oRsErrores.Fields!historia = rsReporte!NroHistoriaClinica
           oRsErrores.Fields!fechaIng = rsReporte!FechaIngreso
           oRsErrores.Fields!Error = Space(1)
           oRsErrores.Update
        Else
           lbNoExisteH = True
           oRsErrores.MoveFirst
           Do While Not oRsErrores.EOF
              If oRsErrores!historia = rsReporte!NroHistoriaClinica And oRsErrores!fechaIng = rsReporte!FechaIngreso Then
                 lbNoExisteH = False
                 Exit Do
              End If
              oRsErrores.MoveNext
           Loop
           If lbNoExisteH = True Then
                oRsErrores.AddNew
                oRsErrores.Fields!historia = rsReporte!NroHistoriaClinica
                oRsErrores.Fields!fechaIng = rsReporte!FechaIngreso
                oRsErrores.Fields!Error = Space(1)
           Else
           '     oRsErrores.Fields!Error = "(misma F.Ingreso)"
           End If
           oRsErrores.Update
        End If
        '
        lcParaMsgbox = " -->Cuenta: " & rsReporte.Fields!idCuentaAtencion
'        'solo Hospitalizados
'        lcDiagn = "/"
'        lcDx1 = "": lcDx2 = "": lcDxTipo1 = "": lcDxTipo2 = "": lcDx3 = "": lcDx4 = "": lcDxTipo3 = "": lcDxTipo4 = ""
'        Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(rsReporte.Fields!idAtencion, 3, oConexion)
'        oRsTmp.Filter = "IdTipoDiagnostico=301"   'vivos-princial
'        If oRsTmp.RecordCount > 0 Then
'          '*****vivos
'          oRsTmp.MoveFirst
'          lcDx1 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
'          lcDxTipo1 = "D"
'          lcDiagn = lcDiagn & lcDx1 & "/"
'          oRsTmp.Filter = "IdTipoDiagnostico=302"
'          If oRsTmp.RecordCount > 0 Then
'            oRsTmp.MoveFirst
'            lcDx2 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
'            lcDxTipo2 = "P"
'            lcDiagn = lcDiagn & lcDx2 & "/"
'          End If
'          oRsTmp.Close
'          Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(rsReporte.Fields!idAtencion, 2, oConexion)
'          If oRsTmp.RecordCount > 0 Then
'            oRsTmp.MoveFirst
'            lcDx3 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
'            lcDxTipo3 = "P"
'            lcDiagn = lcDiagn & lcDx3 & "/"
'            oRsTmp.MoveNext
'            If Not oRsTmp.EOF Then
'                lcDx4 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
'                lcDxTipo4 = "P"
'                lcDiagn = lcDiagn & lcDx4 & "/"
'            End If
'          End If
'        Else
'          '******fallecidos
'          Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(rsReporte.Fields!idAtencion, 4, oConexion)
'          If oRsTmp.RecordCount > 0 Then
'                oRsTmp.Filter = "IdTipoDiagnostico=303"    'Final
'                If oRsTmp.RecordCount > 0 Then
'                  oRsTmp.MoveFirst
'                  lcDx1 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
'                  lcDxTipo1 = "D"
'                  lcDiagn = lcDiagn & lcDx1 & "/"
'                End If
'                oRsTmp.Filter = "IdTipoDiagnostico=304 or IdTipoDiagnostico=305"    'fallecidos - Final
'                If oRsTmp.RecordCount > 0 Then
'                  oRsTmp.MoveFirst
'                  lcDx2 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
'                  lcDxTipo2 = "P"
'                  lcDiagn = lcDiagn & lcDx2 & "/"
'                End If
'          End If
'        End If
'        oRsTmp.Filter = ""
'        If oRsTmp.RecordCount > 0 Then
'           oRsTmp.MoveFirst
'           Do While Not oRsTmp.EOF
'              If InStr(lcDiagn, Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)) = 0 Then
'                 If lcDx3 = "" Then
'                    lcDx3 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
'                    lcDxTipo3 = IIf(oRsTmp!idTipoDiagnostico = 301 Or oRsTmp!idTipoDiagnostico = 303, "D", "P")
'                    lcDiagn = lcDiagn & lcDx3 & "/"
'                 Else
'                    lcDx4 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
'                    lcDxTipo4 = IIf(oRsTmp!idTipoDiagnostico = 301 Or oRsTmp!idTipoDiagnostico = 303, "D", "P")
'                    Exit Do
'                 End If
'              End If
'              oRsTmp.MoveNext
'           Loop
'        End If
'        oRsTmp.Close
        lcDiagn = "/"
        lcDx1 = "": lcDx2 = "": lcDxTipo1 = "": lcDxTipo2 = "": lcDx3 = "": lcDx4 = "": lcDxTipo3 = "": lcDxTipo4 = ""
        Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(rsReporte.Fields!idAtencion, 3, oConexion)
        oRsTmp.Filter = "IdTipoDiagnostico=301"   'vivos-princial
        If oRsTmp.RecordCount > 0 Then
          '*****vivos
'          oRsTmp.MoveFirst
'          lcDx1 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
'          lcDxTipo1 = "D"
'          lcDiagn = lcDiagn & lcDx1 & "/"
          oRsTmp.MoveFirst
          Do While Not oRsTmp.EOF
            If lcDx1 = "" Then
                lcDx1 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                lcDxTipo1 = "D"
                lcDiagn = lcDiagn & lcDx1 & "/"
            ElseIf lcDx2 = "" Then
                lcDx2 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                lcDxTipo2 = "D"
                lcDiagn = lcDiagn & lcDx2 & "/"
            ElseIf lcDx3 = "" Then
                lcDx3 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                lcDxTipo3 = "D"
                lcDiagn = lcDiagn & lcDx3 & "/"
            End If
            oRsTmp.MoveNext
          Loop
          '
          oRsTmp.Filter = "IdTipoDiagnostico=302"
          If oRsTmp.RecordCount > 0 Then
            oRsTmp.MoveFirst
            Do While Not oRsTmp.EOF
                If lcDx2 = "" Then
                    lcDx2 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                    lcDxTipo2 = "P"
                    lcDiagn = lcDiagn & lcDx2 & "/"
                ElseIf lcDx3 = "" Then
                    lcDx3 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                    lcDxTipo3 = "P"
                    lcDiagn = lcDiagn & lcDx3 & "/"
                ElseIf lcDx4 = "" Then
                    lcDx4 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                    lcDxTipo4 = "P"
                    lcDiagn = lcDiagn & lcDx4 & "/"
                End If
                oRsTmp.MoveNext
            Loop
          End If
          oRsTmp.Close
          
          Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(rsReporte.Fields!idAtencion, 2, oConexion)
          If lbConsideraDxIngresos = True Then
          If oRsTmp.RecordCount > 0 Then
            oRsTmp.MoveFirst
            Do While Not oRsTmp.EOF
                If lcDx2 = "" Then
                    lcDx2 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                    lcDxTipo2 = "P"
                    lcDiagn = lcDiagn & lcDx2 & "/"
                ElseIf lcDx3 = "" Then
                    lcDx3 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                    lcDxTipo3 = "P"
                    lcDiagn = lcDiagn & lcDx3 & "/"
                ElseIf lcDx4 = "" Then
                    lcDx4 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                    lcDxTipo4 = "P"
                    lcDiagn = lcDiagn & lcDx4 & "/"
                End If
                oRsTmp.MoveNext
            Loop
          End If
          End If
        Else
          '******fallecidos
          Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(rsReporte.Fields!idAtencion, 4, oConexion)
          If oRsTmp.RecordCount > 0 Then
                oRsTmp.Filter = "IdTipoDiagnostico=303"    'Final
                If oRsTmp.RecordCount > 0 Then
                  oRsTmp.MoveFirst
                  lcDx1 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                  lcDxTipo1 = "D"
                  lcDiagn = lcDiagn & lcDx1 & "/"
                End If
                oRsTmp.Filter = "IdTipoDiagnostico=304 or IdTipoDiagnostico=305"    'fallecidos - Final
                If oRsTmp.RecordCount > 0 Then
                  oRsTmp.MoveFirst
                  lcDx2 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                  lcDxTipo2 = "P"
                  lcDiagn = lcDiagn & lcDx2 & "/"
                End If
          End If
        End If
        oRsTmp.Filter = ""
        If lbConsideraDxIngresos = True Then
        If oRsTmp.RecordCount > 0 Then
           oRsTmp.MoveFirst
           Do While Not oRsTmp.EOF
              If InStr(lcDiagn, Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)) = 0 Then
                 If lcDx2 = "" Then
                    lcDx2 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                    lcDxTipo2 = IIf(oRsTmp!idTipoDiagnostico = 301 Or oRsTmp!idTipoDiagnostico = 303, "D", "P")
                    lcDiagn = lcDiagn & lcDx2 & "/"
                 ElseIf lcDx3 = "" Then
                    lcDx3 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                    lcDxTipo3 = IIf(oRsTmp!idTipoDiagnostico = 301 Or oRsTmp!idTipoDiagnostico = 303, "D", "P")
                    lcDiagn = lcDiagn & lcDx3 & "/"
                 ElseIf lcDx4 = "" Then
                    lcDx4 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                    lcDxTipo4 = IIf(oRsTmp!idTipoDiagnostico = 301 Or oRsTmp!idTipoDiagnostico = 303, "D", "P")
                    Exit Do
                 End If
              End If
              oRsTmp.MoveNext
           Loop
        End If
        End If
        oRsTmp.Close
          



        'nacimientos - solo hospitalizacion
        lnRNvivos = 0: lnRNmuertos = 0: ldFechaNacimiento = 0: lcProfParto = ""
        If ml_Valor <> 2 Then
           Set oRsTmp1 = mo_AdminAdmision.AtencionesNacimientosSeleccionarPorAtencion(rsReporte!idAtencion, oConexion)
           If oRsTmp1.RecordCount > 0 Then
              lcProfParto = "1"
              If rsReporte!idTipoEmpleado = 28 Then   '****obstret****
                 lcProfParto = "2"
              End If
              
              oRsTmp1.MoveFirst
              ldFechaNacimiento = oRsTmp1!FechaNacimiento
              Do While Not oRsTmp1.EOF
                 If oRsTmp1!idCondicionRN = 1 Then
                    lnRNvivos = lnRNvivos + 1
                 Else
                    lnRNmuertos = lnRNmuertos + 1
                 End If
                 oRsTmp1.MoveNext
              Loop
           End If
           oRsTmp1.Close
        End If
        'cpt
        lcCpt1 = "": lcCpt2 = "": lcCpt3 = "": lcCpt4 = ""
        Set oRsTmp1 = mo_ReglasAdmision.BuscaOtrosCPTyExamenesparaFormatoHIS(rsReporte!idCuentaAtencion, _
                                                   sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion, oConexion)
        If oRsTmp1.RecordCount > 0 Then
             oRsTmp1.MoveFirst
             Do While Not oRsTmp1.EOF
               If Len(Trim(oRsTmp1.Fields("codigo").Value)) = 5 Then
                    If lcCpt1 = "" Then
                       lcCpt1 = Trim(oRsTmp1.Fields("codigo").Value)
                    ElseIf lcCpt2 = "" Then
                       lcCpt2 = Trim(oRsTmp1.Fields("codigo").Value)
                    ElseIf lcCpt3 = "" Then
                       lcCpt3 = Trim(oRsTmp1.Fields("codigo").Value)
                    Else
                       lcCpt4 = Trim(oRsTmp1.Fields("codigo").Value)
                    End If
               End If
               oRsTmp1.MoveNext
             Loop
        End If
        oRsTmp1.Close
        
        
        '
        lcMotivo = " "
        Set oRsTmp1 = mo_AdminAdmision.AtencionesEmergenciaXidAtencion(rsReporte.Fields!idAtencion, oConexion)
        If oRsTmp1.RecordCount > 0 Then
           If Not IsNull(oRsTmp1.Fields!motivoSEM) Then
              lcMotivo = oRsTmp1.Fields!motivoSEM
           End If
        End If
        oRsTmp1.Close
        '
        lcDireccionDomicilio = "": lcNombreAcompaniante = "": lcDNIacompainante = "0": lcNroEmergencia = ""
        Set oRsTmp1 = mo_AdminAdmision.atencionesDatosAdicionalesXfiltro("dbo.AtencionesDatosAdicionales.idAtencion=" & rsReporte.Fields!idAtencion, oConexion)
        If oRsTmp1.RecordCount > 0 Then
           If Not IsNull(oRsTmp1.Fields!DireccionDomicilio) Then
              lcDireccionDomicilio = oRsTmp1.Fields!DireccionDomicilio
           End If
           If Not IsNull(oRsTmp1.Fields!NombreAcompaniante) Then
              lcNombreAcompaniante = oRsTmp1.Fields!NombreAcompaniante
           End If
           If Not IsNull(oRsTmp1.Fields!AcompanianteDNI) Then
              lcDNIacompainante = oRsTmp1.Fields!AcompanianteDNI
           End If
           If Not IsNull(oRsTmp1.Fields!emergenciaCorrelativo) Then
              lcNroEmergencia = Mid(oRsTmp1.Fields!emergenciaCorrelativo, 5, 20)
           End If
        End If
        oRsTmp1.Close
        '
        lcDestinoAtencion = " "
        If Not IsNull(rsReporte.Fields!IdDestinoAtencion) Then
            Set oRsTmp1 = mo_ReglasAdmision.TiposDestinoAtencionXidDestinoAtencion(rsReporte.Fields!IdDestinoAtencion, oConexion)
            If oRsTmp1.RecordCount > 0 Then
               If IsNull(oRsTmp1.Fields!destinoSEM) Then
                    MsgBox "Existen Servicios de Emergencia sin CODIGO SEM", vbInformation, "Mensaje"  'debb 05/12/2012
                    Exit Sub
               Else
                    lcDestinoAtencion = oRsTmp1.Fields!destinoSEM
               End If
            End If
            oRsTmp1.Close
        End If
        '
        lcReferido = " "
        If Not IsNull(rsReporte.Fields!idEstablecimientoDestino) Then
            Set oRsTmp1 = mo_ReglasComunes.EstablecimientosSeleccionarPorFiltro("idEstablecimiento=" & rsReporte.Fields!idEstablecimientoDestino, oConexion)
            If oRsTmp1.RecordCount > 0 Then
               lcReferido = Right("0000000000" & Trim(str(Val(oRsTmp1.Fields!codigo))), 10)
            End If
            oRsTmp1.Close
        End If
        '
        lcCodSal = "."
        If Not (IsNull(rsReporte!idTipoDocumento) Or IsNull(rsReporte!dni) Or IsNull(rsReporte!idColegioHIS)) Then
           lcCodSal = Trim(str(rsReporte!idTipoDocumento)) & Trim(rsReporte!dni) & rsReporte!idColegioHIS
        End If
        '
        If Val(lcRenaes) = 6918 Then      'hosp tarapoto
           lnEstancia = lcBuscaParametro.DiasDelPacienteEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, rsReporte!FechaEgreso, rsReporte!horaEgreso, lcHoraEstanciaMax, True)
        Else
           lnEstancia = lcBuscaParametro.DiasDelPacienteEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, rsReporte!FechaEgreso, rsReporte!horaEgreso, lcHoraEstanciaMax)
        End If
        If lnEstancia = 0 Then lnEstancia = 1
        '
        lcUsuario9 = " ": ldUsuarioFecha9 = 0
        If lbMuestraUsuario = True Then
            Set rsServicio = mo_ReglasSeguridad.AuditoriaFiltrarCitasPorIdAtencion(rsReporte!idAtencion, _
                                   IIf(ml_Valor = 2, sghAdmisionEmergencia, sghAdmisionHospitalizacion), _
                                   oConexion)
            If rsServicio.RecordCount > 0 Then
               lcUsuario9 = IIf(IsNull(rsServicio!dUsuario), "", rsServicio!dUsuario)
               ldUsuarioFecha9 = IIf(IsNull(rsServicio!fechaHora), 0, Format(rsServicio!fechaHora, sighentidades.DevuelveFechaSoloFormato_DMY))
            End If
            rsServicio.Close
        End If
        '
        If ml_Valor = 2 Then
            '*********************************Emergencia
            lcServicio = IIf(IsNull(rsReporte.Fields!codigoServicioSem), ".", rsReporte.Fields!codigoServicioSem)
            lcMedico = IIf(IsNull(rsReporte.Fields!Colegiatura), ".", Left(rsReporte.Fields!Colegiatura, 5))
            lnEsObservacionEmergencia = 0
            If Not IsNull(rsReporte.Fields!EsObservacionEmergencia) Then
               If rsReporte.Fields!EsObservacionEmergencia = True Then
                    lnEsObservacionEmergencia = 1
                    Set oRsTmp1 = mo_ReglasAdmision.AtencionesEstanciaHospitalariaXidAtencionIdServicioIngreso(rsReporte.Fields!IdServicioIngreso, rsReporte.Fields!idAtencion, oConexion)
                    If oRsTmp1.RecordCount > 0 Then
                       If oRsTmp1.Fields!EsObservacionEmergencia = True Then
                          '***El Servicio de Ingreso y el Servicio de Alta son OBSERVACIONES DE EMERGENCIA
                          '***Solo se muestra datos del Servicio de Alta ya no es necesario mostrar los datos
                          '***de OBSERVACIONES
                          lnEsObservacionEmergencia = 0
                       Else
                          '***El servicio de Ingreso es un CONSULTORIO y el servicio de alta es OBSERVACION
                          '***Hubo transferencia desde CONSULTORIO hacia OBSERVACION
                          '***Muestra datos del CONSULTORIO y datos DE OBSERVACION
                          lcServicioObs = lcServicio
                          lcMedicoObs = lcCodSal
                          lcDxObs = lcDx1
                          ldFechaIngObs = oRsTmp1.Fields!fechaDesocupacion
                          lcHoraIngObs = oRsTmp1.Fields!horaDesocupacion
                          'debb 05/12/2012
                          lcCamaObs = " "
                          If IsNull(rsReporte!IdCamaEgreso) Then  'debb 05/12/2012
                               MsgBox "Para la cuenta. " & rsReporte.Fields!idCuentaAtencion & " no se registró cama, siendo un SERVICIO DE OBSERVACION", vbInformation, "Mensaje"
                               Exit Sub
                          End If
                          Set oRsTmp2 = mo_ReglasHoteleria.CamasSeleccionarPorIdentificador(rsReporte!IdCamaEgreso, oConexion)
                          If oRsTmp2.RecordCount > 0 Then
                               lcCamaObs = Trim(oRsTmp2!codigo)
                          End If
                          oRsTmp2.Close
                          'debb 05/12/2012
                          ldFechaSalObs = rsReporte.Fields!FechaEgreso
                          lcHoraSalObs = rsReporte.Fields!horaEgreso
                          lcServicio = IIf(IsNull(oRsTmp1.Fields!codigoServicioSem), ".", oRsTmp1.Fields!codigoServicioSem)
                          lcMedico = IIf(IsNull(oRsTmp1.Fields!Colegiatura), ".", Left(oRsTmp1.Fields!Colegiatura, 5))
                          lnEstancia = DateDiff("h", Format(ldFechaIngObs, "dd/mm/yyyy") & " " & lcHoraIngObs, _
                                        Format(ldFechaSalObs, "dd/mm/yyyy") & " " & lcHoraSalObs)

                       End If
                    End If
                    oRsTmp1.Close
               Else
               
                     '***El servicio de ALTA MEDICA es un CONSULTORIO
                     Set oRsTmp1 = mo_ReglasAdmision.AtencionesEstanciaHospitalariaXidAtencionIdServicioIngreso(0, rsReporte.Fields!idAtencion, oConexion)
                     oRsTmp1.Filter = "esObservacionEmergencia=1"
                     If oRsTmp1.RecordCount > 0 Then
                           lnEsObservacionEmergencia = 1
                           '***El servicio de ALTA MEDICA es un CONSULTORIO
                           '***Hubo al menos una TRANSFERENCIA hacia OBSERVACION (o se ingresó de frente a OBSERVACION)
                           '***Muestra datos DE OBSERVACION
                           
                           lcServicioObs = lcServicio
                           
                           If IsNull(oRsTmp1!lcColegiaturaMedico) Then
                              MsgBox "Para la cuenta. " & rsReporte.Fields!idCuentaAtencion & " no se registró COLEGIATURA DEL MEDICO (buscar la CUENTA ->MODIFICAR -> TRANSFERENCIA, luego ir a PROGRAMACION GENERAL->PROFESIONALES DE SALUD", vbInformation, "Mensaje"
                              Exit Sub
                           End If
                           lcMedicoObs = oRsTmp1!lcColegiaturaMedico
                           
                           lcDxObs = lcDx1
                           ldFechaIngObs = oRsTmp1!fechaOcupacion
                           lcHoraIngObs = oRsTmp1!horaOcupacion
                           'debb 05/12/2012
                           If IsNull(oRsTmp1.Fields!cama) Then  'debb 05/12/2012
                              MsgBox "Para la cuenta. " & rsReporte.Fields!idCuentaAtencion & " no se registró cama, siendo un SERVICIO DE OBSERVACION", vbInformation, "Mensaje"
                              Exit Sub
                           Else
                              lcCamaObs = oRsTmp1.Fields!cama
                           End If
                           'debb 05/12/2012
                           ldFechaSalObs = oRsTmp1!fechaDesocupacion
                           lcHoraSalObs = oRsTmp1!horaDesocupacion
                           'lcServicio = IIf(IsNull(oRsTmp1.Fields!codigoServicioSEM), ".", oRsTmp1.Fields!codigoServicioSEM)
                           'lcMedico = IIf(IsNull(oRsTmp1.Fields!Colegiatura), ".", Left(oRsTmp1.Fields!Colegiatura, 5))
                           lnEstancia = DateDiff("h", Format(ldFechaIngObs, "dd/mm/yyyy") & " " & lcHoraIngObs, _
                                                      Format(ldFechaSalObs, "dd/mm/yyyy") & " " & lcHoraSalObs)
                       
                    End If
                End If
            End If
            
            oRsFox.AddNew
            oRsFox.Fields!renIpress = lcRenaes
            oRsFox.Fields!e_ubig = lcUbigeoEESS
            oRsFox.Fields!cod_disa = lcDisa
            oRsFox.Fields!cod_red = lcRed
            oRsFox.Fields!cod_mred = lcMred
            oRsFox.Fields!fecate = Format(rsReporte.Fields!FechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY)
            oRsFox.Fields!horate = rsReporte!HoraIngreso
            If Val(lcRenaes) = 6918 Then
                oRsFox.Fields!numHc = lcNroEmergencia    'solo HOspital Tarapoto
            Else
                If Not IsNull(rsReporte.Fields!NroHistoriaClinica) Then
                   oRsFox.Fields!numHc = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(str(rsReporte.Fields!NroHistoriaClinica)), False)
                Else
                   oRsFox.Fields!numHc = "sin HC"
                End If
            End If
            If IsNull(rsReporte.Fields!IdDocIdentidad) Or IsNull(rsReporte.Fields!nroDocumento) Then
               oRsFox.Fields!doc_iden = "0"
            ElseIf rsReporte.Fields!IdDocIdentidad >= 1 And rsReporte.Fields!IdDocIdentidad <= 4 Then
               oRsFox.Fields!doc_iden = Trim(str(rsReporte.Fields!IdDocIdentidad)) & Left(rsReporte.Fields!nroDocumento, 10)
            Else
               oRsFox.Fields!doc_iden = "0"
            End If
            oRsFox.Fields!etnia = IIf(IsNull(rsReporte!IdEtnia), "80", rsReporte!IdEtnia)
            oRsFox.Fields!financia = rsReporte.Fields!CodigoFuenteFinanciamientoSEM
            oRsFox.Fields!Sexo = Trim(str(rsReporte.Fields!idTipoSexo))
            oRsFox.Fields!Edad = rsReporte.Fields!Edad
            oRsFox.Fields!TipoEdad = rsReporte.Fields!tipoEdadSEM
            oRsFox.Fields!nomb = Left(Trim(rsReporte.Fields!PrimerNombre) & " " & IIf(IsNull(rsReporte.Fields!SegundoNombre), "", Trim(rsReporte.Fields!SegundoNombre)), 50)
            oRsFox.Fields!apell = Left(Trim(rsReporte.Fields!ApellidoPaterno) & " " & Trim(rsReporte.Fields!ApellidoMaterno), 50)
            oRsFox.Fields!direcc = Left(lcDireccionDomicilio, 100)
            If Not IsNull(rsReporte!IdDistritoDomicilio) Then
               oRsFox.Fields!ubig_resha = Right("0" & Trim(str(rsReporte!IdDistritoDomicilio)), 6)
            Else
               oRsFox.Fields!ubig_resha = "000000"
            End If
            '
            If Not IsNull(rsReporte.Fields!IdDistritoProcedencia) Then
               oRsFox.Fields!ubig_proce = Right("0" & Trim(str(rsReporte.Fields!IdDistritoProcedencia)), 6)
            Else
               If Not IsNull(rsReporte!IdDistritoDomicilio) Then
                  oRsFox.Fields!ubig_proce = Right("0" & Trim(str(rsReporte!IdDistritoDomicilio)), 6)
               Else
                  oRsFox.Fields!ubig_proce = "000000"
               End If
            End If
            '
            oRsFox.Fields!acompana = Left(lcNombreAcompaniante, 50)
            oRsFox.Fields!adoc_iden = lcDNIacompainante
            oRsFox.Fields!motaten = lcMotivo
            oRsFox.Fields!sitocurren = ""
            oRsFox.Fields!UPS = IIf(IsNull(rsReporte.Fields!codigoServicioSem), ".", rsReporte.Fields!codigoServicioSem)
            oRsFox.Fields!coddiag1 = lcDx1
            oRsFox.Fields!tipdiag1 = lcDxTipo1
            oRsFox.Fields!coddiag2 = lcDx2
            oRsFox.Fields!tipdiag2 = lcDxTipo2
            oRsFox.Fields!coddiag3 = lcDx3
            oRsFox.Fields!tipdiag3 = lcDxTipo3
            oRsFox.Fields!coddiag4 = lcDx4
            oRsFox.Fields!tipdiag4 = lcDxTipo4
            oRsFox.Fields!codcpt1 = lcCpt1
            oRsFox.Fields!codcpt2 = lcCpt2
            oRsFox.Fields!codcpt3 = lcCpt3
            oRsFox.Fields!codcpt4 = lcCpt4
            oRsFox.Fields!condicion = IIf(IsNull(rsReporte.Fields!tiposaltaSEM), "", rsReporte.Fields!tiposaltaSEM)
            oRsFox.Fields!fecegr = Format(rsReporte.Fields!FechaEgreso, sighentidades.DevuelveFechaSoloFormato_DMY)
            oRsFox.Fields!horegr = rsReporte.Fields!horaEgreso
            oRsFox.Fields!destino = lcDestinoAtencion
            oRsFox.Fields!des_eess = lcReferido
            oRsFox.Fields!des_ups = ""
            oRsFox.Fields!codpsal = lcCodSal
            oRsFox.Fields!observ = lnEsObservacionEmergencia
            If lnEsObservacionEmergencia = 1 Then
                oRsFox.Fields!ofecing = ldFechaIngObs
                oRsFox.Fields!ohoring = lcHoraIngObs
                oRsFox.Fields!ofecegr = ldFechaSalObs
                oRsFox.Fields!ohoregr = lcHoraSalObs
                oRsFox.Fields!totalest = lnEstancia
                oRsFox.Fields!ocama = lcCamaObs
                oRsFox.Fields!ocodpsal = lcMedicoObs
                oRsFox.Fields!ocoddiag1 = lcDxObs
                oRsFox.Fields!ocoddiag2 = ""
            Else
                oRsFox.Fields!ofecing = 0
                oRsFox.Fields!ohoring = ""
                oRsFox.Fields!ofecegr = 0
                oRsFox.Fields!ohoregr = ""
                oRsFox.Fields!totalest = lnEstancia
                oRsFox.Fields!ocama = ""
                oRsFox.Fields!ocodpsal = ""
                oRsFox.Fields!ocoddiag1 = ""
                oRsFox.Fields!ocoddiag2 = ""
            End If
            oRsFox.Fields!FechaReg = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
            oRsFox.Fields!Estado = "1"
            oRsFox.Fields!Usuario = lcUsuario9
            oRsFox.Fields!usuariof = ldUsuarioFecha9
            oRsFox.Fields!prioridad = " "
            oRsFox.Fields!form_ing = " "
            oRsFox.Update
            
        Else
            '*********************************Hospitalizacion
            oRsFox.AddNew
            oRsFox.Fields!renIpress = lcRenaes
            oRsFox.Fields!e_ubig = lcUbigeoEESS
            oRsFox.Fields!e_cdpto = Left(lcUbigeoEESS, 2)
            oRsFox.Fields!e_cprov = Mid(lcUbigeoEESS, 3, 2)
            oRsFox.Fields!e_cdist = Right(lcUbigeoEESS, 2)
            oRsFox.Fields!cod_disa = lcDisa
            oRsFox.Fields!cod_red = lcRed
            oRsFox.Fields!cod_mred = lcMred
            If Not IsNull(rsReporte.Fields!NroHistoriaClinica) Then
               If Val(lcRenaes) = 6918 Then   'tarapoto
                  oRsFox.Fields!numHc = Right("000000" & Trim(str(rsReporte.Fields!NroHistoriaClinica)), 6)
               Else
                  oRsFox.Fields!numHc = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(str(rsReporte.Fields!NroHistoriaClinica)), False)
               End If
            Else
               oRsFox.Fields!numHc = "sin HC"
            End If
            oRsFox.Fields!nomb = Left(Trim(rsReporte.Fields!PrimerNombre) & " " & IIf(IsNull(rsReporte.Fields!SegundoNombre), "", Trim(rsReporte.Fields!SegundoNombre)), 50)
            oRsFox.Fields!apell = Left(Trim(rsReporte.Fields!ApellidoPaterno) & " " & Trim(rsReporte.Fields!ApellidoMaterno), 50)
            If IsNull(rsReporte.Fields!IdDocIdentidad) Or IsNull(rsReporte.Fields!nroDocumento) Then
               oRsFox.Fields!doc_iden = "0"
            ElseIf rsReporte.Fields!IdDocIdentidad >= 1 And rsReporte.Fields!IdDocIdentidad <= 4 Then
               oRsFox.Fields!doc_iden = Trim(str(rsReporte.Fields!IdDocIdentidad)) & Left(rsReporte.Fields!nroDocumento, 10)
            Else
               oRsFox.Fields!doc_iden = "0"
            End If
            
            
            oRsFox.Fields!etnia = IIf(IsNull(rsReporte!IdEtnia), "80", rsReporte!IdEtnia)
            oRsFox.Fields!Sexo = Trim(str(rsReporte.Fields!idTipoSexo))
            oRsFox.Fields!Edad = rsReporte.Fields!Edad
            oRsFox.Fields!TipoEdad = rsReporte.Fields!tipoEdadSEM
            If IsNull(rsReporte!IdDistritoDomicilio) Then
               oRsFox.Fields!Ubigeo = "000000"
               oRsFox.Fields!cdpto = "00"
               oRsFox.Fields!cprov = "00"
               oRsFox.Fields!cdist = "00"
            Else
               oRsFox.Fields!Ubigeo = Right("0" & Trim(str(rsReporte!IdDistritoDomicilio)), 6)
               oRsFox.Fields!cdpto = Left(Right("0" & Trim(str(rsReporte!IdDistritoDomicilio)), 6), 2)
               oRsFox.Fields!cprov = Mid(Right("0" & Trim(str(rsReporte!IdDistritoDomicilio)), 6), 3, 2)
               oRsFox.Fields!cdist = Right("0" & Trim(str(rsReporte!IdDistritoDomicilio)), 2)
            End If
            oRsFox.Fields!fecIng = Format(rsReporte!FechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY)
            oRsFox.Fields!fecegr = Format(rsReporte.Fields!FechaEgreso, sighentidades.DevuelveFechaSoloFormato_DMY)
            oRsFox.Fields!totalest = lnEstancia
            
            
            oRsFox.Fields!UPS = IIf(IsNull(rsReporte.Fields!codigoServicioSem), "000000", rsReporte.Fields!codigoServicioSem)
            oRsFox.Fields!condicion = IIf(IsNull(rsReporte.Fields!tiposaltaSEM), "", rsReporte.Fields!tiposaltaSEM)
            oRsFox.Fields!financia = rsReporte.Fields!CodigoFuenteFinanciamientoSEM
            oRsFox.Fields!coddiag1 = lcDx1
            oRsFox.Fields!coddiag2 = lcDx2
            oRsFox.Fields!coddiag3 = lcDx3
            oRsFox.Fields!coddiag4 = lcDx4
            oRsFox.Fields!cemorb1 = ""
            oRsFox.Fields!cemorb2 = ""
            oRsFox.Fields!codcpt1 = lcCpt1
            oRsFox.Fields!codcpt2 = lcCpt2
            oRsFox.Fields!codcpt3 = lcCpt3
            oRsFox.Fields!codcpt4 = lcCpt4
            oRsFox.Fields!estadio = ""
            oRsFox.Fields!valor_t = ""
            oRsFox.Fields!valor_n = ""
            oRsFox.Fields!valor_m = ""
            oRsFox.Fields!tratamien = ""
            oRsFox.Fields!prof_parto = lcProfParto
            If ldFechaNacimiento <> 0 Then
               oRsFox.Fields!FecParto = Format(ldFechaNacimiento, sighentidades.DevuelveFechaSoloFormato_DMY)
            Else
               oRsFox.Fields!FecParto = 0
            End If
            oRsFox.Fields!rnvivo = lnRNvivos
            oRsFox.Fields!rnmuerto = lnRNmuertos
            oRsFox.Fields!codpsal = lcCodSal
            oRsFox.Fields!FechaReg = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
            oRsFox.Fields!Estado = 1
            oRsFox.Fields!Usuario = lcUsuario9
            oRsFox.Fields!usuariof = ldUsuarioFecha9
            oRsFox.Update
        End If
        lnNumMov = lnNumMov + 1
        rsReporte.MoveNext
      Loop
      Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "A", 0, "Atenciones", oConexion, 500 + 140, mo_lcNombrePc, "Exportar al SEM_2016: " & lcFechaInicio & " al " & lcFechaFinal)       '500+ ListBarReporte.idReporte
    End If
    '*** problemas con DUPLICADOS
    oRsErrores.Filter = "error<>' '"
    If oRsErrores.RecordCount > 0 Then
       oRsErrores.MoveFirst
       Do While Not oRsErrores.EOF
            lcMensajes99 = lcMensajes99 & oRsErrores!Error & "  Historia: " & oRsErrores!historia & " , F.Ingreso:" & _
                              oRsErrores!fechaIng & " solucionar en FACTURACION->ESTADO CUENTA" & Chr(13)
            oRsErrores.MoveNext
       Loop
    End If
    If lcMensajes99 <> "" Then
       MsgBox "Errores de DATOS:" & Chr(13) & lcMensajes99, vbInformation, "Errores de registro de Información"
    End If
    
    oRsFox.Close
    oConexionFox.Close
    oConexion.Close
    Set oConexion = Nothing
    Set oConexionFox = Nothing
    Set oRsFox = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_AdminAdmision = Nothing
    Set mo_ReglasSeguridad = Nothing
    Set mo_ReglasAdmision = Nothing
    Set mo_ReglasServiciosHosp = Nothing
    Set mo_ReglasComunes = Nothing
    Set rsReporte = Nothing
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing
    Set rsServicio = Nothing
    Set oRsErrores = Nothing
    Set mo_ReglasHoteleria = Nothing
    Set oRsTmp2 = Nothing
    Exit Sub
ErrorProceso:
    MsgBox Err.Description & Chr(13) & lcParaMsgbox    '
    Exit Sub
    Resume
End Sub



'debb-15/06/2016
Sub ExportaSEM2017(lcFechaInicio As String, lcFechaFinal As String, lcCorrelativo As String, _
                    ml_Valor As Integer, ml_Tabla As String, lbMuestraUsuario As Boolean)
    On Error GoTo ErrorProceso
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_AdminAdmision As New ReglasAdmision
    Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_ReglasServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
    Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
    Dim mo_ReglasHoteleria As New SIGHNegocios.ReglasHoteleria
    Dim rsReporte As New Recordset
    Dim oRsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim rsOcupacion As New Recordset
    Dim oRsFox As New Recordset
    Dim rsServicio As New Recordset
    Dim oRsErrores As New Recordset
    Dim lbPrimerError As Boolean, lbNoExisteH As Boolean
    Dim oConexionFox As New Connection
    Dim oConexion As New Connection
    Dim lcDx1 As String, lcDx2 As String
    Dim lnNumMov As Long, lcSql As String, lnEstancia As Integer
    Dim lcFec1 As String, lcFec2 As String
    Dim lcMotivo As String, lcDestinoAtencion As String, lcReferido As String, lcUsuario As String
    Dim lcDireccionDomicilio As String, lcNombreAcompaniante As String
    Dim lnEsObservacionEmergencia As Integer, lcServicio As String, lcMedico As String, lcServicioObs As String
    Dim lcMedicoObs As String, ldFechaIngObs As Date, lcHoraIngObs As String, lcCamaObs As String, ldFechaSalObs As Date
    Dim lcHoraSalObs As String, lcDxObs As String, lcHoraEstanciaMax As String, lnTotal As Long
    Dim lcRenaes As String, lcUbigeoEESS As String, lcDisa As String, lcRed As String, lcMred As String
    Dim lcCodSal As String, lnRNmuertos As Long, lnRNvivos As Long, ldFechaNacimiento As Date, lcProfParto As String
    Dim lcCpt1 As String, lcCpt2 As String, lcCpt3 As String, lcCpt4 As String, lcDxTipo1 As String, lcDxTipo2 As String
    Dim lcDNIacompainante As String, lcNroEmergencia As String, lcUsuario9 As String, ldUsuarioFecha9 As Date
    Dim lcDx3 As String, lcDx4 As String, lcDxTipo3 As String, lcDxTipo4 As String, lcDiagn As String
    Dim lcPrioridad As String, lcForm_ing As String, lcMensajes99 As String, lcUPS11 As String
    Const lcEtniaDefault As String = "58"

    If ml_Valor = 2 Then
       FileCopy App.Path + "\archivos\sem_em1.dbf", App.Path + "\archivos\sem_emer.dbf"
    Else
       FileCopy App.Path + "\archivos\sem_eg1.dbf", App.Path + "\archivos\sem_egre.dbf"
    End If
    
    lcRenaes = Right("00000000" & lcBuscaParametro.SeleccionaFilaParametro(208), 10)
    lcUbigeoEESS = lcBuscaParametro.SeleccionaFilaParametro(242)
    lcHoraEstanciaMax = lcBuscaParametro.SeleccionaFilaParametro(201)
    lcDisa = lcBuscaParametro.SeleccionaFilaParametro(239)
    lcRed = lcBuscaParametro.SeleccionaFilaParametro(240)
    lcMred = lcBuscaParametro.SeleccionaFilaParametro(241)
        '
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
        '
    oConexionFox.CommandTimeout = 300
    lcParaMsgbox = " (Eliminar Datos)"
    
    lcMensajes99 = ""
    
    
    
    If ml_Valor = 2 Then
        '*********************************Emergencia
        
        oConexionFox.Open "DSN=SEM"
        ml_Tabla = "sem_emer"
        mo_ReglasComunes.AbreTablaFox oRsFox, oConexionFox, ml_Tabla, "", ""
        If oRsFox.RecordCount > 0 Then
          oRsFox.MoveFirst
          Do While Not oRsFox.EOF
            If oRsFox.Fields!fecate >= CDate(lcFechaInicio) And oRsFox.Fields!fecate <= CDate(lcFechaFinal) Then
              oRsFox.Delete
              oRsFox.Update
            End If
            oRsFox.MoveNext
          Loop
        End If
        oRsFox.Close
        '
        mo_ReglasComunes.AbreTablaFox oRsFox, oConexionFox, ml_Tabla, "", ""
    Else
        '*********************************Hospitalizacion
        oConexionFox.Open "DSN=SEM"
        ml_Tabla = "sem_egre"
        mo_ReglasComunes.AbreTablaFox oRsFox, oConexionFox, ml_Tabla, "", ""
        If oRsFox.RecordCount > 0 Then
          oRsFox.MoveFirst
          Do While Not oRsFox.EOF
            If oRsFox.Fields!fecIng >= CDate(lcFechaInicio) And oRsFox.Fields!fecIng <= CDate(lcFechaFinal) Then
              oRsFox.Delete
              oRsFox.Update
            End If
            oRsFox.MoveNext
          Loop
        End If
        oRsFox.Close
        '
        mo_ReglasComunes.AbreTablaFox oRsFox, oConexionFox, ml_Tabla, "", ""
    End If
 '   DoEvents
    '
    lcUsuario = " "
    Set oRsTmp1 = mo_ReglasComunes.EmpleadosSeleccionarPorIdEmpleado(ml_idUsuario)
    If oRsTmp1.RecordCount > 0 Then
       If Not IsNull(oRsTmp1.Fields!Usuario) Then
          lcUsuario = oRsTmp1.Fields!Usuario
       End If
    End If
    oRsTmp1.Close
    '
    Set rsReporte = mo_ReglasAdmision.AtencionesSeleccionarHospEmergPorFechasDeEgresoMedico(CDate(lcFechaInicio), CDate(lcFechaFinal), "FECHAEGRESO")
    rsReporte.Filter = "codigoCondicionAltaSEM<>null and idTipoServicio=" & ml_Valor
    lnTotal = rsReporte.RecordCount
    
    If lnTotal = 0 Then
      MsgBox "No existe Información  con esos Datos", vbInformation, "Mensaje"
    Else
      '
      lbPrimerError = True
      With oRsErrores
            .Fields.Append "Historia", adInteger
            .Fields.Append "FechaIng", adDate
            .Fields.Append "Error", adVarChar, 50, adFldIsNullable
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .Open
      End With
      '
      lnNumMov = Val(lcCorrelativo)
      rsReporte.MoveFirst
      Do While Not rsReporte.EOF
If rsReporte!idCuentaAtencion = 214945 Then
lcParaMsgbox = ""
End If
        '
'        If rsReporte!IdEtnia <> lcEtniaDefault Then
'           lcMensajes99 = lcMensajes99 & " HC=" & Trim(Str(rsReporte.Fields!NroHistoriaClinica)) & " tiene ETNIA diferente a: " & lcEtniaDefault & Chr(13)
'        End If
        If Len(Trim(rsReporte.Fields!CodigoFuenteFinanciamientoSEM)) <> 2 Then
           lcCpt1 = lcCpt1 & " N° Cuenta=" & Trim(str(rsReporte.Fields!idCuentaAtencion)) & " tiene FUENTE FINANCIAMIENTO diferente a 2 dígitos, vefiricar FUENTE FINANCIAMIENTO en CUENTA y luego en 'Fact-Config->Fuente Financiamiento->código sistema SEM'" & Chr(13)
        End If
        '
        If lbPrimerError = True Then
           lbPrimerError = False
           oRsErrores.AddNew
           oRsErrores.Fields!historia = rsReporte!NroHistoriaClinica
           oRsErrores.Fields!fechaIng = rsReporte!FechaIngreso
           oRsErrores.Fields!Error = Space(1)
           oRsErrores.Update
        Else
           lbNoExisteH = True
           oRsErrores.MoveFirst
           Do While Not oRsErrores.EOF
              If oRsErrores!historia = rsReporte!NroHistoriaClinica And oRsErrores!fechaIng = rsReporte!FechaIngreso Then
                 lbNoExisteH = False
                 Exit Do
              End If
              oRsErrores.MoveNext
           Loop
           If lbNoExisteH = True Then
                oRsErrores.AddNew
                oRsErrores.Fields!historia = rsReporte!NroHistoriaClinica
                oRsErrores.Fields!fechaIng = rsReporte!FechaIngreso
                oRsErrores.Fields!Error = Space(1)
           Else
            '    oRsErrores.Fields!Error = "(misma F.Ingreso)"
           End If
           oRsErrores.Update
        End If
        '
        If IsNull(rsReporte.Fields!codigoServicioSem) Then
           Set rsOcupacion = mo_AdminAdmision.EstanciaHospitalariaSeleccionarPorAtencion(rsReporte!idAtencion, 0, oConexion)
           lcUPS11 = ".."
           If rsOcupacion.RecordCount > 0 Then
              rsOcupacion.MoveFirst
              rsOcupacion.Find "idServicio=" & rsReporte!IdServicioEgreso
              If Not rsOcupacion.EOF Then
                 If rsOcupacion!EsObservacionEmergencia = True Then
                    rsOcupacion.MoveFirst
                    Do While Not rsOcupacion.EOF
                       If Not IsNull(rsOcupacion!codigoServicioSem) Then
                          lcUPS11 = rsOcupacion!codigoServicioSem
                       End If
                       rsOcupacion.MoveNext
                    Loop
                 End If
              End If
           End If
           rsOcupacion.Close
        Else
           lcUPS11 = rsReporte.Fields!codigoServicioSem
        End If
        
        '
        lcParaMsgbox = " -->Cuenta: " & rsReporte.Fields!idCuentaAtencion
        'solo Hospitalizados
        lcDiagn = "/"
        lcDx1 = "": lcDx2 = "": lcDxTipo1 = "": lcDxTipo2 = "": lcDx3 = "": lcDx4 = "": lcDxTipo3 = "": lcDxTipo4 = ""
        Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(rsReporte.Fields!idAtencion, 3, oConexion)
        oRsTmp.Filter = "IdTipoDiagnostico=301"   'vivos-princial
        If oRsTmp.RecordCount > 0 Then
          '*****vivos
          oRsTmp.MoveFirst
          Do While Not oRsTmp.EOF
            If lcDx1 = "" Then
                lcDx1 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                lcDxTipo1 = "D"
                lcDiagn = lcDiagn & lcDx1 & "/"
            ElseIf lcDx2 = "" Then
                lcDx2 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                lcDxTipo2 = "D"
                lcDiagn = lcDiagn & lcDx2 & "/"
            ElseIf lcDx3 = "" Then
                lcDx3 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                lcDxTipo3 = "D"
                lcDiagn = lcDiagn & lcDx3 & "/"
            End If
            oRsTmp.MoveNext
          Loop
'          oRsTmp.Filter = "IdTipoDiagnostico=302"
'          If oRsTmp.RecordCount > 0 Then
'            oRsTmp.MoveFirst
'            lcDx2 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
'            lcDxTipo2 = "P"
'            lcDiagn = lcDiagn & lcDx2 & "/"
'          End If
          oRsTmp.Filter = "IdTipoDiagnostico=302"
          If oRsTmp.RecordCount > 0 Then
            oRsTmp.MoveFirst
            Do While Not oRsTmp.EOF
                If lcDx2 = "" Then
                    lcDx2 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                    lcDxTipo2 = "P"
                    lcDiagn = lcDiagn & lcDx2 & "/"
                ElseIf lcDx3 = "" Then
                    lcDx3 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                    lcDxTipo3 = "P"
                    lcDiagn = lcDiagn & lcDx3 & "/"
                ElseIf lcDx4 = "" Then
                    lcDx4 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                    lcDxTipo4 = "P"
                    lcDiagn = lcDiagn & lcDx4 & "/"
                End If
                oRsTmp.MoveNext
            Loop
          End If
          oRsTmp.Close
          Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(rsReporte.Fields!idAtencion, 2, oConexion)
          oRsTmp.Filter = "idTipoDiagnostico<>null"
          If oRsTmp.RecordCount > 0 Then
            oRsTmp.MoveFirst
            Do While Not oRsTmp.EOF
                If lcDx2 = "" Then
                    lcDx2 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                    lcDxTipo2 = "P"
                    lcDiagn = lcDiagn & lcDx2 & "/"
                ElseIf lcDx3 = "" Then
                    lcDx3 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                    lcDxTipo3 = "P"
                    lcDiagn = lcDiagn & lcDx3 & "/"
                ElseIf lcDx4 = "" Then
                    lcDx4 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                    lcDxTipo4 = "P"
                    lcDiagn = lcDiagn & lcDx4 & "/"
                End If
                oRsTmp.MoveNext
            Loop
          End If
        Else
          '******fallecidos
          Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(rsReporte.Fields!idAtencion, 4, oConexion)
          If oRsTmp.RecordCount > 0 Then
                oRsTmp.Filter = "IdTipoDiagnostico=303"    'Final
                If oRsTmp.RecordCount > 0 Then
                  oRsTmp.MoveFirst
                  lcDx1 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                  lcDxTipo1 = "D"
                  lcDiagn = lcDiagn & lcDx1 & "/"
                End If
                oRsTmp.Filter = "IdTipoDiagnostico=304 or IdTipoDiagnostico=305"    'fallecidos - Final
                If oRsTmp.RecordCount > 0 Then
                  oRsTmp.MoveFirst
                  lcDx2 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                  lcDxTipo2 = "P"
                  lcDiagn = lcDiagn & lcDx2 & "/"
                End If
          End If
        End If
        oRsTmp.Filter = "idTipoDiagnostico<>null"   'oRsTmp.Filter = ""
        If oRsTmp.RecordCount > 0 Then
           oRsTmp.MoveFirst
           Do While Not oRsTmp.EOF
              If InStr(lcDiagn, Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)) = 0 Then
                 If lcDx2 = "" Then
                    lcDx2 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                    lcDxTipo2 = IIf(oRsTmp!idTipoDiagnostico = 301 Or oRsTmp!idTipoDiagnostico = 303, "D", "P")
                    lcDiagn = lcDiagn & lcDx2 & "/"
                 ElseIf lcDx3 = "" Then
                    lcDx3 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                    lcDxTipo3 = IIf(oRsTmp!idTipoDiagnostico = 301 Or oRsTmp!idTipoDiagnostico = 303, "D", "P")
                    lcDiagn = lcDiagn & lcDx3 & "/"
                 ElseIf lcDx4 = "" Then
                    lcDx4 = Left(sighentidades.DevuelveCodigoDxSinPUNTO(oRsTmp.Fields!CodigoCie2004), 4)
                    lcDxTipo4 = IIf(oRsTmp!idTipoDiagnostico = 301 Or oRsTmp!idTipoDiagnostico = 303, "D", "P")
                    Exit Do
                 End If
              End If
              oRsTmp.MoveNext
           Loop
        End If
        oRsTmp.Close
        'nacimientos - solo hospitalizacion
        lnRNvivos = 0: lnRNmuertos = 0: ldFechaNacimiento = 0: lcProfParto = ""
        If ml_Valor <> 2 Then
           Set oRsTmp1 = mo_AdminAdmision.AtencionesNacimientosSeleccionarPorAtencion(rsReporte!idAtencion, oConexion)
           If oRsTmp1.RecordCount > 0 Then
              lcProfParto = "1"
              If rsReporte!idTipoEmpleado = 28 Or rsReporte!idTipoEmpleado = 109 Then   '****obstretra/medico-gineco-obstetra****
                 lcProfParto = "2"
              End If
              
              oRsTmp1.MoveFirst
              ldFechaNacimiento = oRsTmp1!FechaNacimiento
              Do While Not oRsTmp1.EOF
                 If oRsTmp1!idCondicionRN = 1 Then
                    lnRNvivos = lnRNvivos + 1
                 Else
                    lnRNmuertos = lnRNmuertos + 1
                 End If
                 oRsTmp1.MoveNext
              Loop
           End If
           oRsTmp1.Close
        End If
        'cpt
        lcCpt1 = "": lcCpt2 = "": lcCpt3 = "": lcCpt4 = ""
        Set oRsTmp1 = mo_ReglasAdmision.BuscaOtrosCPTyExamenesparaFormatoHIS(rsReporte!idCuentaAtencion, _
                                        sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion, oConexion)
        If oRsTmp1.RecordCount > 0 Then
             oRsTmp1.MoveFirst
             Do While Not oRsTmp1.EOF
               If Len(Trim(oRsTmp1.Fields("codigo").Value)) = 5 Then
                    If lcCpt1 = "" Then
                       lcCpt1 = Trim(oRsTmp1.Fields("codigo").Value)
                    ElseIf lcCpt2 = "" Then
                       lcCpt2 = Trim(oRsTmp1.Fields("codigo").Value)
                    ElseIf lcCpt3 = "" Then
                       lcCpt3 = Trim(oRsTmp1.Fields("codigo").Value)
                    Else
                       lcCpt4 = Trim(oRsTmp1.Fields("codigo").Value)
                    End If
               End If
               oRsTmp1.MoveNext
             Loop
        End If
        oRsTmp1.Close
        '
        lcPrioridad = " "
        If Not IsNull(rsReporte!tipoGravedadSEM) Then
           lcPrioridad = rsReporte!tipoGravedadSEM
        End If
        '
        lcForm_ing = " "
        lcMotivo = " "
        Set oRsTmp1 = mo_AdminAdmision.AtencionesEmergenciaXidAtencion(rsReporte.Fields!idAtencion, oConexion)
        If oRsTmp1.RecordCount > 0 Then
           If Not IsNull(oRsTmp1.Fields!motivoSEM) Then
              lcMotivo = oRsTmp1.Fields!motivoSEM
           End If
           If Not IsNull(oRsTmp1!comoLlego) Then
              lcForm_ing = Trim(str(oRsTmp1!comoLlego))
           End If
        End If
        oRsTmp1.Close
        '
        lcDireccionDomicilio = "": lcNombreAcompaniante = "": lcDNIacompainante = "0": lcNroEmergencia = ""
        Set oRsTmp1 = mo_AdminAdmision.atencionesDatosAdicionalesXfiltro("dbo.AtencionesDatosAdicionales.idAtencion=" & rsReporte.Fields!idAtencion, oConexion)
        If oRsTmp1.RecordCount > 0 Then
           If Not IsNull(oRsTmp1.Fields!DireccionDomicilio) Then
              lcDireccionDomicilio = oRsTmp1.Fields!DireccionDomicilio
           End If
           If Not IsNull(oRsTmp1.Fields!NombreAcompaniante) Then
              lcNombreAcompaniante = oRsTmp1.Fields!NombreAcompaniante
           End If
           If Not IsNull(oRsTmp1.Fields!AcompanianteDNI) Then
              lcDNIacompainante = oRsTmp1.Fields!AcompanianteDNI
           End If
           If Not IsNull(oRsTmp1.Fields!emergenciaCorrelativo) Then
              lcNroEmergencia = Mid(oRsTmp1.Fields!emergenciaCorrelativo, 5, 20)
           End If
        End If
        oRsTmp1.Close
        '
        lcDestinoAtencion = " "
        If Not IsNull(rsReporte.Fields!IdDestinoAtencion) Then
            Set oRsTmp1 = mo_ReglasAdmision.TiposDestinoAtencionXidDestinoAtencion(rsReporte.Fields!IdDestinoAtencion, oConexion)
            If oRsTmp1.RecordCount > 0 Then
               If IsNull(oRsTmp1.Fields!destinoSEM) Then
                    MsgBox "Existen Servicios de Emergencia sin CODIGO SEM", vbInformation, "Mensaje"  'debb 05/12/2012
                    Exit Sub
               Else
                    lcDestinoAtencion = oRsTmp1.Fields!destinoSEM
               End If
            End If
            oRsTmp1.Close
        End If
        '
        lcReferido = " "
        If Not IsNull(rsReporte.Fields!idEstablecimientoDestino) Then
            Set oRsTmp1 = mo_ReglasComunes.EstablecimientosSeleccionarPorFiltro("idEstablecimiento=" & rsReporte.Fields!idEstablecimientoDestino, oConexion)
            If oRsTmp1.RecordCount > 0 Then
               lcReferido = Right("0000000000" & Trim(str(Val(oRsTmp1.Fields!codigo))), 10)
            End If
            oRsTmp1.Close
        End If
        '
        lcCodSal = "."
        If Not (IsNull(rsReporte!idTipoDocumento) Or IsNull(rsReporte!dni) Or IsNull(rsReporte!idColegioHIS)) Then
           lcCodSal = Trim(str(rsReporte!idTipoDocumento)) & Trim(rsReporte!dni) & rsReporte!idColegioHIS
        End If
        '
        If Val(lcRenaes) = 6918 Then      'hosp tarapoto
           lnEstancia = lcBuscaParametro.DiasDelPacienteEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, rsReporte!FechaEgreso, rsReporte!horaEgreso, lcHoraEstanciaMax, True)
        Else
           lnEstancia = lcBuscaParametro.DiasDelPacienteEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, rsReporte!FechaEgreso, rsReporte!horaEgreso, lcHoraEstanciaMax)
        End If
        If lnEstancia = 0 Then lnEstancia = 1
        '
        lcUsuario9 = " ": ldUsuarioFecha9 = 0
        If lbMuestraUsuario = True Then
            Set rsServicio = mo_ReglasSeguridad.AuditoriaFiltrarCitasPorIdAtencion(rsReporte!idAtencion, _
                                   IIf(ml_Valor = 2, sghAdmisionEmergencia, sghAdmisionHospitalizacion), _
                                   oConexion)
            If rsServicio.RecordCount > 0 Then
               lcUsuario9 = IIf(IsNull(rsServicio!dUsuario), "", rsServicio!dUsuario)
               ldUsuarioFecha9 = IIf(IsNull(rsServicio!fechaHora), 0, Format(rsServicio!fechaHora, sighentidades.DevuelveFechaSoloFormato_DMY))
            End If
            rsServicio.Close
        End If
        '
        If ml_Valor = 2 Then
            '*********************************Emergencia
            If lcUPS11 = ".." Then
                MsgBox "Para la cuenta. " & rsReporte.Fields!idCuentaAtencion & " no se registró CODIGO SERVICIO SEM (buscar la CUENTA ->ALTA MEDICA ->SERVICIO EGRESO, luego ir a GENERAL->SERVICIOS", vbInformation, "Mensaje"
                Exit Sub
            End If
            lcServicio = IIf(IsNull(rsReporte.Fields!codigoServicioSem), ".", rsReporte.Fields!codigoServicioSem)
            If IsNull(rsReporte.Fields!Colegiatura) Then
                MsgBox "Para la cuenta. " & rsReporte.Fields!idCuentaAtencion & " no se registró COLEGIATURA DEL MEDICO (buscar la CUENTA ->ALTA MEDICA ->MEDICO EGRESO, luego ir a PROGRAMACION GENERAL->PROFESIONALES DE SALUD", vbInformation, "Mensaje"
                Exit Sub
            End If
            lcMedico = IIf(IsNull(rsReporte.Fields!Colegiatura), ".", Left(rsReporte.Fields!Colegiatura, 5))
            lnEsObservacionEmergencia = 0
            If rsReporte.Fields!EsObservacionEmergencia = True Then
                lnEsObservacionEmergencia = 1
                Set oRsTmp1 = mo_ReglasAdmision.AtencionesEstanciaHospitalariaXidAtencionIdServicioIngreso(rsReporte.Fields!IdServicioIngreso, rsReporte.Fields!idAtencion, oConexion)
                If oRsTmp1.RecordCount > 0 Then
                   If oRsTmp1.Fields!EsObservacionEmergencia = True Then
                      '***El Servicio de Ingreso y el Servicio de Alta son OBSERVACIONES DE EMERGENCIA
                      '***Solo se muestra datos del Servicio de Alta ya no es necesario mostrar los datos
                      '***de OBSERVACIONES
                      lnEsObservacionEmergencia = 0
                   Else
                      '***El servicio de Ingreso es un CONSULTORIO y el servicio de alta es OBSERVACION
                      '***Hubo transferencia desde CONSULTORIO hacia OBSERVACION
                      '***Muestra datos del CONSULTORIO y datos DE OBSERVACION
                      lcServicioObs = lcServicio
                      
                      If IsNull(oRsTmp1!lcColegiaturaMedico) Then
                         MsgBox "Para la cuenta. " & rsReporte.Fields!idCuentaAtencion & " no se registró COLEGIATURA DEL MEDICO (buscar la CUENTA ->MODIFICAR -> MEDICO INGRESO, luego ir a PROGRAMACION GENERAL->PROFESIONALES DE SALUD", vbInformation, "Mensaje"
                         Exit Sub
                      End If
                      lcMedicoObs = oRsTmp1!lcColegiaturaMedico
                      
                      lcDxObs = lcDx1
                      ldFechaIngObs = oRsTmp1.Fields!fechaDesocupacion
                      lcHoraIngObs = oRsTmp1.Fields!horaDesocupacion
                      'debb 05/12/2012
                      lcCamaObs = " "
                      If IsNull(rsReporte!IdCamaEgreso) Then  'debb 05/12/2012
                         MsgBox "Para la cuenta. " & rsReporte.Fields!idCuentaAtencion & " no se registró cama, siendo un SERVICIO DE OBSERVACION", vbInformation, "Mensaje"
                         Exit Sub
                      End If
                      Set oRsTmp2 = mo_ReglasHoteleria.CamasSeleccionarPorIdentificador(rsReporte!IdCamaEgreso, oConexion)
                      If oRsTmp2.RecordCount > 0 Then
                         lcCamaObs = Trim(oRsTmp2!codigo)
                      End If
                      oRsTmp2.Close
                      'debb 05/12/2012
                      ldFechaSalObs = rsReporte.Fields!FechaEgreso
                      lcHoraSalObs = rsReporte.Fields!horaEgreso
                      
                      If IsNull(oRsTmp1.Fields!codigoServicioSem) Then
                          MsgBox "Para la cuenta. " & rsReporte.Fields!idCuentaAtencion & " no se registró CODIGO SERVICIO SEM (buscar la CUENTA ->MODIFICAR->SERVICIO INGRESO, luego ir a GENERAL->SERVICIOS", vbInformation, "Mensaje"
                          Exit Sub
                      End If
                      lcServicio = IIf(IsNull(oRsTmp1.Fields!codigoServicioSem), ".", oRsTmp1.Fields!codigoServicioSem)
                      
                      
                      If IsNull(oRsTmp1.Fields!Colegiatura) Then
                         MsgBox "Para la cuenta. " & rsReporte.Fields!idCuentaAtencion & " no se registró COLEGIATURA DEL MEDICO (buscar la CUENTA ->MODIFICAR ->MEDICO INGRESO, luego ir a PROGRAMACION GENERAL->PROFESIONALES DE SALUD", vbInformation, "Mensaje"
                         Exit Sub
                      End If
                      lcMedico = IIf(IsNull(oRsTmp1.Fields!Colegiatura), ".", Left(oRsTmp1.Fields!Colegiatura, 5))
                      
                      lnEstancia = DateDiff("h", Format(ldFechaIngObs, "dd/mm/yyyy") & " " & lcHoraIngObs & ":00", _
                                                 Format(ldFechaSalObs, "dd/mm/yyyy") & " " & lcHoraSalObs & ":00")
                   End If
                End If
                oRsTmp1.Close
            Else
                '***El servicio de ALTA MEDICA es un CONSULTORIO
                Set oRsTmp1 = mo_ReglasAdmision.AtencionesEstanciaHospitalariaXidAtencionIdServicioIngreso(0, rsReporte.Fields!idAtencion, oConexion)
                oRsTmp1.Filter = "esObservacionEmergencia=1"
                If oRsTmp1.RecordCount > 0 Then
                      lnEsObservacionEmergencia = 1
                      '***El servicio de ALTA MEDICA es un CONSULTORIO
                      '***Hubo al menos una TRANSFERENCIA hacia OBSERVACION (o se ingresó de frente a OBSERVACION)
                      '***Muestra datos DE OBSERVACION
                      
                      lcServicioObs = lcServicio
                      
                      If IsNull(oRsTmp1!lcColegiaturaMedico) Then
                         MsgBox "Para la cuenta. " & rsReporte.Fields!idCuentaAtencion & " no se registró COLEGIATURA DEL MEDICO (buscar la CUENTA ->MODIFICAR -> TRANSFERENCIA, luego ir a PROGRAMACION GENERAL->PROFESIONALES DE SALUD", vbInformation, "Mensaje"
                         Exit Sub
                      End If
                      lcMedicoObs = oRsTmp1!lcColegiaturaMedico
                      
                      lcDxObs = lcDx1
                      ldFechaIngObs = oRsTmp1!fechaOcupacion
                      lcHoraIngObs = oRsTmp1!horaOcupacion
                      'debb 05/12/2012
                      If IsNull(oRsTmp1.Fields!cama) Then  'debb 05/12/2012
                         MsgBox "Para la cuenta. " & rsReporte.Fields!idCuentaAtencion & " no se registró cama, siendo un SERVICIO DE OBSERVACION", vbInformation, "Mensaje"
                         Exit Sub
                      Else
                         lcCamaObs = oRsTmp1.Fields!cama
                      End If
                      'debb 05/12/2012
                      ldFechaSalObs = oRsTmp1!fechaDesocupacion
                      lcHoraSalObs = oRsTmp1!horaDesocupacion
                      'lcServicio = IIf(IsNull(oRsTmp1.Fields!codigoServicioSEM), ".", oRsTmp1.Fields!codigoServicioSEM)
                      'lcMedico = IIf(IsNull(oRsTmp1.Fields!Colegiatura), ".", Left(oRsTmp1.Fields!Colegiatura, 5))
                      lnEstancia = DateDiff("h", Format(ldFechaIngObs, "dd/mm/yyyy") & " " & lcHoraIngObs, _
                                                 Format(ldFechaSalObs, "dd/mm/yyyy") & " " & lcHoraSalObs)
                End If
                oRsTmp1.Close
            End If
            
            oRsFox.AddNew
            oRsFox.Fields!renIpress = lcRenaes
            oRsFox.Fields!e_ubig = lcUbigeoEESS
            oRsFox.Fields!cod_disa = lcDisa
            oRsFox.Fields!cod_red = lcRed
            oRsFox.Fields!cod_mred = lcMred
            oRsFox.Fields!fecate = Format(rsReporte.Fields!FechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY)
            oRsFox.Fields!horate = rsReporte!HoraIngreso
            If Val(lcRenaes) = 6918 Then
                oRsFox.Fields!numHc = lcNroEmergencia    'solo HOspital Tarapoto
            Else
                If Not IsNull(rsReporte.Fields!NroHistoriaClinica) Then
                   oRsFox.Fields!numHc = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(str(rsReporte.Fields!NroHistoriaClinica)), False)
                Else
                   oRsFox.Fields!numHc = "sin HC"
                End If
            End If
            If IsNull(rsReporte.Fields!IdDocIdentidad) Or IsNull(rsReporte.Fields!nroDocumento) Then
               oRsFox.Fields!doc_iden = "0"
            ElseIf rsReporte.Fields!IdDocIdentidad >= 1 And rsReporte.Fields!IdDocIdentidad <= 4 Then
               oRsFox.Fields!doc_iden = Trim(str(rsReporte.Fields!IdDocIdentidad)) & Left(rsReporte.Fields!nroDocumento, 10)
            Else
               oRsFox.Fields!doc_iden = "0"
            End If
            oRsFox.Fields!etnia = IIf(IsNull(rsReporte!IdEtnia), "80", rsReporte!IdEtnia)
            oRsFox.Fields!financia = rsReporte.Fields!CodigoFuenteFinanciamientoSEM
            oRsFox.Fields!Sexo = Trim(str(rsReporte.Fields!idTipoSexo))
            oRsFox.Fields!Edad = rsReporte.Fields!Edad
            oRsFox.Fields!TipoEdad = rsReporte.Fields!tipoEdadSEM
            oRsFox.Fields!nomb = Left(Trim(rsReporte.Fields!PrimerNombre) & " " & IIf(IsNull(rsReporte.Fields!SegundoNombre), "", Trim(rsReporte.Fields!SegundoNombre)), 50)
            oRsFox.Fields!apell = Left(Trim(rsReporte.Fields!ApellidoPaterno) & " " & Trim(rsReporte.Fields!ApellidoMaterno), 50)
            oRsFox.Fields!direcc = Left(lcDireccionDomicilio, 100)
            If Not IsNull(rsReporte!IdDistritoDomicilio) Then
               oRsFox.Fields!ubig_resha = Right("0" & Trim(str(rsReporte!IdDistritoDomicilio)), 6)
               oRsFox.Fields!sitocurren = Right("0" & Trim(str(rsReporte!IdDistritoDomicilio)), 6)
            Else
               oRsFox.Fields!ubig_resha = "000000"
               oRsFox.Fields!sitocurren = "000000"
            End If
            '
            If Not IsNull(rsReporte.Fields!IdDistritoProcedencia) Then
               oRsFox.Fields!ubig_proce = Right("0" & Trim(str(rsReporte.Fields!IdDistritoProcedencia)), 6)
            Else
               If Not IsNull(rsReporte!IdDistritoDomicilio) Then
                  oRsFox.Fields!ubig_proce = Right("0" & Trim(str(rsReporte!IdDistritoDomicilio)), 6)
               Else
                  oRsFox.Fields!ubig_proce = "000000"
               End If
            End If
            '
            oRsFox.Fields!acompana = Left(lcNombreAcompaniante, 50)
            oRsFox.Fields!adoc_iden = lcDNIacompainante
            oRsFox.Fields!motaten = lcMotivo
            
            
            oRsFox.Fields!UPS = lcUPS11
            oRsFox.Fields!coddiag1 = lcDx1
            oRsFox.Fields!tipdiag1 = lcDxTipo1
            oRsFox.Fields!coddiag2 = lcDx2
            oRsFox.Fields!tipdiag2 = lcDxTipo2
            oRsFox.Fields!coddiag3 = lcDx3
            oRsFox.Fields!tipdiag3 = lcDxTipo3
            oRsFox.Fields!coddiag4 = lcDx4
            oRsFox.Fields!tipdiag4 = lcDxTipo4
            oRsFox.Fields!codcpt1 = lcCpt1
            oRsFox.Fields!codcpt2 = lcCpt2
            oRsFox.Fields!codcpt3 = lcCpt3
            oRsFox.Fields!codcpt4 = lcCpt4
            oRsFox.Fields!condicion = IIf(IsNull(rsReporte.Fields!tiposaltaSEM), "", rsReporte.Fields!tiposaltaSEM)
            oRsFox.Fields!fecegr = Format(rsReporte.Fields!FechaEgreso, sighentidades.DevuelveFechaSoloFormato_DMY)
            oRsFox.Fields!horegr = rsReporte.Fields!horaEgreso
            oRsFox.Fields!destino = lcDestinoAtencion
            oRsFox.Fields!des_eess = lcReferido
            oRsFox.Fields!des_ups = ""
            oRsFox.Fields!codpsal = lcCodSal
            oRsFox.Fields!observ = lnEsObservacionEmergencia
            If lnEsObservacionEmergencia = 1 Then
                oRsFox.Fields!ofecing = ldFechaIngObs
                oRsFox.Fields!ohoring = lcHoraIngObs
                oRsFox.Fields!ofecegr = ldFechaSalObs
                oRsFox.Fields!ohoregr = lcHoraSalObs
                oRsFox.Fields!totalest = lnEstancia
                oRsFox.Fields!ocama = lcCamaObs
                oRsFox.Fields!ocodpsal = lcMedicoObs
                oRsFox.Fields!ocoddiag1 = lcDxObs
                oRsFox.Fields!ocoddiag2 = ""
            Else
                oRsFox.Fields!ofecing = 0
                oRsFox.Fields!ohoring = ""
                oRsFox.Fields!ofecegr = 0
                oRsFox.Fields!ohoregr = ""
                oRsFox.Fields!totalest = lnEstancia
                oRsFox.Fields!ocama = ""
                oRsFox.Fields!ocodpsal = ""
                oRsFox.Fields!ocoddiag1 = ""
                oRsFox.Fields!ocoddiag2 = ""
            End If
            oRsFox.Fields!FechaReg = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
            oRsFox.Fields!Estado = "1"
            oRsFox.Fields!Usuario = lcUsuario9
            oRsFox.Fields!usuariof = ldUsuarioFecha9
            oRsFox.Fields!prioridad = lcPrioridad
            oRsFox.Fields!form_ing = lcForm_ing
            
            oRsFox.Update
            
        Else
            '*********************************Hospitalizacion
            oRsFox.AddNew
            oRsFox.Fields!renIpress = lcRenaes
            oRsFox.Fields!e_ubig = lcUbigeoEESS
            oRsFox.Fields!e_cdpto = Left(lcUbigeoEESS, 2)
            oRsFox.Fields!e_cprov = Mid(lcUbigeoEESS, 3, 2)
            oRsFox.Fields!e_cdist = Right(lcUbigeoEESS, 2)
            oRsFox.Fields!cod_disa = lcDisa
            oRsFox.Fields!cod_red = lcRed
            oRsFox.Fields!cod_mred = lcMred
            If Not IsNull(rsReporte.Fields!NroHistoriaClinica) Then
               If Val(lcRenaes) = 6918 Then   'tarapoto
                  oRsFox.Fields!numHc = Right("000000" & Trim(str(rsReporte.Fields!NroHistoriaClinica)), 6)
               Else
                  oRsFox.Fields!numHc = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(str(rsReporte.Fields!NroHistoriaClinica)), False)
               End If
            Else
               oRsFox.Fields!numHc = "sin HC"
            End If
            oRsFox.Fields!nomb = Left(Trim(rsReporte.Fields!PrimerNombre) & " " & IIf(IsNull(rsReporte.Fields!SegundoNombre), "", Trim(rsReporte.Fields!SegundoNombre)), 50)
            oRsFox.Fields!apell = Left(Trim(rsReporte.Fields!ApellidoPaterno) & " " & Trim(rsReporte.Fields!ApellidoMaterno), 50)
            If IsNull(rsReporte.Fields!IdDocIdentidad) Or IsNull(rsReporte.Fields!nroDocumento) Then
               oRsFox.Fields!doc_iden = "0"
            ElseIf rsReporte.Fields!IdDocIdentidad >= 1 And rsReporte.Fields!IdDocIdentidad <= 4 Then
               oRsFox.Fields!doc_iden = Trim(str(rsReporte.Fields!IdDocIdentidad)) & Left(rsReporte.Fields!nroDocumento, 10)
            Else
               oRsFox.Fields!doc_iden = "0"
            End If
            
            
            oRsFox.Fields!etnia = IIf(IsNull(rsReporte!IdEtnia), "80", rsReporte!IdEtnia)
            oRsFox.Fields!Sexo = Trim(str(rsReporte.Fields!idTipoSexo))
            oRsFox.Fields!Edad = rsReporte.Fields!Edad
            oRsFox.Fields!TipoEdad = rsReporte.Fields!tipoEdadSEM
            If IsNull(rsReporte!IdDistritoDomicilio) Then
               oRsFox.Fields!Ubigeo = "000000"
               oRsFox.Fields!cdpto = "00"
               oRsFox.Fields!cprov = "00"
               oRsFox.Fields!cdist = "00"
            Else
               oRsFox.Fields!Ubigeo = Right("0" & Trim(str(rsReporte!IdDistritoDomicilio)), 6)
               oRsFox.Fields!cdpto = Left(Right("0" & Trim(str(rsReporte!IdDistritoDomicilio)), 6), 2)
               oRsFox.Fields!cprov = Mid(Right("0" & Trim(str(rsReporte!IdDistritoDomicilio)), 6), 3, 2)
               oRsFox.Fields!cdist = Right("0" & Trim(str(rsReporte!IdDistritoDomicilio)), 2)
            End If
            oRsFox.Fields!fecIng = Format(rsReporte!FechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY)
            oRsFox.Fields!fecegr = Format(rsReporte.Fields!FechaEgreso, sighentidades.DevuelveFechaSoloFormato_DMY)
            oRsFox.Fields!totalest = lnEstancia
            
            
            oRsFox.Fields!UPS = IIf(IsNull(rsReporte.Fields!codigoServicioSem), "000000", rsReporte.Fields!codigoServicioSem)
            oRsFox.Fields!condicion = IIf(IsNull(rsReporte.Fields!tiposaltaSEM), "", rsReporte.Fields!tiposaltaSEM)
            oRsFox.Fields!financia = rsReporte.Fields!CodigoFuenteFinanciamientoSEM
            oRsFox.Fields!coddiag1 = lcDx1
            oRsFox.Fields!coddiag2 = lcDx2
            oRsFox.Fields!coddiag3 = lcDx3
            oRsFox.Fields!coddiag4 = lcDx4
            oRsFox.Fields!cemorb1 = ""
            oRsFox.Fields!cemorb2 = ""
            oRsFox.Fields!codcpt1 = lcCpt1
            oRsFox.Fields!codcpt2 = lcCpt2
            oRsFox.Fields!codcpt3 = lcCpt3
            oRsFox.Fields!codcpt4 = lcCpt4
            oRsFox.Fields!estadio = ""
            oRsFox.Fields!valor_t = ""
            oRsFox.Fields!valor_n = ""
            oRsFox.Fields!valor_m = ""
            oRsFox.Fields!tratamien = ""
            oRsFox.Fields!prof_parto = lcProfParto
            If ldFechaNacimiento <> 0 Then
               oRsFox.Fields!FecParto = Format(ldFechaNacimiento, sighentidades.DevuelveFechaSoloFormato_DMY)
            Else
               oRsFox.Fields!FecParto = 0
            End If
            oRsFox.Fields!rnvivo = lnRNvivos
            oRsFox.Fields!rnmuerto = lnRNmuertos
            oRsFox.Fields!codpsal = lcCodSal
            oRsFox.Fields!FechaReg = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
            oRsFox.Fields!Estado = 1
            oRsFox.Fields!Usuario = lcUsuario9
            oRsFox.Fields!usuariof = ldUsuarioFecha9
            oRsFox.Update
        End If
        lnNumMov = lnNumMov + 1
        rsReporte.MoveNext
      Loop
      Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "A", 0, "Atenciones", oConexion, 500 + 140, mo_lcNombrePc, "Exportar al SEM_2016: " & lcFechaInicio & " al " & lcFechaFinal)       '500+ ListBarReporte.idReporte
    End If
    '*** problemas con DUPLICADOS
    oRsErrores.Filter = "error<>' '"
    If oRsErrores.RecordCount > 0 Then
       oRsErrores.MoveFirst
       Do While Not oRsErrores.EOF
            lcMensajes99 = lcMensajes99 & oRsErrores!Error & "  Historia: " & oRsErrores!historia & " , F.Ingreso:" & _
                              oRsErrores!fechaIng & " solucionar en FACTURACION->ESTADO CUENTA" & Chr(13)
            oRsErrores.MoveNext
       Loop
    End If
    If lcMensajes99 <> "" Then
       MsgBox "Errores de DATOS:" & Chr(13) & lcMensajes99, vbInformation, "Errores de registro de Información"
    End If
    
    oRsFox.Close
    oConexionFox.Close
    oConexion.Close
    Set oConexion = Nothing
    Set oConexionFox = Nothing
    Set oRsFox = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_AdminAdmision = Nothing
    Set mo_ReglasSeguridad = Nothing
    Set mo_ReglasAdmision = Nothing
    Set mo_ReglasServiciosHosp = Nothing
    Set mo_ReglasComunes = Nothing
    Set mo_ReglasHoteleria = Nothing
    Set rsReporte = Nothing
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set rsServicio = Nothing
    Set oRsErrores = Nothing
    Set rsOcupacion = Nothing
    Exit Sub
ErrorProceso:
    MsgBox Err.Description & Chr(13) & lcParaMsgbox    '
    Exit Sub
    Resume
End Sub








Sub ReportePartidaREsumen(ByRef mrs_Tmp As Recordset, mda_FechaInicio As Date, mda_FechaFin As Date, _
                          ml_IdPartidaFiltro As Long, lnTipoProceso As sghTipoProceso, _
                          moConexion As Connection, lbSeCreaTemporal As Boolean, lnIdCajero As Long, _
                          lbSoloCreditos As Boolean)


        Dim lnExoneracionCamaSP As Double, lnExoneracionSOPSP As Double, lnExoneracionRestoSP As Double, lnExoneracionItemSP As Double, lnExoneracionConsultaSP As Double
        Dim lcTexto3 As String, lcTexto2 As String, lcTexto1 As String
        Dim lRecordCount As Long, lnLineas As Long, lnIdTipoServicio As Long, lnIdCentroCosto As Long
        Dim lnPagoCta As Double, lnExoneracion As Double, lnAnulado As Double, lnImpTotal As Double
        Dim lnCama As Double, lnSoloSOP As Double, lnResto As Double, lnQueda As Double
        Dim lnAnuladoCama As Double, lnImptotalCama As Double, lnExoneracionConsulta As Double
        Dim lnAnuladoSOP As Double, lnImptotalSOP As Double
        Dim lnAnuladoResto As Double, lnImptotalResto As Double
        Dim lnConsulta As Double, lnAnuladoConsulta As Double
        Dim lnImptotalConsulta As Double
        Dim lnImpLaboratorio As Double, lnExoneracionLaboratorioSP As Double
        Dim lnImpImagenes As Double, lnExoneracionImagenesSP As Double
        Dim lnAnuladoLaboratorio As Double, lnImpTotalLaboratorio As Double
        Dim lnAnuladoImagenes As Double, lnImpTotalImagenes As Double
        Dim lbPrimeraVez As Boolean, lbEncontroDato As Boolean
        Dim rsTmp As New Recordset
        Dim rsReporte As New Recordset
        Dim oRsExoneradoBoleta As New Recordset
        Dim oRsItemsXCuenta As New Recordset
        Dim rsParamet As New Recordset
        Dim rsTmp1 As New Recordset
        Dim rsTmp2 As New Recordset
        Dim rsTmp3 As New Recordset
        
        
        
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
        Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
        Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
        Dim rsTmp999 As New Recordset
        
        Dim lnExoneracionImpDetall As Double, lnExoneracionCanDetall As Long
        Dim lnAnuladoImpDetall As Double, lnAnuladoCanDetall As Long
        Dim lnImptotalDetall As Double, lnCanTotalDetall As Long
        Dim lnIdProductoDetall As Long, lnCodigoDetall As String, lnDescripcionDetall As String, lnPrecioDetall As Double
        Dim lbHallo As Boolean, lcPie As String
        Dim lnBoletaExo As Double, lnBoletaTot As Double, lnBoletaAnu As Double
        Dim lnBoletaExo1 As Double, lnBoletaTot1 As Double, lnBoletaAnu1 As Double
        Dim lnIdPartida100 As Long, lnIdProducto100 As Long, lcCodigo100 As String, lcProducto100 As String
        Dim lnPrecioVenta100 As Double, lnCantidadPagar100 As Long, lnTotalPagar100 As Double
        Dim lnIdServiciosHospitalarios As Long, lnIdReembolsosServicios As Long
        Dim lnIdEstadoComprobante As Long, wxParametro549 As String, lnIdPartida999 As Long, lcFiltro99 As String
        '
        On Error GoTo ErrRXCC
        '
        Set rsTmp3 = mo_ReglasCaja.CajaCajaSegunFiltro("")
        '
        lnIdServiciosHospitalarios = Val(lcBuscaParametro.SeleccionaFilaParametro(245))
        lnIdReembolsosServicios = Val(lcBuscaParametro.SeleccionaFilaParametro(251))
        wxParametro549 = lcBuscaParametro.SeleccionaFilaParametro(549)
        '
        lcTexto3 = "..comienza a generar Temporal..."
        If lbSeCreaTemporal = True Then
            If ml_IdPartidaFiltro = 0 Then
                With mrs_Tmp
                   .Fields.Append "idPartida", adInteger
                   .Fields.Append "Codigo", adVarChar, 10, adFldIsNullable
                   .Fields.Append "Descripcion", adVarChar, 150, adFldIsNullable
                   .Fields.Append "ImpAnulado", adDouble
                   .Fields.Append "ImpExonerado", adDouble
                   .Fields.Append "ImpNormal", adDouble
                   .Fields.Append "ImpCancelado", adDouble
                   .Fields.Append "ImpSubTotal", adDouble
                   .LockType = adLockOptimistic
                   .Open
                End With
            Else
                With mrs_Tmp
                   .Fields.Append "Identificador", adInteger
                   .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
                   .Fields.Append "Descripcion", adVarChar, 300, adFldIsNullable
                   .Fields.Append "ImpAnulado", adDouble
                   .Fields.Append "ImpExonerado", adDouble
                   .Fields.Append "ImpNormal", adDouble
                   .Fields.Append "ImpCancelado", adDouble
                   .Fields.Append "idPartida", adInteger
                   .Fields.Append "ImpSubTotal", adDouble
                   .LockType = adLockOptimistic
                   .Open
                End With
            End If
        End If
        '
        lcTexto3 = "..llena Temporal..."
        If ml_IdPartidaFiltro = 0 And mrs_Tmp.RecordCount = 0 Then
            'Llena Temporal con Partidas
            Set rsTmp = mo_ReglasComunes.FactPartidasPresupuestalesSeleccionarTodas(moConexion)
            If rsTmp.RecordCount = 0 Then
               MsgBox "Llene tabla PARTIDAS"
               Exit Sub
            End If
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
               mrs_Tmp.AddNew
               mrs_Tmp.Fields!IdPartida = rsTmp.Fields!IdPartidaPresupuestal
               mrs_Tmp.Fields!codigo = rsTmp.Fields!codigo
               mrs_Tmp.Fields!Descripcion = rsTmp.Fields!Descripcion
               mrs_Tmp.Update
               rsTmp.MoveNext
            Loop
            rsTmp.Close
        End If
        lcTexto3 = "..comienza a jalar boletas servicios..."
        'Boletas Servicios
        
        Set rsReporte = mo_ReglasFacturacion.ServicioConsolidado(0, mda_FechaInicio, mda_FechaFin, 0, 0)
        '
        lcFiltro99 = ""
        If lnIdCajero > 0 Then
           lcFiltro99 = "idCajero=" & lnIdCajero
        End If
        If lcFiltro99 = "" Then
           lcFiltro99 = IIf(lbSoloCreditos = True, "TieneCredito='C'", "TieneCredito=null")
        Else
           lcFiltro99 = lcFiltro99 & IIf(lbSoloCreditos = True, " and TieneCredito='C'", " and TieneCredito=null")
        End If
        rsReporte.Filter = lcFiltro99
        '
        lRecordCount = rsReporte.RecordCount
        If lRecordCount > 0 Then
           Set rsTmp999 = mo_ReglasCaja.CajaComprobantesPagoXnroSerieYDocumentoFechas(mda_FechaInicio, mda_FechaFin, moConexion)
           lnLineas = 0

           rsReporte.MoveFirst
           Do While Not rsReporte.EOF
              lcTexto2 = rsReporte.Fields!NroSerie + rsReporte.Fields!nroDocumento
              lnPagoCta = rsReporte.Fields!Adelantos
              lnExoneracion = rsReporte.Fields!exoneraciones
              lnIdEstadoComprobante = rsReporte.Fields!IdEstadoComprobante
              If rsReporte.Fields!IdEstadoComprobante = 9 Then
                 lnAnulado = rsReporte.Fields!totalPagado
                 lnImpTotal = 0
              Else
                 lnAnulado = 0
                 lnImpTotal = rsReporte.Fields!totalPagado
              End If
              '
              lnIdTipoServicio = 0
              lbPrimeraVez = True
              '
              Set oRsExoneradoBoleta = mo_ReglasCaja.FacturacionServicioFinanciamientosExoneracionesEnBoleta(rsReporte.Fields!IdComprobantePago, moConexion)
              '
              lnBoletaExo = rsReporte.Fields!exoneraciones
              lnBoletaTot = rsReporte.Fields!totalPagado: lnBoletaAnu = 0
              Select Case rsReporte.Fields!IdEstadoComprobante
              Case 6   '***Devolucion
                   lnBoletaTot = -rsReporte.Fields!totalPagado
              Case 9   '***anulado
                   lnBoletaTot = 0: lnBoletaAnu = rsReporte.Fields!totalPagado
              End Select
              lnBoletaExo1 = 0: lnBoletaTot1 = 0: lnBoletaAnu1 = 0
              '
              Do While Not rsReporte.EOF And lcTexto2 = (rsReporte.Fields!NroSerie + rsReporte.Fields!nroDocumento)
If Val(rsReporte.Fields!NroSerie) = 22 And Val(rsReporte.Fields!nroDocumento) = 1372884 Then
lcTexto1 = ""
End If
                    lnLineas = lnLineas + 1
                    '
                    If lbPrimeraVez = True Then
                       lbPrimeraVez = False
                       
                       '*****Emergencia
                       lcTexto3 = "..comienza: Emergencia..."
                       lnConsulta = 0: lnResto = 0: lnImpLaboratorio = 0: lnImpImagenes = 0
                       'Set rsTmp = CajaComprobantesPagoXnroSerieYDocumento(rsReporte.Fields!NroSerie, rsReporte.Fields!NroDocumento, moConexion)
                       rsTmp999.Filter = "idComprobantePago=" & rsReporte.Fields!IdComprobantePago
                       
                       
                          If rsTmp999.RecordCount > 0 Then
                             '
                             lnExoneracionConsultaSP = 0: lnExoneracionRestoSP = 0
                             lnExoneracionLaboratorioSP = 0: lnExoneracionImagenesSP = 0
                             '
                             lbEncontroDato = True
                             rsTmp999.MoveFirst
                             Do While Not rsTmp999.EOF
                             
                                '
                                lnExoneracionItemSP = 0
                                If oRsExoneradoBoleta.RecordCount > 0 Then
                                   oRsExoneradoBoleta.MoveFirst
                                   Do While Not oRsExoneradoBoleta.EOF
                                      If oRsExoneradoBoleta.Fields!idOrden = rsTmp999.Fields!idOrden And oRsExoneradoBoleta.Fields!idProducto = rsTmp999.Fields!idProducto Then
                                         lnExoneracionItemSP = lnExoneracionItemSP + oRsExoneradoBoleta.Fields!totalFinanciado
                                      End If
                                      oRsExoneradoBoleta.MoveNext
                                   Loop
                                End If
                                '
                                lnConsulta = lnConsulta + rsTmp999.Fields!Total - lnExoneracionItemSP
                                lnExoneracionConsultaSP = lnExoneracionConsultaSP + lnExoneracionItemSP
                                '
                                If lnExoneracionItemSP > 0 Then
                                   lnExoneracionCanDetall = 1: lnExoneracionImpDetall = lnExoneracionItemSP
                                Else
                                   lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
                                End If
                                Select Case rsReporte.Fields!IdEstadoComprobante
                                Case 6        '***Devolucion
                                   lnImptotalDetall = -(rsTmp999.Fields!Importe - lnExoneracionItemSP): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                   lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                Case 9        '***Anulado
                                   lnImptotalDetall = 0: lnCanTotalDetall = 0
                                   lnAnuladoImpDetall = (rsTmp999.Fields!Importe - lnExoneracionItemSP): lnAnuladoCanDetall = rsTmp999.Fields!Cantidad
                                   lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
                                Case Else     '***Pagado
                                   lnImptotalDetall = (rsTmp999.Fields!Importe - lnExoneracionItemSP): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                   lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                   If lnPagoCta > 0 Then
                                      lnImptotalDetall = lnImptotalDetall - lnPagoCta
                                      lnPagoCta = 0
                                   End If
                                End Select
                                lnIdProductoDetall = rsTmp999.Fields!idProducto: lnCodigoDetall = rsTmp999.Fields!codigo
                                lnDescripcionDetall = rsTmp999.Fields!nombreProducto: lnPrecioDetall = rsTmp999.Fields!precio
                                '
                                lnBoletaExo1 = lnBoletaExo1 + lnExoneracionImpDetall
                                lnBoletaTot1 = lnBoletaTot1 + lnImptotalDetall
                                lnBoletaAnu1 = lnBoletaAnu1 + lnAnuladoImpDetall
                                '
                                lnIdPartida100 = rsTmp999.Fields!IdPartida
                                lnIdProducto100 = lnIdProductoDetall
                                lcCodigo100 = lnCodigoDetall
                                lcProducto100 = lnDescripcionDetall
                                lnPrecioVenta100 = lnPrecioDetall
                                lnCantidadPagar100 = rsTmp999.Fields!Cantidad
                                lnTotalPagar100 = rsTmp999.Fields!Total
                                '
                                Select Case rsTmp999.Fields!idProducto
                                Case lnIdServiciosHospitalarios And Not IsNull(rsReporte.Fields!idCuentaAtencion)
                                     '******Servicios Hospitalarios, habrá que desagregarlos en 2 Partes: Farmacia y el resto en ServiciosHospitalarios
                                     mo_ReglasFacturacion.DesagregaCuentasItemsXServHosp oRsItemsXCuenta, _
                                                          rsReporte.Fields!idCuentaAtencion, rsTmp999.Fields!Total, moConexion, _
                                                          lnIdServiciosHospitalarios
                                     If oRsItemsXCuenta.RecordCount > 0 Then
                                        oRsItemsXCuenta.MoveFirst
                                        Do While Not oRsItemsXCuenta.EOF
                                            lnExoneracionImpDetall = 0
                                            lnExoneracionCanDetall = 0
                                            Select Case rsReporte.Fields!IdEstadoComprobante
                                            Case 6        '***Devolucion
                                            Case 9        '***Anulado
                                               lnImptotalDetall = 0
                                               lnCanTotalDetall = 0
                                               lnAnuladoImpDetall = oRsItemsXCuenta.Fields!ImporteProrrateado
                                               lnAnuladoCanDetall = oRsItemsXCuenta.Fields!Cantidad
                                               lnExoneracionImpDetall = 0: lnExoneracionCanDetall = 0
                                            Case Else     '***Pagado
                                               lnImptotalDetall = oRsItemsXCuenta.Fields!ImporteProrrateado
                                               lnCanTotalDetall = oRsItemsXCuenta.Fields!Cantidad
                                               lnAnuladoImpDetall = 0
                                               lnAnuladoCanDetall = 0
                                            End Select
                                            If ml_IdPartidaFiltro > 0 Then
                                                lnIdProductoDetall = oRsItemsXCuenta.Fields!idProducto
                                                lnCodigoDetall = oRsItemsXCuenta.Fields!codigo
                                                lnDescripcionDetall = oRsItemsXCuenta.Fields!nombreProducto
                                                lnPrecioDetall = oRsItemsXCuenta.Fields!precio
                                                GrabaDetalleEnTmpPartida ml_IdPartidaFiltro, oRsItemsXCuenta.Fields!IdPartida, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, oRsItemsXCuenta.Fields!Cantidad, oRsItemsXCuenta.Fields!ImporteProrrateado, lnTipoProceso
                                            Else
                                                mrs_Tmp.MoveFirst
                                                mrs_Tmp.Find "idPartida=" & oRsItemsXCuenta.Fields!IdPartida
                                                If Not mrs_Tmp.EOF Then
                                                   mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
                                                   mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
                                                   mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
                                                   mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
                                                   mrs_Tmp.Update
                                                End If
                                            End If
                                            oRsItemsXCuenta.MoveNext
                                        Loop
                                     End If
                                Case lnIdReembolsosServicios
                                     '******Reembolsos de Servicios, habrá que desagregarlos por Cuenta->Procedimientos
                                     mo_ReglasFacturacion.DesagregaCuentasItemsXreembolso oRsItemsXCuenta, rsReporte.Fields!IdComprobantePago, _
                                                          rsTmp999.Fields!Total, moConexion
                                     If oRsItemsXCuenta.RecordCount > 0 Then
                                        oRsItemsXCuenta.MoveFirst
                                        Do While Not oRsItemsXCuenta.EOF
                                            lnExoneracionImpDetall = 0
                                            lnExoneracionCanDetall = 0
                                            Select Case rsReporte.Fields!IdEstadoComprobante
                                            Case 6        '***Devolucion
                                            Case 9        '***Anulado
                                               lnImptotalDetall = 0
                                               lnCanTotalDetall = 0
                                               lnAnuladoImpDetall = oRsItemsXCuenta.Fields!ImporteProrrateado
                                               lnAnuladoCanDetall = oRsItemsXCuenta.Fields!Cantidad
                                            Case Else     '***Pagado
                                               lnImptotalDetall = oRsItemsXCuenta.Fields!ImporteProrrateado
                                               lnCanTotalDetall = oRsItemsXCuenta.Fields!Cantidad
                                               lnAnuladoImpDetall = 0
                                               lnAnuladoCanDetall = 0
                                            End Select
                                            If ml_IdPartidaFiltro > 0 Then
                                                lnIdProductoDetall = oRsItemsXCuenta.Fields!idProducto
                                                lnCodigoDetall = oRsItemsXCuenta.Fields!codigo
                                                lnDescripcionDetall = oRsItemsXCuenta.Fields!nombreProducto
                                                lnPrecioDetall = oRsItemsXCuenta.Fields!precio
                                                GrabaDetalleEnTmpPartida ml_IdPartidaFiltro, oRsItemsXCuenta.Fields!IdPartida, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, oRsItemsXCuenta.Fields!Cantidad, oRsItemsXCuenta.Fields!ImporteProrrateado, lnTipoProceso
                                            Else
                                                mrs_Tmp.MoveFirst
                                                mrs_Tmp.Find "idPartida=" & oRsItemsXCuenta.Fields!IdPartida
                                                If Not mrs_Tmp.EOF Then
                                                   mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
                                                   mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
                                                   mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
                                                   mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
                                                   mrs_Tmp.Update
                                                End If
                                            End If
                                            oRsItemsXCuenta.MoveNext
                                        Loop
                                     End If
                                Case Else
                                     lnIdPartida999 = rsTmp999!IdPartida
                                     If lnIdProductoDetall = Val(wxParametro549) Then
                                       rsTmp3.MoveFirst
                                       rsTmp3.Find "idCaja=" & rsReporte!idCaja
                                       If Not rsTmp3.EOF Then
                                          If Not IsNull(rsTmp3!IdPartida) Then
                                             lnIdPartida999 = rsTmp3!IdPartida
                                          End If
                                       End If
                                     End If
                                     If ml_IdPartidaFiltro > 0 Then
                                        GrabaDetalleEnTmpPartida ml_IdPartidaFiltro, lnIdPartida999, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, rsTmp999.Fields!Cantidad, rsTmp999.Fields!Total, lnTipoProceso
                                     Else
                                        mrs_Tmp.MoveFirst
                                        mrs_Tmp.Find "idPartida=" & lnIdPartida999
                                        If Not mrs_Tmp.EOF Then
                                          mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
                                          mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
                                          mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
                                          mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
                                          mrs_Tmp.Update
                                        End If
                                     End If
                                End Select
                                rsTmp999.MoveNext
                             Loop
                          End If
                         ' rsTmp.Close
                    End If
                    rsReporte.MoveNext
                    If rsReporte.EOF Then
                       Exit Do
                    End If
              Loop
              '
              oRsExoneradoBoleta.Close
              If lnIdEstadoComprobante = 9 Then
                 lnBoletaExo1 = lnBoletaExo
              End If
              'Chequea Redondeo, para que cuadre con Totales de la  Boleta
              If lnBoletaExo <> lnBoletaExo1 Or lnBoletaTot <> lnBoletaTot1 Or lnBoletaAnu <> lnBoletaAnu1 Then
                    lnAnuladoImpDetall = 0
                    lnExoneracionImpDetall = 0
                    lnImptotalDetall = 0
                    If lnBoletaExo <> lnBoletaExo1 Then
                          lnExoneracionImpDetall = lnBoletaExo - lnBoletaExo1
                    End If
                    If lnBoletaTot <> lnBoletaTot1 Then
                          lnImptotalDetall = lnBoletaTot - lnBoletaTot1
                    End If
                    If lnBoletaAnu <> lnBoletaAnu1 Then
                          lnAnuladoImpDetall = lnBoletaAnu - lnBoletaAnu1
                    End If
                    If ml_IdPartidaFiltro > 0 Then
                       GrabaDetalleEnTmpPartida ml_IdPartidaFiltro, lnIdPartida100, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProducto100, lcCodigo100, lcProducto100, lnPrecioVenta100, lnCantidadPagar100, lnTotalPagar100, lnTipoProceso
                    Else
                       mrs_Tmp.MoveFirst
                       mrs_Tmp.Find "idPartida=" & lnIdPartida100
                       If Not mrs_Tmp.EOF Then
                          mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
                          mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
                          mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
                          mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
                          mrs_Tmp.Update
                       End If
                    End If
              End If
              '
           Loop
        End If
        'Farmacia
        lcTexto3 = "..comienza: MEDICAMENTOS emitidos en CAJA SERVICIO..."
        Set rsReporte = Nothing
        Set rsReporte = mo_ReglasFacturacion.FarmaciaConsolidado(0, mda_FechaInicio, mda_FechaFin, 0, 0)
        '
        lcFiltro99 = ""
        If lnIdCajero > 0 Then
           lcFiltro99 = "idCajero=" & lnIdCajero
        End If
        If lcFiltro99 = "" Then
           lcFiltro99 = IIf(lbSoloCreditos = True, "TieneCredito='C'", "TieneCredito=null")
        Else
           lcFiltro99 = lcFiltro99 & IIf(lbSoloCreditos = True, " and TieneCredito='C'", " and TieneCredito=null")
        End If
        rsReporte.Filter = lcFiltro99
        '
        lRecordCount = rsReporte.RecordCount
        If lRecordCount > 0 Then
           Set oRsExoneradoBoleta = mo_ReglasCaja.FacturacionBienesFinanciamientosExoneracionesEnBoletaTodos(moConexion)
           lnLineas = 0
           '
           rsReporte.MoveFirst
           lnExoneracion = 0: lnAnulado = 0: lnImpTotal = 0
           Do While Not rsReporte.EOF
              lcTexto2 = rsReporte.Fields!NroSerie + rsReporte.Fields!nroDocumento
              lnPagoCta = rsReporte.Fields!Adelantos
              lnExoneracion = lnExoneracion + rsReporte.Fields!exoneraciones
              If rsReporte.Fields!IdEstadoComprobante = 9 Then
                 lnAnulado = lnAnulado + rsReporte.Fields!totalPagado
              Else
                 lnImpTotal = lnImpTotal + rsReporte.Fields!totalPagado
              End If
              

              'Set oRsExoneradoBoleta = FacturacionBienesFinanciamientosExoneracionesEnBoleta(rsReporte.Fields!IdComprobantePago, moConexion)
              oRsExoneradoBoleta.Filter = "idComprobantePago=" & rsReporte.Fields!IdComprobantePago
             '
              lnBoletaExo = rsReporte.Fields!exoneraciones
              lnBoletaTot = rsReporte.Fields!totalPagado: lnBoletaAnu = 0
              lnIdEstadoComprobante = rsReporte.Fields!IdEstadoComprobante
              If rsReporte.Fields!IdEstadoComprobante = 9 Then
                 lnBoletaTot = 0: lnBoletaAnu = rsReporte.Fields!totalPagado
              End If
              lnBoletaExo1 = 0: lnBoletaTot1 = 0: lnBoletaAnu1 = 0
              '
              Do While Not rsReporte.EOF And lcTexto2 = (rsReporte.Fields!NroSerie + rsReporte.Fields!nroDocumento)
                    lnIdPartida100 = rsReporte.Fields!IdPartida
                    lnIdProducto100 = rsReporte.Fields!idProducto
                    lcCodigo100 = rsReporte.Fields!codigo
                    lcProducto100 = rsReporte.Fields!Producto
                    lnPrecioVenta100 = rsReporte.Fields!PrecioVenta
                    lnCantidadPagar100 = rsReporte.Fields!CantidadPagar
                    lnTotalPagar100 = rsReporte.Fields!TotalPagar
                    lnLineas = lnLineas + 1
                    lnExoneracionImpDetall = 0
                    lnExoneracionCanDetall = 0
                    If oRsExoneradoBoleta.RecordCount > 0 Then
                       oRsExoneradoBoleta.MoveFirst
                       Do While Not oRsExoneradoBoleta.EOF
                          If oRsExoneradoBoleta.Fields!MovTipo = rsReporte.Fields!MovTipo And oRsExoneradoBoleta.Fields!movNumero = rsReporte.Fields!movNumero And oRsExoneradoBoleta.Fields!idProducto = rsReporte.Fields!idProducto Then
                             lnExoneracionImpDetall = lnExoneracionImpDetall + oRsExoneradoBoleta.Fields!totalFinanciado
                             lnExoneracionCanDetall = lnExoneracionCanDetall + oRsExoneradoBoleta.Fields!CantidadFinanciada
                          End If
                          oRsExoneradoBoleta.MoveNext
                       Loop
                    End If
                    '
                    lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                    lnImptotalDetall = 0: lnCanTotalDetall = 0
                    
                    Select Case rsReporte.Fields!IdEstadoComprobante
                    Case 9      '***anulado
                        lnAnuladoImpDetall = rsReporte.Fields!TotalPagar - lnExoneracionImpDetall
                        lnAnuladoCanDetall = rsReporte.Fields!CantidadPagar
                        lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
                        
                    Case Else   '***pagado
                        lnImptotalDetall = rsReporte.Fields!TotalPagar - lnExoneracionImpDetall
                        lnCanTotalDetall = rsReporte.Fields!CantidadPagar
                        If lnPagoCta > 0 Then
                           lnImptotalDetall = lnImptotalDetall - lnPagoCta
                           lnPagoCta = 0
                        End If
                    End Select
                    '
                    lnBoletaExo1 = lnBoletaExo1 + lnExoneracionImpDetall
                    lnBoletaTot1 = lnBoletaTot1 + lnImptotalDetall
                    lnBoletaAnu1 = lnBoletaAnu1 + lnAnuladoImpDetall
                    '
                    If ml_IdPartidaFiltro > 0 Then
                       GrabaDetalleEnTmpPartida ml_IdPartidaFiltro, rsReporte.Fields!IdPartida, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, rsReporte.Fields!idProducto, rsReporte.Fields!codigo, rsReporte.Fields!Producto, rsReporte.Fields!PrecioVenta, rsReporte.Fields!CantidadPagar, rsReporte.Fields!TotalPagar, lnTipoProceso
                    Else
                       mrs_Tmp.MoveFirst
                       mrs_Tmp.Find "idPartida=" & rsReporte.Fields!IdPartida
                       If Not mrs_Tmp.EOF Then
                          mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
                          mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
                          mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
                          mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
                          mrs_Tmp.Update
                       End If
                    End If
                    
                    '
                    rsReporte.MoveNext
                    If rsReporte.EOF Then
                       Exit Do
                    End If
              Loop
              'oRsExoneradoBoleta.Close
              If lnIdEstadoComprobante = 9 Then
                 lnBoletaExo1 = lnBoletaExo
              End If
              'Chequea Redondeo, para que cuadre con Totales de la  Boleta
              If lnBoletaExo <> lnBoletaExo1 Or lnBoletaTot <> lnBoletaTot1 Or lnBoletaAnu <> lnBoletaAnu1 Then
                    lnAnuladoImpDetall = 0
                    lnExoneracionImpDetall = 0
                    lnImptotalDetall = 0
                    If lnBoletaExo <> lnBoletaExo1 Then
                          lnExoneracionImpDetall = lnBoletaExo - lnBoletaExo1
                    End If
                    If lnBoletaTot <> lnBoletaTot1 Then
                          lnImptotalDetall = lnBoletaTot - lnBoletaTot1
                    End If
                    If lnBoletaAnu <> lnBoletaAnu1 Then
                          lnAnuladoImpDetall = lnBoletaAnu - lnBoletaAnu1
                    End If
                    If ml_IdPartidaFiltro > 0 Then
                       GrabaDetalleEnTmpPartida ml_IdPartidaFiltro, lnIdPartida100, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProducto100, lcCodigo100, lcProducto100, lnPrecioVenta100, lnCantidadPagar100, lnTotalPagar100, lnTipoProceso
                    Else
                       mrs_Tmp.MoveFirst
                       mrs_Tmp.Find "idPartida=" & lnIdPartida100
                       If Not mrs_Tmp.EOF Then
                          mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
                          mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
                          mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
                          mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
                          mrs_Tmp.Update
                       End If
                    End If
              End If
              '
           Loop
        End If
        If mrs_Tmp.RecordCount > 0 Then
           mrs_Tmp.MoveFirst
           Do While Not mrs_Tmp.EOF
              mrs_Tmp!ImpSubTotal = mrs_Tmp!ImpNormal - (mrs_Tmp!impAnulado + mrs_Tmp!impExonerado)
              mrs_Tmp.Update
              mrs_Tmp.MoveNext
           Loop
        End If
        
        '***********************notas de Credito ***********   'kike 2017
        If ml_IdPartidaFiltro = 0 Then
            Set rsTmp = mo_ReglasCaja.NotaCreditoDevueltosPorNumYFecha("", "", mda_FechaInicio, mda_FechaFin)
            lcTexto1 = ""
            If lnIdCajero > 0 Then
               lcTexto1 = "idCajero=" & lnIdCajero
            End If
            If lcTexto1 <> "" Then
               rsTmp.Filter = lcTexto1
            End If
            If rsTmp.RecordCount > 0 Then
                lnTotalPagar100 = 0
                rsTmp.MoveFirst
                Do While Not rsTmp.EOF
                   lnTotalPagar100 = lnTotalPagar100 + rsTmp!Total
                   rsTmp.MoveNext
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!IdPartida = 12345
                mrs_Tmp.Fields!codigo = "00000"
                mrs_Tmp.Fields!Descripcion = "NOTAS DE CREDITO"
                mrs_Tmp.Fields!impAnulado = 0
                mrs_Tmp.Fields!impExonerado = 0
                mrs_Tmp.Fields!ImpNormal = -lnTotalPagar100
                mrs_Tmp.Fields!ImpCancelado = -lnTotalPagar100
                mrs_Tmp.Update
            End If
            rsTmp.Close
        End If
        
        Set rsTmp = Nothing
        Set rsReporte = Nothing
        Set oRsExoneradoBoleta = Nothing
        Set oRsItemsXCuenta = Nothing
        Set rsParamet = Nothing
        Set rsTmp1 = Nothing
        Set rsTmp2 = Nothing
        Set lcBuscaParametro = Nothing
        'Set moConexion = Nothing
        Set mo_ReglasComunes = Nothing
        Set rsTmp999 = Nothing
        Set mo_ReglasCaja = Nothing
        Set rsTmp3 = Nothing
        Exit Sub
ErrRXCC:
        If Err.Number = 94 Then
           MsgBox Err.Description & Chr(13) & Chr(13) & "Boleta N° " & lcTexto2 & "     facturacionServiciosPagos.idProducto: " & Trim(str(lnIdProductoDetall))
        Else
           MsgBox Err.Description & Chr(13) & "Boleta: " & lcTexto2
        End If
        Exit Sub
        Resume
End Sub

'debb-16/05/2016
Sub GrabaDetalleEnTmpPartida(mlIdPartidaFiltro As Long, lnIdPartida As Long, mrs_Tmp As Recordset, _
                             lnExoneracionImpDetall As Double, lnExoneracionCanDetall As Long, _
                             lnAnuladoImpDetall As Double, lnAnuladoCanDetall As Long, _
                             lnImptotalDetall As Double, lnCanTotalDetall As Long, lnIdProducto1 As Long, _
                             lcCodigo1 As String, lcDescripcion1 As String, lnPrecio1 As Double, _
                             lnCantidad1 As Long, lnTotal1 As Double, lnTipoProceso As sghTipoProceso)

        Dim lbPrimeraVez As Boolean
        Dim lbContinuaTMP As Boolean
        lbContinuaTMP = False
        If lnTipoProceso = sghProcesaYgraba Then
               lbContinuaTMP = True
        ElseIf lnTipoProceso = sghSoloParaReporte Then
            If mlIdPartidaFiltro = lnIdPartida Then
               lbContinuaTMP = True
            End If
        End If
        If lbContinuaTMP = True Then
            lbPrimeraVez = True
            If mrs_Tmp.RecordCount > 0 Then
               mrs_Tmp.MoveFirst
               Do While Not mrs_Tmp.EOF
                  If mrs_Tmp.Fields!identificador = lnIdProducto1 And mrs_Tmp.Fields!IdPartida = lnIdPartida Then
                     lbPrimeraVez = False
                     Exit Do
                  End If
                  mrs_Tmp.MoveNext
               Loop
            End If
            If lbPrimeraVez = True Then
                  mrs_Tmp.AddNew
                  mrs_Tmp.Fields!identificador = lnIdProducto1
                  mrs_Tmp.Fields!Descripcion = Left(lcDescripcion1, 150)
                  mrs_Tmp.Fields!codigo = lcCodigo1
                  mrs_Tmp.Fields!IdPartida = lnIdPartida
            End If
            mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
            mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
            mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
            mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
            mrs_Tmp.Update
        End If



End Sub





'lnDesde=1(desde boton Paquetes), 2(desde boton Copia REceta)
Function CargaDatosEnTemporalesRecetas(lnDesde As Long, oRsTmp1 As Recordset) As Recordset
        Dim oRsItems As New Recordset
        With oRsItems
              .Fields.Append "IdPuntoCarga", adInteger, 4, adFldIsNullable
              .Fields.Append "IdProducto", adInteger, 4, adFldIsNullable
              .Fields.Append "Producto", adVarChar, 255, adFldIsNullable
              .Fields.Append "cantidad", adInteger
              .CursorType = adOpenKeyset
              .LockType = adLockOptimistic
              .Open
        End With
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              If lnDesde = 1 Then      'desde boton PAQUETES
                    oRsItems.AddNew
                    oRsItems.Fields!IdPuntoCarga = oRsTmp1!IdPuntoCarga
                    oRsItems.Fields!idProducto = oRsTmp1!idProducto
                    oRsItems.Fields!Producto = Left(Trim(oRsTmp1!codigo) & "//" & oRsTmp1!Descripcion, 255)
                    oRsItems.Fields!Cantidad = oRsTmp1!Cantidad
                    oRsItems.Update
              ElseIf lnDesde = 2 Then   'desde boton COPIAR RECETA ANTERIOR
                    If oRsTmp1!IdPuntoCarga = sghPtoCargaFarmacia Then
                        oRsItems.AddNew
                        oRsItems.Fields!IdPuntoCarga = sghPtoCargaFarmacia
                        oRsItems.Fields!idProducto = oRsTmp1!idItem
                        oRsItems.Fields!Producto = Left(Trim(oRsTmp1!codigo) & "//" & oRsTmp1!nombre, 255)
                        oRsItems.Fields!Cantidad = oRsTmp1!cantidadPedida
                        oRsItems.Update
                    Else
                        oRsItems.AddNew
                        oRsItems.Fields!IdPuntoCarga = oRsTmp1!IdPuntoCarga
                        oRsItems.Fields!idProducto = oRsTmp1!idItem
                        oRsItems.Fields!Producto = Left(Trim(oRsTmp1!codigo) & "//" & oRsTmp1!nombre, 255)
                        oRsItems.Fields!Cantidad = oRsTmp1!cantidadPedida
                        oRsItems.Update
                    End If
              End If
              oRsTmp1.MoveNext
           Loop
        End If
        Set CargaDatosEnTemporalesRecetas = oRsItems

End Function


'debb-11/04/2016
Sub AgregaItemsDeReceta(lnIdPuntoCarga As sghPuntosCargaBasicos, oRsItemsElegidos As Recordset, _
                ByRef oRsPatologia As Recordset, ByRef oRsAnatomia As Recordset, _
                ByRef oRsBanco As Recordset, ByRef oRsRayosX As Recordset, _
                ByRef oRsEcografiaO As Recordset, ByRef oRsEcografiaG As Recordset, _
                ByRef oRsTomografia As Recordset, ByRef oRsFarmacia As Recordset, _
                lnMaximoItems As Long, lnIdDosisDefault As Long, ml_IdTipoFinanciamiento As Long, _
                lbVariosFormatosFua As Boolean, lnIdFarmaciaElegida As Long, lbSeEligioPaquete As Boolean, _
                oRsDevuelveTodosLosItemsServ As Recordset, oRsDevuelveTodosLosItemsFarm As Recordset, lcDx As String)
    Dim lcCabecera As String, lnPrecio As Double
    Dim lnIdProducto As Long, lcProducto As String
    Dim lbContinuar As Boolean, lnTotalReg As Long
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim lbEstaVacia As Boolean
    Select Case lnIdPuntoCarga
    Case sghPtoCargaPatologiaClinica
        If oRsDevuelveTodosLosItemsServ.RecordCount > 0 Then
           lbEstaVacia = IIf(oRsPatologia.RecordCount = 0, True, False)
           oRsDevuelveTodosLosItemsServ.MoveFirst
           Do While Not oRsDevuelveTodosLosItemsServ.EOF
                lnIdProducto = oRsDevuelveTodosLosItemsServ!idProducto
                If lnIdProducto > 0 Then
                    lnPrecio = DevuelvePrecioItem(lnIdProducto, sghPtoCargaPatologiaClinica, ml_IdTipoFinanciamiento)
                    If lbEstaVacia = False Then
                        oRsPatologia.MoveFirst
                        oRsPatologia.Find "id=" & lnIdProducto
                        If oRsPatologia.EOF Then
                            oRsPatologia.AddNew
                            If lbVariosFormatosFua = True Then
                               oRsPatologia.Fields!fua = 1
                            End If
                            oRsPatologia.Fields!ID = lnIdProducto
                            oRsPatologia.Fields!procedimiento = Trim(oRsDevuelveTodosLosItemsServ!codigo) & "//" & oRsDevuelveTodosLosItemsServ!nombreProducto
                        End If
                    Else
                        oRsPatologia.AddNew
                        If lbVariosFormatosFua = True Then
                           oRsPatologia.Fields!fua = 1
                        End If
                        oRsPatologia.Fields!ID = lnIdProducto
                        oRsPatologia.Fields!procedimiento = Trim(oRsDevuelveTodosLosItemsServ!codigo) & "//" & oRsDevuelveTodosLosItemsServ!nombreProducto
                        sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, Trim(oRsDevuelveTodosLosItemsServ!codigo)
                    End If
                    oRsPatologia.Fields!Cantidad = oRsDevuelveTodosLosItemsServ!Cantidad
                    oRsPatologia.Fields!precio = lnPrecio
                    If lnPrecio > 0 Then
                       oRsPatologia.Fields!hayCpt = True
                    End If
                    oRsPatologia.Fields!saldoActual = 0
                    oRsPatologia.Fields!Observaciones = oRsDevuelveTodosLosItemsServ!Observaciones
                    If lcDx <> "" Then
                       oRsPatologia.Fields!dx = lcDx
                    End If
                    oRsPatologia.Update
                    
                End If
                oRsDevuelveTodosLosItemsServ.MoveNext
           Loop
        ElseIf oRsItemsElegidos.RecordCount > 0 Then
           oRsItemsElegidos.MoveFirst
           Do While Not oRsItemsElegidos.EOF
                lnIdProducto = oRsItemsElegidos.Fields!idProducto
                lcProducto = oRsItemsElegidos.Fields!Producto
                lbContinuar = True
                'debb-24/06/2015
                lnTotalReg = oRsPatologia.RecordCount
                If lnTotalReg >= lnMaximoItems Then
                   MsgBox "Solo puede registrar hasta " & Trim(str(lnMaximoItems)) & " items", vbInformation, "Receta"
                   lbContinuar = False
                End If
                '
                If lnTotalReg > 0 And lbContinuar = True Then        'debb-24/06/2015
                   oRsPatologia.MoveFirst
                   oRsPatologia.Find "id=" & lnIdProducto
                   If Not oRsPatologia.EOF Then
                      lbContinuar = False
                   End If
                End If
                If lbContinuar = True Then
                    lnPrecio = DevuelvePrecioItem(lnIdProducto, sghPtoCargaPatologiaClinica, ml_IdTipoFinanciamiento)
                    oRsPatologia.AddNew
                    If lbVariosFormatosFua = True Then
                       oRsPatologia.Fields!fua = 1
                    End If
                    oRsPatologia.Fields!ID = lnIdProducto
                    oRsPatologia.Fields!procedimiento = lcProducto
                    If lbSeEligioPaquete = True Then
                       oRsPatologia.Fields!Cantidad = oRsItemsElegidos!Cantidad
                    Else
                       oRsPatologia.Fields!Cantidad = 1
                    End If
                    oRsPatologia.Fields!precio = lnPrecio
                    If lnPrecio > 0 Then
                       oRsPatologia.Fields!hayCpt = True
                    End If
                    oRsPatologia.Fields!saldoActual = 0
                    If lcDx <> "" Then
                       oRsPatologia.Fields!dx = lcDx
                    End If
                    oRsPatologia.Update
                    sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, Left(lcProducto, 7)
                End If
                oRsItemsElegidos.MoveNext
           Loop
        End If
    Case sghPtoCargaAnatomiaPatologica1
        If oRsDevuelveTodosLosItemsServ.RecordCount > 0 Then
           lbEstaVacia = IIf(oRsAnatomia.RecordCount = 0, True, False)
           oRsDevuelveTodosLosItemsServ.MoveFirst
           Do While Not oRsDevuelveTodosLosItemsServ.EOF
                lnIdProducto = oRsDevuelveTodosLosItemsServ!idProducto
                If lnIdProducto > 0 Then
                    lnPrecio = DevuelvePrecioItem(lnIdProducto, sghPtoCargaAnatomiaPatologica1, ml_IdTipoFinanciamiento)
                    If lbEstaVacia = False Then
                        oRsAnatomia.MoveFirst
                        oRsAnatomia.Find "id=" & lnIdProducto
                        If oRsAnatomia.EOF Then
                            oRsAnatomia.AddNew
                            If lbVariosFormatosFua = True Then
                               oRsAnatomia.Fields!fua = 1
                            End If
                            oRsAnatomia.Fields!ID = lnIdProducto
                            oRsAnatomia.Fields!procedimiento = Trim(oRsDevuelveTodosLosItemsServ!codigo) & "//" & oRsDevuelveTodosLosItemsServ!nombreProducto
                        End If
                    Else
                        oRsAnatomia.AddNew
                        If lbVariosFormatosFua = True Then
                           oRsAnatomia.Fields!fua = 1
                        End If
                        oRsAnatomia.Fields!ID = lnIdProducto
                        oRsAnatomia.Fields!procedimiento = Trim(oRsDevuelveTodosLosItemsServ!codigo) & "//" & oRsDevuelveTodosLosItemsServ!nombreProducto
                        sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, Trim(oRsDevuelveTodosLosItemsServ!codigo)
                    End If
                    oRsAnatomia.Fields!Cantidad = oRsDevuelveTodosLosItemsServ!Cantidad
                    oRsAnatomia.Fields!precio = lnPrecio
                    If lnPrecio > 0 Then
                       oRsAnatomia.Fields!hayCpt = True
                    End If
                    oRsAnatomia.Fields!saldoActual = 0
                    oRsAnatomia.Fields!Observaciones = oRsDevuelveTodosLosItemsServ!Observaciones
                    If lcDx <> "" Then
                       oRsAnatomia.Fields!dx = lcDx
                    End If
                    oRsAnatomia.Update
                End If
                oRsDevuelveTodosLosItemsServ.MoveNext
           Loop
        ElseIf oRsItemsElegidos.RecordCount > 0 Then
           oRsItemsElegidos.MoveFirst
           Do While Not oRsItemsElegidos.EOF
                lnIdProducto = oRsItemsElegidos.Fields!idProducto
                lcProducto = oRsItemsElegidos.Fields!Producto
                lbContinuar = True
                'debb-24/06/2015
                lnTotalReg = oRsAnatomia.RecordCount
                If lnTotalReg >= lnMaximoItems Then
                   MsgBox "Solo puede registrar hasta " & Trim(str(lnMaximoItems)) & " items", vbInformation, "Receta"
                   lbContinuar = False
                End If
                '
                If lnTotalReg > 0 And lbContinuar = True Then        'debb-24/06/2015
                   oRsAnatomia.MoveFirst
                   oRsAnatomia.Find "id=" & lnIdProducto
                   If Not oRsAnatomia.EOF Then
                      lbContinuar = False
                   End If
                End If

                If lbContinuar = True Then
                    lnPrecio = DevuelvePrecioItem(lnIdProducto, sghPtoCargaAnatomiaPatologica1, ml_IdTipoFinanciamiento)
                    oRsAnatomia.AddNew
                    If lbVariosFormatosFua = True Then
                       oRsAnatomia.Fields!fua = 1
                    End If
                    oRsAnatomia.Fields!ID = lnIdProducto
                    oRsAnatomia.Fields!procedimiento = lcProducto
                    If lbSeEligioPaquete = True Then
                       oRsAnatomia.Fields!Cantidad = oRsItemsElegidos!Cantidad
                    Else
                       oRsAnatomia.Fields!Cantidad = 1
                    End If
                    oRsAnatomia.Fields!precio = lnPrecio
                    If lnPrecio > 0 Then
                       oRsAnatomia.Fields!hayCpt = True
                    End If
                    oRsAnatomia.Fields!saldoActual = 0
                    If lcDx <> "" Then
                       oRsAnatomia.Fields!dx = lcDx
                    End If
                    oRsAnatomia.Update
                    sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, Left(lcProducto, 7)

                End If
                oRsItemsElegidos.MoveNext
           Loop
        End If
    Case sghPtoCargaBancoSangre1
        If oRsDevuelveTodosLosItemsServ.RecordCount > 0 Then
           lbEstaVacia = IIf(oRsBanco.RecordCount = 0, True, False)
           oRsDevuelveTodosLosItemsServ.MoveFirst
           Do While Not oRsDevuelveTodosLosItemsServ.EOF
                lnIdProducto = oRsDevuelveTodosLosItemsServ!idProducto
                If lnIdProducto > 0 Then
                    lnPrecio = DevuelvePrecioItem(lnIdProducto, sghPtoCargaBancoSangre1, ml_IdTipoFinanciamiento)
                    If lbEstaVacia = False Then
                        oRsBanco.MoveFirst
                        oRsBanco.Find "id=" & lnIdProducto
                        If oRsBanco.EOF Then
                            oRsBanco.AddNew
                            If lbVariosFormatosFua = True Then
                               oRsBanco.Fields!fua = 1
                            End If
                            oRsBanco.Fields!ID = lnIdProducto
                            oRsBanco.Fields!procedimiento = Trim(oRsDevuelveTodosLosItemsServ!codigo) & "//" & oRsDevuelveTodosLosItemsServ!nombreProducto
                            sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, Trim(oRsDevuelveTodosLosItemsServ!codigo)
                        End If
                    Else
                        oRsBanco.AddNew
                        If lbVariosFormatosFua = True Then
                           oRsBanco.Fields!fua = 1
                        End If
                        oRsBanco.Fields!ID = lnIdProducto
                        oRsBanco.Fields!procedimiento = Trim(oRsDevuelveTodosLosItemsServ!codigo) & "//" & oRsDevuelveTodosLosItemsServ!nombreProducto
                    End If
                    oRsBanco.Fields!Cantidad = oRsDevuelveTodosLosItemsServ!Cantidad
                    oRsBanco.Fields!precio = lnPrecio
                    If lnPrecio > 0 Then
                       oRsBanco.Fields!hayCpt = True
                    End If
                    oRsBanco.Fields!saldoActual = 0
                    oRsBanco.Fields!Observaciones = oRsDevuelveTodosLosItemsServ!Observaciones
                    If lcDx <> "" Then
                       oRsBanco.Fields!dx = lcDx
                    End If
                    oRsBanco.Update
                End If
                oRsDevuelveTodosLosItemsServ.MoveNext
           Loop
        ElseIf oRsItemsElegidos.RecordCount > 0 Then
           oRsItemsElegidos.MoveFirst
           Do While Not oRsItemsElegidos.EOF
                lnIdProducto = oRsItemsElegidos.Fields!idProducto
                lcProducto = oRsItemsElegidos.Fields!Producto
                lbContinuar = True
                'debb-24/06/2015
                lnTotalReg = oRsBanco.RecordCount
                If lnTotalReg >= lnMaximoItems Then
                   MsgBox "Solo puede registrar hasta " & Trim(str(lnMaximoItems)) & " items", vbInformation, "Receta"
                   lbContinuar = False
                End If
                '
                If lnTotalReg > 0 And lbContinuar = True Then        'debb-24/06/2015
                   oRsBanco.MoveFirst
                   oRsBanco.Find "id=" & lnIdProducto
                   If Not oRsBanco.EOF Then
                      lbContinuar = False
                   End If
                End If

                If lbContinuar = True Then
                    lnPrecio = DevuelvePrecioItem(lnIdProducto, sghPtoCargaBancoSangre1, ml_IdTipoFinanciamiento)
                    oRsBanco.AddNew
                    If lbVariosFormatosFua = True Then
                       oRsBanco.Fields!fua = 1
                    End If
                    oRsBanco.Fields!ID = lnIdProducto
                    oRsBanco.Fields!procedimiento = lcProducto
                    If lbSeEligioPaquete = True Then
                       oRsBanco.Fields!Cantidad = oRsItemsElegidos!Cantidad
                    Else
                       oRsBanco.Fields!Cantidad = 1
                    End If
                    oRsBanco.Fields!precio = lnPrecio
                    If lnPrecio > 0 Then
                       oRsBanco.Fields!hayCpt = True
                    End If
                    oRsBanco.Fields!saldoActual = 0
                    If lcDx <> "" Then
                       oRsBanco.Fields!dx = lcDx
                    End If
                    oRsBanco.Update
                    sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, Left(lcProducto, 7)
                End If
                oRsItemsElegidos.MoveNext
           Loop
        End If
    Case sghPtoCargaRayosX
        If oRsDevuelveTodosLosItemsServ.RecordCount > 0 Then
           lbEstaVacia = IIf(oRsRayosX.RecordCount = 0, True, False)
           oRsDevuelveTodosLosItemsServ.MoveFirst
           Do While Not oRsDevuelveTodosLosItemsServ.EOF
                lnIdProducto = oRsDevuelveTodosLosItemsServ!idProducto
                If lnIdProducto > 0 Then
                    lnPrecio = DevuelvePrecioItem(lnIdProducto, sghPtoCargaRayosX, ml_IdTipoFinanciamiento)
                    If lbEstaVacia = False Then
                        oRsRayosX.MoveFirst
                        oRsRayosX.Find "id=" & lnIdProducto
                        If oRsRayosX.EOF Then
                            oRsRayosX.AddNew
                            If lbVariosFormatosFua = True Then
                               oRsRayosX.Fields!fua = 1
                            End If
                            oRsRayosX.Fields!ID = lnIdProducto
                            oRsRayosX.Fields!procedimiento = Trim(oRsDevuelveTodosLosItemsServ!codigo) & "//" & oRsDevuelveTodosLosItemsServ!nombreProducto
                            sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, Trim(oRsDevuelveTodosLosItemsServ!codigo)
                        End If
                    Else
                        oRsRayosX.AddNew
                        If lbVariosFormatosFua = True Then
                           oRsRayosX.Fields!fua = 1
                        End If
                        oRsRayosX.Fields!ID = lnIdProducto
                        oRsRayosX.Fields!procedimiento = Trim(oRsDevuelveTodosLosItemsServ!codigo) & "//" & oRsDevuelveTodosLosItemsServ!nombreProducto
                    End If
                    oRsRayosX.Fields!Cantidad = oRsDevuelveTodosLosItemsServ!Cantidad
                    oRsRayosX.Fields!precio = lnPrecio
                    If lnPrecio > 0 Then
                       oRsRayosX.Fields!hayCpt = True
                    End If
                    oRsRayosX.Fields!saldoActual = 0
                    oRsRayosX.Fields!Observaciones = oRsDevuelveTodosLosItemsServ!Observaciones
                    If lcDx <> "" Then
                       oRsRayosX.Fields!dx = lcDx
                    End If
                    oRsRayosX.Update
                End If
                oRsDevuelveTodosLosItemsServ.MoveNext
           Loop
        ElseIf oRsItemsElegidos.RecordCount > 0 Then
           oRsItemsElegidos.MoveFirst
           Do While Not oRsItemsElegidos.EOF
                lnIdProducto = oRsItemsElegidos.Fields!idProducto
                lcProducto = oRsItemsElegidos.Fields!Producto
                lbContinuar = True
                'debb-24/06/2015
                lnTotalReg = oRsRayosX.RecordCount
                If lnTotalReg >= lnMaximoItems Then
                   MsgBox "Solo puede registrar hasta " & Trim(str(lnMaximoItems)) & " items", vbInformation, "Receta"
                   lbContinuar = False
                End If
                '
                If lnTotalReg > 0 And lbContinuar = True Then     'debb-24/06/2015
                   oRsRayosX.MoveFirst
                   oRsRayosX.Find "id=" & lnIdProducto
                   If Not oRsRayosX.EOF Then
                      lbContinuar = False
                   End If
                End If
                If lbContinuar = True Then
                    lnPrecio = DevuelvePrecioItem(lnIdProducto, sghPtoCargaRayosX, ml_IdTipoFinanciamiento)
                    oRsRayosX.AddNew
                    If lbVariosFormatosFua = True Then
                       oRsRayosX.Fields!fua = 1
                    End If
                    oRsRayosX.Fields!ID = lnIdProducto
                    oRsRayosX.Fields!procedimiento = lcProducto
                    If lbSeEligioPaquete = True Then
                       oRsRayosX.Fields!Cantidad = oRsItemsElegidos!Cantidad
                    Else
                       oRsRayosX.Fields!Cantidad = 1
                    End If
                    oRsRayosX.Fields!precio = lnPrecio
                    If lnPrecio > 0 Then
                       oRsRayosX.Fields!hayCpt = True
                    End If
                    oRsRayosX.Fields!saldoActual = 0
                    If lcDx <> "" Then
                       oRsRayosX.Fields!dx = lcDx
                    End If
                    oRsRayosX.Update
                    sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, Left(lcProducto, 7)
                End If
                oRsItemsElegidos.MoveNext
           Loop
        End If
    Case sghPtoCargaEcogObstetrica
        If oRsDevuelveTodosLosItemsServ.RecordCount > 0 Then
           lbEstaVacia = IIf(oRsEcografiaO.RecordCount = 0, True, False)
           oRsDevuelveTodosLosItemsServ.MoveFirst
           Do While Not oRsDevuelveTodosLosItemsServ.EOF
                lnIdProducto = oRsDevuelveTodosLosItemsServ!idProducto
                If lnIdProducto > 0 Then
                    lnPrecio = DevuelvePrecioItem(lnIdProducto, sghPtoCargaEcogObstetrica, ml_IdTipoFinanciamiento)
                    If lbEstaVacia = False Then
                        oRsEcografiaO.MoveFirst
                        oRsEcografiaO.Find "id=" & lnIdProducto
                        If oRsEcografiaO.EOF Then
                            oRsEcografiaO.AddNew
                            If lbVariosFormatosFua = True Then
                               oRsEcografiaO.Fields!fua = 1
                            End If
                            oRsEcografiaO.Fields!ID = lnIdProducto
                            oRsEcografiaO.Fields!procedimiento = Trim(oRsDevuelveTodosLosItemsServ!codigo) & "//" & oRsDevuelveTodosLosItemsServ!nombreProducto
                            sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, Trim(oRsDevuelveTodosLosItemsServ!codigo)
                        End If
                    Else
                        oRsEcografiaO.AddNew
                        If lbVariosFormatosFua = True Then
                           oRsEcografiaO.Fields!fua = 1
                        End If
                        oRsEcografiaO.Fields!ID = lnIdProducto
                        oRsEcografiaO.Fields!procedimiento = Trim(oRsDevuelveTodosLosItemsServ!codigo) & "//" & oRsDevuelveTodosLosItemsServ!nombreProducto
                    End If
                    oRsEcografiaO.Fields!Cantidad = oRsDevuelveTodosLosItemsServ!Cantidad
                    oRsEcografiaO.Fields!precio = lnPrecio
                    If lnPrecio > 0 Then
                       oRsEcografiaO.Fields!hayCpt = True
                    End If
                    oRsEcografiaO.Fields!saldoActual = 0
                    oRsEcografiaO.Fields!Observaciones = oRsDevuelveTodosLosItemsServ!Observaciones
                    If lcDx <> "" Then
                       oRsEcografiaO.Fields!dx = lcDx
                    End If
                    oRsEcografiaO.Update
                End If
                oRsDevuelveTodosLosItemsServ.MoveNext
           Loop
        ElseIf oRsItemsElegidos.RecordCount > 0 Then
           oRsItemsElegidos.MoveFirst
           Do While Not oRsItemsElegidos.EOF
                  lnIdProducto = oRsItemsElegidos.Fields!idProducto
                  lcProducto = oRsItemsElegidos.Fields!Producto
                  lbContinuar = True
                  'debb-24/06/2015
                  lnTotalReg = oRsEcografiaO.RecordCount
                  If lnTotalReg >= lnMaximoItems Then
                     MsgBox "Solo puede registrar hasta " & Trim(str(lnMaximoItems)) & " items", vbInformation, "Receta"
                     lbContinuar = False
                  End If
                  '
                  If lnTotalReg > 0 And lbContinuar = True Then        'debb-24/06/2015
                     oRsEcografiaO.MoveFirst
                     oRsEcografiaO.Find "id=" & lnIdProducto
                     If Not oRsEcografiaO.EOF Then
                        lbContinuar = False
                     End If
                  End If

                  If lbContinuar = True Then
                        lnPrecio = DevuelvePrecioItem(lnIdProducto, sghPtoCargaEcogObstetrica, ml_IdTipoFinanciamiento)
                        oRsEcografiaO.AddNew
                        If lbVariosFormatosFua = True Then
                           oRsEcografiaO.Fields!fua = 1
                        End If
                        oRsEcografiaO.Fields!ID = lnIdProducto
                        oRsEcografiaO.Fields!procedimiento = lcProducto
                        If lbSeEligioPaquete = True Then
                           oRsEcografiaO.Fields!Cantidad = oRsItemsElegidos!Cantidad
                        Else
                           oRsEcografiaO.Fields!Cantidad = 1
                        End If
                        oRsEcografiaO.Fields!precio = lnPrecio
                        If lnPrecio > 0 Then
                           oRsEcografiaO.Fields!hayCpt = True
                        End If
                        oRsEcografiaO.Fields!saldoActual = 0
                        If lcDx <> "" Then
                           oRsEcografiaO.Fields!dx = lcDx
                        End If
                        oRsEcografiaO.Update
                        sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, Left(lcProducto, 7)
                  End If
                  oRsItemsElegidos.MoveNext
           Loop
        End If
    Case sghPtoCargaEcogGeneral
        If oRsDevuelveTodosLosItemsServ.RecordCount > 0 Then
           lbEstaVacia = IIf(oRsEcografiaG.RecordCount = 0, True, False)
           oRsDevuelveTodosLosItemsServ.MoveFirst
           Do While Not oRsDevuelveTodosLosItemsServ.EOF
                lnIdProducto = oRsDevuelveTodosLosItemsServ!idProducto
                If lnIdProducto > 0 Then
                    lnPrecio = DevuelvePrecioItem(lnIdProducto, sghPtoCargaEcogGeneral, ml_IdTipoFinanciamiento)
                    If lbEstaVacia = False Then
                        oRsEcografiaG.MoveFirst
                        oRsEcografiaG.Find "id=" & lnIdProducto
                        If oRsEcografiaG.EOF Then
                            oRsEcografiaG.AddNew
                            If lbVariosFormatosFua = True Then
                               oRsEcografiaG.Fields!fua = 1
                            End If
                            oRsEcografiaG.Fields!ID = lnIdProducto
                            oRsEcografiaG.Fields!procedimiento = Trim(oRsDevuelveTodosLosItemsServ!codigo) & "//" & oRsDevuelveTodosLosItemsServ!nombreProducto
                            sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, Trim(oRsDevuelveTodosLosItemsServ!codigo)
                        End If
                    Else
                        oRsEcografiaG.AddNew
                        If lbVariosFormatosFua = True Then
                           oRsEcografiaG.Fields!fua = 1
                        End If
                        oRsEcografiaG.Fields!ID = lnIdProducto
                        oRsEcografiaG.Fields!procedimiento = Trim(oRsDevuelveTodosLosItemsServ!codigo) & "//" & oRsDevuelveTodosLosItemsServ!nombreProducto
                    End If
                    oRsEcografiaG.Fields!Cantidad = oRsDevuelveTodosLosItemsServ!Cantidad
                    oRsEcografiaG.Fields!precio = lnPrecio
                    If lnPrecio > 0 Then
                       oRsEcografiaG.Fields!hayCpt = True
                    End If
                    oRsEcografiaG.Fields!saldoActual = 0
                    oRsEcografiaG.Fields!Observaciones = oRsDevuelveTodosLosItemsServ!Observaciones
                    If lcDx <> "" Then
                       oRsEcografiaG.Fields!dx = lcDx
                    End If
                    oRsEcografiaG.Update
                End If
                oRsDevuelveTodosLosItemsServ.MoveNext
           Loop
        ElseIf oRsItemsElegidos.RecordCount > 0 Then
           oRsItemsElegidos.MoveFirst
           Do While Not oRsItemsElegidos.EOF
                lnIdProducto = oRsItemsElegidos.Fields!idProducto
                lcProducto = oRsItemsElegidos.Fields!Producto
                lbContinuar = True
                'debb-24/06/2015
                lnTotalReg = oRsEcografiaG.RecordCount
                If lnTotalReg >= lnMaximoItems Then
                   MsgBox "Solo puede registrar hasta " & Trim(str(lnMaximoItems)) & " items", vbInformation, "Receta"
                   lbContinuar = False
                End If
                '
                If lnTotalReg > 0 And lbContinuar = True Then        'debb-24/06/2015
                   oRsEcografiaG.MoveFirst
                   oRsEcografiaG.Find "id=" & lnIdProducto
                   If Not oRsEcografiaG.EOF Then
                      lbContinuar = False
                   End If
                End If

                If lbContinuar = True Then
                     lnPrecio = DevuelvePrecioItem(lnIdProducto, sghPtoCargaEcogGeneral, ml_IdTipoFinanciamiento)
                    oRsEcografiaG.AddNew
                    If lbVariosFormatosFua = True Then
                        oRsEcografiaG.Fields!fua = 1
                    End If
                    oRsEcografiaG.Fields!ID = lnIdProducto
                    oRsEcografiaG.Fields!procedimiento = lcProducto
                    If lbSeEligioPaquete = True Then
                       oRsEcografiaG.Fields!Cantidad = oRsItemsElegidos!Cantidad
                    Else
                       oRsEcografiaG.Fields!Cantidad = 1
                    End If
                    oRsEcografiaG.Fields!precio = lnPrecio
                    If lnPrecio > 0 Then
                       oRsEcografiaG.Fields!hayCpt = True
                    End If
                    oRsEcografiaG.Fields!saldoActual = 0
                    If lcDx <> "" Then
                       oRsEcografiaG.Fields!dx = lcDx
                    End If
                    oRsEcografiaG.Update
                    sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, Left(lcProducto, 7)
                End If
                oRsItemsElegidos.MoveNext
           Loop
        End If
    Case sghPtoCargaTomografia
        If oRsDevuelveTodosLosItemsServ.RecordCount > 0 Then
           lbEstaVacia = IIf(oRsTomografia.RecordCount = 0, True, False)
           oRsDevuelveTodosLosItemsServ.MoveFirst
           Do While Not oRsDevuelveTodosLosItemsServ.EOF
                lnIdProducto = oRsDevuelveTodosLosItemsServ!idProducto
                If lnIdProducto > 0 Then
                    lnPrecio = DevuelvePrecioItem(lnIdProducto, sghPtoCargaTomografia, ml_IdTipoFinanciamiento)
                    If lbEstaVacia = False Then
                        oRsTomografia.MoveFirst
                        oRsTomografia.Find "id=" & lnIdProducto
                        If oRsTomografia.EOF Then
                            oRsTomografia.AddNew
                            If lbVariosFormatosFua = True Then
                               oRsTomografia.Fields!fua = 1
                            End If
                            oRsTomografia.Fields!ID = lnIdProducto
                            oRsTomografia.Fields!procedimiento = Trim(oRsDevuelveTodosLosItemsServ!codigo) & "//" & oRsDevuelveTodosLosItemsServ!nombreProducto
                            sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, Trim(oRsDevuelveTodosLosItemsServ!codigo)
                        End If
                    Else
                        oRsTomografia.AddNew
                        If lbVariosFormatosFua = True Then
                           oRsTomografia.Fields!fua = 1
                        End If
                        oRsTomografia.Fields!ID = lnIdProducto
                        oRsTomografia.Fields!procedimiento = Trim(oRsDevuelveTodosLosItemsServ!codigo) & "//" & oRsDevuelveTodosLosItemsServ!nombreProducto
                    End If
                    oRsTomografia.Fields!Cantidad = oRsDevuelveTodosLosItemsServ!Cantidad
                    oRsTomografia.Fields!precio = lnPrecio
                    If lnPrecio > 0 Then
                       oRsTomografia.Fields!hayCpt = True
                    End If
                    oRsTomografia.Fields!saldoActual = 0
                    oRsTomografia.Fields!Observaciones = oRsDevuelveTodosLosItemsServ!Observaciones
                    If lcDx <> "" Then
                       oRsTomografia.Fields!dx = lcDx
                    End If
                    oRsTomografia.Update
                End If
                oRsDevuelveTodosLosItemsServ.MoveNext
           Loop
        ElseIf oRsItemsElegidos.RecordCount > 0 Then
           oRsItemsElegidos.MoveFirst
           Do While Not oRsItemsElegidos.EOF
                lnIdProducto = oRsItemsElegidos.Fields!idProducto
                lcProducto = oRsItemsElegidos.Fields!Producto
                lbContinuar = True
                'debb-24/06/2015
                lnTotalReg = oRsTomografia.RecordCount
                If lnTotalReg >= lnMaximoItems Then
                   MsgBox "Solo puede registrar hasta " & Trim(str(lnMaximoItems)) & " items", vbInformation, "Receta"
                   lbContinuar = False
                End If
                '
                If lnTotalReg > 0 And lbContinuar = True Then        'debb-24/06/2015
                   oRsTomografia.MoveFirst
                   oRsTomografia.Find "id=" & lnIdProducto
                   If Not oRsTomografia.EOF Then
                      lbContinuar = False
                   End If
                End If
                If lbContinuar = True Then
                    lnPrecio = DevuelvePrecioItem(lnIdProducto, sghPtoCargaTomografia, ml_IdTipoFinanciamiento)
                    oRsTomografia.AddNew
                    If lbVariosFormatosFua = True Then
                       oRsTomografia.Fields!fua = 1
                    End If
                    oRsTomografia.Fields!ID = lnIdProducto
                    oRsTomografia.Fields!procedimiento = lcProducto
                    If lbSeEligioPaquete = True Then
                       oRsTomografia.Fields!Cantidad = oRsItemsElegidos!Cantidad
                    Else
                       oRsTomografia.Fields!Cantidad = 1
                    End If
                    oRsTomografia.Fields!precio = lnPrecio
                    If lnPrecio > 0 Then
                       oRsTomografia.Fields!hayCpt = True
                    End If
                    oRsTomografia.Fields!saldoActual = 0
                    If lcDx <> "" Then
                       oRsTomografia.Fields!dx = lcDx
                    End If
                    oRsTomografia.Update
                    sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, Left(lcProducto, 7)
                End If
                oRsItemsElegidos.MoveNext
           Loop
        End If
    Case sghPtoCargaFarmacia
        If oRsDevuelveTodosLosItemsFarm.RecordCount > 0 Then
           lbEstaVacia = IIf(oRsFarmacia.RecordCount = 0, True, False)
           oRsDevuelveTodosLosItemsFarm.MoveFirst
           Do While Not oRsDevuelveTodosLosItemsFarm.EOF
                lnIdProducto = oRsDevuelveTodosLosItemsFarm!idProducto
                If lnIdProducto > 0 Then
                    If lbEstaVacia = False Then
                        oRsFarmacia.MoveFirst
                        oRsFarmacia.Find "id=" & lnIdProducto
                        If oRsFarmacia.EOF Then
                            oRsFarmacia.AddNew
                            If lbVariosFormatosFua = True Then
                               oRsFarmacia.Fields!fua = 1
                            End If
                            oRsFarmacia.Fields!ID = oRsDevuelveTodosLosItemsFarm!idProducto
                            oRsFarmacia.Fields!procedimiento = Trim(oRsDevuelveTodosLosItemsFarm!codigo) & "//" & oRsDevuelveTodosLosItemsFarm!nombreProducto
                        End If
                    Else
                        oRsFarmacia.AddNew
                        If lbVariosFormatosFua = True Then
                           oRsFarmacia.Fields!fua = 1
                        End If
                        oRsFarmacia.Fields!ID = oRsDevuelveTodosLosItemsFarm!idProducto
                        oRsFarmacia.Fields!procedimiento = Trim(oRsDevuelveTodosLosItemsFarm!codigo) & "//" & oRsDevuelveTodosLosItemsFarm!nombreProducto
                        sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, Trim(oRsDevuelveTodosLosItemsFarm!codigo)
                    End If
                    oRsFarmacia.Fields!Cantidad = oRsDevuelveTodosLosItemsFarm!Cantidad
                    oRsFarmacia.Fields!haySaldo = True
                    oRsFarmacia.Fields!saldoActual = 0
                    oRsFarmacia.Fields!IdAlmacen = lnIdFarmaciaElegida
                    oRsFarmacia.Fields!Almacen = ""
                    oRsFarmacia.Fields!precio = DevuelvePrecioItem(lnIdProducto, sghPtoCargaFarmacia, ml_IdTipoFinanciamiento)
                    oRsFarmacia.Fields!idDosisRecetada = oRsDevuelveTodosLosItemsFarm!idDosisRecetada
                    oRsFarmacia.Fields!IdViaAdministracion = oRsDevuelveTodosLosItemsFarm!IdViaAdministracion
                    oRsFarmacia.Fields!Observaciones = oRsDevuelveTodosLosItemsFarm!Observaciones
                    If lcDx <> "" Then
                       oRsFarmacia.Fields!dx = lcDx
                    End If
                    oRsFarmacia.Update
                End If
                oRsDevuelveTodosLosItemsFarm.MoveNext
           Loop
        ElseIf Not oRsItemsElegidos Is Nothing Then
            If oRsItemsElegidos.RecordCount > 0 Then
               oRsItemsElegidos.MoveFirst
               Do While Not oRsItemsElegidos.EOF
                    lnIdProducto = oRsItemsElegidos.Fields!idProducto
                    lcProducto = IIf(IsNull(oRsItemsElegidos.Fields!Producto), "", oRsItemsElegidos.Fields!Producto)
                    lbContinuar = True
                    'debb-24/06/2015
                    lnTotalReg = oRsFarmacia.RecordCount
                    If lnTotalReg >= lnMaximoItems Then
                       MsgBox "Solo puede registrar hasta " & Trim(str(lnMaximoItems)) & " items", vbInformation, "Receta"
                       lbContinuar = False
                    End If
                    '
                    If lnTotalReg > 0 And lbContinuar = True Then        'debb-24/06/2015
                       oRsFarmacia.MoveFirst
                       oRsFarmacia.Find "id=" & lnIdProducto
                       If Not oRsFarmacia.EOF Then
                          lbContinuar = False
                       End If
                    End If
    
                    If lbContinuar = True Then
                        Dim oRsTmp As New Recordset
                        Dim lnSaldoActual As Long, lcAlmacen As String, lnIdAlmacen As Long
                        Set oRsTmp = mo_ReglasFarmacia.farmSaldoSoloFarmaciasSismed(lnIdProducto)
                        
                        If lnIdFarmaciaElegida > 1 Then
                           oRsTmp.Filter = "idAlmacen=" & Trim(str(lnIdFarmaciaElegida))
                        End If

                        lnSaldoActual = 0
                        lnPrecio = 0
                        lcAlmacen = ""
                        If oRsTmp.RecordCount > 0 Then
                           oRsTmp.MoveFirst
                           Do While Not oRsTmp.EOF
                              lnIdAlmacen = oRsTmp.Fields!IdAlmacen
                              lnSaldoActual = 0
                              lnPrecio = oRsTmp.Fields!precio
                              lcAlmacen = oRsTmp.Fields!Descripcion
                              Do While Not oRsTmp.EOF And lnIdAlmacen = oRsTmp.Fields!IdAlmacen
                                 lnSaldoActual = lnSaldoActual + oRsTmp.Fields!Cantidad
                                 oRsTmp.MoveNext
                                 If oRsTmp.EOF Then
                                    Exit Do
                                 End If
                              Loop
                              If lnSaldoActual > 1 Then
                                 Exit Do
                              End If
                           Loop
                        End If
                        oRsTmp.Close
                        Set oRsTmp = Nothing
                        '
                        oRsFarmacia.AddNew
                        If lbVariosFormatosFua = True Then
                           oRsFarmacia.Fields!fua = 1
                        End If
                        oRsFarmacia.Fields!ID = lnIdProducto
                        oRsFarmacia.Fields!procedimiento = lcProducto
                        If lbSeEligioPaquete = True Then
                           oRsFarmacia.Fields!Cantidad = oRsItemsElegidos!Cantidad
                        Else
                           oRsFarmacia.Fields!Cantidad = 1
                        End If
                        oRsFarmacia.Fields!haySaldo = IIf(lnSaldoActual > 0, True, False)
                        oRsFarmacia.Fields!saldoActual = lnSaldoActual
                        oRsFarmacia.Fields!IdAlmacen = lnIdAlmacen
                        oRsFarmacia.Fields!Almacen = lcAlmacen
                        oRsFarmacia.Fields!precio = DevuelvePrecioItem(lnIdProducto, sghPtoCargaFarmacia, ml_IdTipoFinanciamiento)
                        oRsFarmacia.Fields!idDosisRecetada = lnIdDosisDefault
                        oRsFarmacia.Fields!IdViaAdministracion = RecetaDetalleUltimaViaAdministracion(lnIdProducto)    'debb-19/08/2015
'                        oRsFarmacia.Fields!IdViaAdministracion = 0 'Actualizado 26092014
                        If lcDx <> "" Then
                           oRsFarmacia.Fields!dx = lcDx
                        End If
                        oRsFarmacia.Update
                        sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, Left(lcProducto, 7)
                    End If
                    oRsItemsElegidos.MoveNext
               Loop
            End If
        End If
    End Select
    Set mo_ReglasFarmacia = Nothing
End Sub


Function DevuelvePrecioItem(lnIdProducto As Long, lnIdPuntoCarga As Long, ml_IdTipoFinanciamiento As Long, Optional oConexion1 As Connection) As Double
      Dim oRsTmp As New Recordset
      Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
      If lnIdPuntoCarga = sghPtoCargaFarmacia Then
         Set oRsTmp = mo_ReglasComunes.FactCatalogoBienesInsumosHospXfiltro("idProducto=" & lnIdProducto & " and idTipoFinanciamiento=" & ml_IdTipoFinanciamiento)
      Else
         Set oRsTmp = mo_ReglasComunes.FactCatalogoServiciosHospXfiltro("idProducto=" & lnIdProducto & " and idTipoFinanciamiento=" & ml_IdTipoFinanciamiento)
      End If
      DevuelvePrecioItem = 0
      If oRsTmp.RecordCount > 0 Then
         DevuelvePrecioItem = oRsTmp.Fields!PrecioUnitario
      End If
      Set mo_ReglasComunes = Nothing
End Function


'debb-19/08/2015
Function RecetaDetalleUltimaViaAdministracion(lnIdProductoFarmacia As Long) As Long
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetaDetalleXultimaViaAdministracion"
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProductoFarmacia): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    RecetaDetalleUltimaViaAdministracion = 0
    If oRs.RecordCount > 0 Then
       If Not IsNull(oRs!IdViaAdministracion) Then
          RecetaDetalleUltimaViaAdministracion = oRs!IdViaAdministracion
       End If
    End If
    oConexion.Close
    Set oCommand = Nothing
    Set oConexion = Nothing
End Function



Function LLegoAlTopeCitasSIS(ldDiario As Date, ldHoy As Date, lnMaxCuposCitasAdelantadasSIS As Long, _
        lnCuposCitasAdelantadasSIS As Long, lnMaxCuposCitasHoySIS As Long, lnCuposCitasHoySIS As Long) As String
    LLegoAlTopeCitasSIS = ""
    If CDate(Format(ldDiario, sighentidades.DevuelveFechaSoloFormato_DMY)) > ldHoy Then
        If lnMaxCuposCitasAdelantadasSIS <= lnCuposCitasAdelantadasSIS Then
           LLegoAlTopeCitasSIS = "Ya llegó al tope de CUPOS SIS ADELANTADOS para ese CONSULTORIO (citas mayores a HOY)" & Chr(13) & "que es: " & Trim(str(lnMaxCuposCitasAdelantadasSIS))
        End If
    ElseIf CDate(Format(ldDiario, sighentidades.DevuelveFechaSoloFormato_DMY)) = ldHoy Then
        If lnMaxCuposCitasHoySIS <= lnCuposCitasHoySIS Then
           LLegoAlTopeCitasSIS = "Ya llegó al tope de CUPOS SIS para ese CONSULTORIO, que es: " & Trim(str(lnMaxCuposCitasHoySIS))
        End If
    End If
End Function




Sub ExportaDAtosAlHISv4(lcTarde As String, lnChkExportaCPT As Integer, lnIdMesElegido As Integer, lcANIO As String, _
                        lbExportaAnovafis As Boolean, lbAgregaOrdenesMedicas As Boolean, lbExportaAlHisMinsa As Boolean)
              
10            On Error GoTo ErrorProceso
              Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
              Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
              Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
              Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
              Dim mo_ReglasServGeograf As New SIGHNegocios.ReglasServGeograf
              
              Dim oConexion As New Connection
              Dim oConexionFox As New Connection
             
              Dim rsReporte As New Recordset
              Dim oRsTmp As New Recordset
              Dim oRsFox1 As New Recordset
              Dim oRsFox2 As New Recordset
              Dim oRsFox3 As New Recordset
              Dim oRsFox4 As New Recordset
              Dim oRsFox5 As New Recordset
              Dim oRsEdadHis As New Recordset
              Dim lcDx1 As String, lcDx2 As String, lcDx3 As String, lcDx4 As String, lcDx5 As String, lcDx6 As String
              Dim lcTDx1 As String, lcTDx2 As String, lcTDx3 As String, lcTDx4 As String, lcTDx5 As String, lcTDx6 As String
              Dim lcCDx1 As String, lcCDx2 As String, lcCDx3 As String, lcCDx4 As String, lcCDx5 As String, lcCDx6 As String
              Dim lcLDx1 As String, lcLDx2 As String, lcLDx3 As String, lcLDx4 As String, lcLDx5 As String, lcLDx6 As String
              Dim lnTot_reg As Long, lnEstancia As Integer
              Dim lcFec1 As String, lcFec2 As String
              Dim lcMes As String, lnEdadEnAtencion As Integer, lnAsc As Integer
              Dim lnUltimoDiaMes As Integer, lbContinuar As Boolean
              Dim lnIdEdad As Long, lnDesde As Integer, lnHasta As Integer
              Dim lcDistritoDomicilio As String, lcCodigoServicioHIS As String
              Dim lnTotal As Long, lnIdEmpleado As Long, lnNumPag As Long, lnReg As Long
              Dim lnTTot_pag As Long, lnTTot_reg As Long, lnMaxPacientesXhoja As Long, lnMaxHISxLOTE As Long
              Dim lcCod_dig As String, lcCod_2000 As String, lcMt As String, lcNomLote As String
              Dim ldFecha As Date, lbAumentoPagina As Boolean, lbEsCentroSalud As Boolean
              Dim lcCodif As String, lcPlaza As String, lcCod_servsa As String, lcTipoDocumentoHisDoctor As String
              Dim lcFichaOhistoria As String, lcDNIpaciente As String, lcFuenteFinanciamientoPaciente As String, lcEtniaPaciente As String
              
              Dim lnGrupo6DxCPT As Integer
              Dim lnRestoDxCPT As Integer
              Dim lnIndiceDxCPT As Integer
              Dim lnContador As Integer
              Dim lnContDxCpt As Integer
              Const lcCharInicio As String = "0"
              Dim oRsImpHIStodos As New Recordset, lnoRsImpHIStodos As Long
              Dim rsDx1 As New Recordset
              Dim oConexionExterna As New Connection
              Dim oRsNovafis_edades As New Recordset
              Dim lcAD040 As String
              Dim lcRutaExportar As String, lcPaisPaciente As String, lcTipoDocumentoPaciente As String
              Dim lnEdadEnTablaEdades As Long, lcGestante As String, lbNuevoRegistro As Boolean, lcDNIpaciente1 As String
              Dim lnDNImedico1 As String
              Dim lnAscFinal1 As Integer, lnAscFinal2 As Integer, lcFilter As String
              '
20            lcAD040 = lcBuscaParametro.SeleccionaFilaParametro(354)
              '
30            oConexionExterna.CursorLocation = adUseClient
40            oConexionExterna.CommandTimeout = 150
50            oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
              '
              If lbExportaAnovafis = True Then
60               Set oRsNovafis_edades = mo_ReglasComunes.novafis_edadesSeleccionarTodos
              End If
70            lcRutaExportar = lcBuscaParametro.SeleccionaFilaParametro(313)
              '
80            lcMes = Right("0" & Trim(str(lnIdMesElegido)), 2)
              'lcAnio = cmbAnio.Text
90            lcFec1 = ("01/" & lcMes & "/" & lcANIO)
100           lnUltimoDiaMes = DevuelveUltimoDiaDelMes(Val(lcMes), Val(lcANIO))
110           lcFec2 = Right("0" & Trim(str(lnUltimoDiaMes)), 2) & "/" & lcMes & "/" & lcANIO
120           lnMaxPacientesXhoja = Val(lcBuscaParametro.SeleccionaFilaParametro(272))
130           lnMaxHISxLOTE = Val(lcBuscaParametro.SeleccionaFilaParametro(271))
140           If lbExportaAnovafis = True Then
                 '***********exporta al NOVAFIS ****************
150              lcCod_dig = Val(lcBuscaParametro.SeleccionaFilaParametro(270))
160           Else
                 '***********exporta al HIS ****************
170              lcCod_dig = Val(Right(lcBuscaParametro.SeleccionaFilaParametro(270), 3))
180           End If
190           lbEsCentroSalud = IIf(lcBuscaParametro.SeleccionaFilaParametro(282) = "S", True, False)
              '
200           oConexion.CommandTimeout = 300
210           oConexion.CursorLocation = adUseClient
220           oConexion.Open sighentidades.CadenaConexion
              '
230           oConexionFox.CommandTimeout = 300
240           oConexionFox.Open "DSN=his"
              
              '
250           mo_ReglasComunes.PreparaTablasDBFhis oRsFox1, oRsFox2, oRsFox3, oConexionFox
260           If lbExportaAnovafis = True Then
                 '***********exporta al NOVAFIS ****************
270              Set oRsFox2 = Nothing
280              With oRsFox2
290                   .Fields.Append "COD_ESTAB", adVarChar, 9, adFldIsNullable
300                   .Fields.Append "ANO", adInteger
310                   .Fields.Append "MES", adInteger
320                   .Fields.Append "NOM_LOTE", adVarChar, 3, adFldIsNullable
330                   .Fields.Append "tot_pag", adInteger
340                   .CursorType = adOpenDynamic
350                   .LockType = adLockOptimistic
360                   .Open
370              End With
380              Set oRsFox1 = Nothing
390              With oRsFox1
400                   .Fields.Append "COD_ESTAB", adVarChar, 9, adFldIsNullable
410                   .Fields.Append "ANO", adInteger
420                   .Fields.Append "MES", adInteger
430                   .Fields.Append "NOM_LOTE", adVarChar, 3, adFldIsNullable
440                   .Fields.Append "NUM_PAG", adInteger
450                   .Fields.Append "NUM_REG", adInteger
460                   .Fields.Append "DIA", adInteger
470                   .Fields.Append "FICHAFAM", adVarChar, 10, adFldIsNullable
480                   .Fields.Append "UBIGEO", adVarChar, 6, adFldIsNullable
490                   .Fields.Append "EDAD", adInteger
500                   .Fields.Append "TIP_EDAD", adVarChar, 1, adFldIsNullable
510                   .Fields.Append "SEXO", adVarChar, 1, adFldIsNullable
520                   .Fields.Append "ESTABLEC", adVarChar, 1, adFldIsNullable
530                   .Fields.Append "SERVICIO", adVarChar, 1, adFldIsNullable

540                   .Fields.Append "TD1", adVarChar, 1, adFldIsNullable
550                   .Fields.Append "LC1", adVarChar, 3, adFldIsNullable
560                   .Fields.Append "DX1", adVarChar, 6, adFldIsNullable

570                   .Fields.Append "TD2", adVarChar, 1, adFldIsNullable
580                   .Fields.Append "LC2", adVarChar, 3, adFldIsNullable
590                   .Fields.Append "DX2", adVarChar, 6, adFldIsNullable

600                   .Fields.Append "TD3", adVarChar, 1, adFldIsNullable
610                   .Fields.Append "LC3", adVarChar, 3, adFldIsNullable
620                   .Fields.Append "DX3", adVarChar, 6, adFldIsNullable
630                   .Fields.Append "TD4", adVarChar, 1, adFldIsNullable
640                   .Fields.Append "LC4", adVarChar, 3, adFldIsNullable
650                   .Fields.Append "DX4", adVarChar, 6, adFldIsNullable
660                   .Fields.Append "TD5", adVarChar, 1, adFldIsNullable
670                   .Fields.Append "LC5", adVarChar, 3, adFldIsNullable
680                   .Fields.Append "DX5", adVarChar, 6, adFldIsNullable
690                   .Fields.Append "TD6", adVarChar, 1, adFldIsNullable
700                   .Fields.Append "LC6", adVarChar, 3, adFldIsNullable
710                   .Fields.Append "DX6", adVarChar, 6, adFldIsNullable
720                   .Fields.Append "CODPSAL", adVarChar, 11, adFldIsNullable
730                   .Fields.Append "DIGITA", adVarChar, 11, adFldIsNullable
740                   .Fields.Append "COD_SERVSA", adVarChar, 6, adFldIsNullable
750                   .Fields.Append "MT", adVarChar, 1, adFldIsNullable
760                   .Fields.Append "CF", adVarChar, 2, adFldIsNullable
770                   .Fields.Append "ET", adVarChar, 3, adFldIsNullable
780                   .Fields.Append "PAIS", adVarChar, 3, adFldIsNullable
790                   .Fields.Append "COD_DOC", adVarChar, 1, adFldIsNullable
800                   .Fields.Append "FAMILIA", adVarChar, 2, adFldIsNullable
810                   .Fields.Append "DNI", adVarChar, 11, adFldIsNullable
820                   .Fields.Append "EST", adVarChar, 3, adFldIsNullable
830                   .Fields.Append "ED", adInteger
840                   .Fields.Append "ID", adVarChar, 30, adFldIsNullable
                      .Fields.Append "idCuenta", adInteger
850                   .CursorType = adOpenDynamic
860                   .LockType = adLockOptimistic
870                   .Open
880                   .Sort = "cod_estab,ano,mes,nom_lote"
890              End With
900           End If
              '
910           CrearTablaTemporalDxCPT 'Frank 12092014
              '
920           Set rsReporte = mo_ReglasAdmision.AtencionesSeleccionarPorFechasDeIngresoYtipoServicioOconexion(lcMes, lcANIO, oConexion)
      'rsReporte.Filter = "HISfuenteFinanciamiento='1'"
930           lnTotal = rsReporte.RecordCount
940           If lnTotal = 0 Then
950              MsgBox "No existe Información  con esos Datos", vbInformation, "Mensaje"
960           Else
970             ' mo_ProgressRpt1.Min = 0: mo_ProgressRpt1.Max = lnTotal: mo_ProgressRpt1.ShowText = True: mo_ProgressRpt1.Color = vbGreen
980              lnTotal = 0
990              lcCod_2000 = Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9)
1000             ms_MensajeError = ""
                 
1010             rsReporte.MoveFirst
1020             Do While Not rsReporte.EOF
1030                  If IsNull(rsReporte.Fields!loteHis) Then
1040                     ms_MensajeError = ms_MensajeError & "Tiene que ingresar el 'LOTE HIS' para el Médico: " & rsReporte.Fields!dEmpleado & Chr(13)
1050                  ElseIf lbExportaAnovafis = True And UCase(Left(rsReporte!loteHis, 1)) = "V" Then
1060                     ms_MensajeError = ms_MensajeError & "No puede empezar con la letra   V     el 'LOTE HIS' para el Médico: " & rsReporte.Fields!dEmpleado & Chr(13)
1070                  ElseIf lbExportaAnovafis = True And InStr("0123456789", Left(rsReporte!loteHis, 1)) > 0 Then
1080                     ms_MensajeError = ms_MensajeError & "No puede empezar con   NUMERO  el 'LOTE HIS' para el Médico: " & rsReporte.Fields!dEmpleado & Chr(13)
1090                  End If
1100                  lcMt = IIf(rsReporte.Fields!HoraIngreso >= lcTarde, "T", "M")
1110                  lcNomLote = rsReporte.Fields!loteHis & lcCharInicio    'rsReporte.Fields!LoteHis & lcMt
1120                  lnIdEmpleado = rsReporte.Fields!IdEmpleado
1130                  lnTTot_pag = 0: lnTTot_reg = 0
1140                  If IsNull(rsReporte.Fields!codigoServicioHIS) Then
1150                     MsgBox "No se pudo procesar" & Chr(13) & "Tiene que ingresar el 'Codigo HIS' para el Servicio: " & rsReporte.Fields!DServicio
1160                     Exit Do
1170                  End If
1180                  lbAumentoPagina = False
1190                  Do While Not rsReporte.EOF And lnIdEmpleado = rsReporte.Fields!IdEmpleado
                          '
1200                      lcMt = IIf(rsReporte.Fields!HoraIngreso >= lcTarde, "T", "M")
1210                      lnAsc = Asc(lcCharInicio)
                          lnAscFinal1 = 90: lnAscFinal2 = 90
1220                      Do While True
1230                         If InStr("01234567890ABCDEFGHIJKLMNOPQRSTUVWYZabcdefghijklmnopqrstuvwyz", Right(lcNomLote, 1)) = 0 Then
1240                            If Second(Time) <= 10 Then
1250                               lcNomLote = UCase(Left(lcNomLote, 2)) & Chr(lnAsc)
1260                            ElseIf Second(Time) > 10 And Second(Time) <= 20 Then
1270                               lcNomLote = UCase(Left(lcNomLote, 2)) & Chr(lnAsc)
1280                            ElseIf Second(Time) > 20 And Second(Time) <= 30 Then
1290                               lcNomLote = UCase(Mid(lcNomLote, 1, 1)) & UCase(Mid(lcNomLote, 2, 1)) & Chr(lnAsc)
1300                            Else
1310                               lcNomLote = UCase(Mid(lcNomLote, 1, 1)) & UCase(Mid(lcNomLote, 2, 1)) & Chr(lnAsc)
1320                            End If
1330                            lnAsc = lnAsc + 1
1340                            If lnAsc >= 123 Then
1350                               lnAsc = Asc(lcCharInicio)
1360                            End If
1370                         Else
1380                              On Error Resume Next
1390                              mo_ReglasComunes.HisLoteXfiltro oRsFox4, oConexionFox, "ano= " & lcANIO & " and mes=" & lcMes & " and nom_lote='" & lcNomLote & "'"
1400                              If oRsFox4.RecordCount = 0 Then
1410                                 On Error GoTo ErrorProceso
1420                                 Exit Do
1430                              Else
1440                                 lcSql = " "
1450                                  If Second(Time) <= 10 Then
1460                                     lcNomLote = UCase(Left(lcNomLote, 2)) & Chr(lnAsc)
1470                                  ElseIf Second(Time) > 10 And Second(Time) <= 20 Then
1480                                     lcNomLote = UCase(Left(lcNomLote, 2)) & Chr(lnAsc)
1490                                  ElseIf Second(Time) > 20 And Second(Time) <= 30 Then
1500                                     lcNomLote = UCase(Mid(lcNomLote, 1, 1)) & UCase(Mid(lcNomLote, 2, 1)) & Chr(lnAsc)
1510                                  Else
1520                                     lcNomLote = UCase(Mid(lcNomLote, 1, 1)) & UCase(Mid(lcNomLote, 2, 1)) & Chr(lnAsc)
1530                                  End If
1540                                  lnAsc = lnAsc + 1
1550                                  If lnAsc >= 123 Then

                                            '*************
                                            lcNomLote = Chr(lnAscFinal1) & Chr(lnAscFinal2)
                                            lnAscFinal2 = lnAscFinal2 - 1
                                            If lnAscFinal2 = 65 Then
                                               lnAscFinal1 = lnAscFinal1 - 1
                                               lnAscFinal2 = 90
                                            End If
                                             '******

                                            lnAsc = Asc(lcCharInicio)
1570                                  End If
1580                              End If
1590                              On Error GoTo ErrorProceso
1600                         End If
1610                      Loop
                          '
1620                      lnTTot_pag = 0: lnTTot_reg = 0
                          '
1630                      ldFecha = rsReporte.Fields!FechaIngreso
                          '
1640                      If IsNull(rsReporte.Fields!codigoServicioHIS) Then
1650                          ms_MensajeError = ms_MensajeError & "Tiene que ingresar el 'Codigo HIS' para el Servicio: " & rsReporte.Fields!DServicio & Chr(13)
1660                          MsgBox ms_MensajeError, vbInformation, "Mensajes"
1670                          Exit Sub
1680                          lcCodigoServicioHIS = ""
1690                      Else
1700                          lcCodigoServicioHIS = rsReporte.Fields!codigoServicioHIS
1710                      End If
1720                      lcCod_servsa = lcCodigoServicioHIS
1730                      Set oRsImpHIStodos = mo_ReglasAdmision.ServiciosAtenSimultaneaImpHISxUPS(lcCodigoServicioHIS)
1740                      lnoRsImpHIStodos = oRsImpHIStodos.RecordCount
                          
                          '
1750                      lnDNImedico1 = "00000000"
1760                      If IsNull(rsReporte.Fields!dni) Or IsNull(rsReporte.Fields!idColegioHIS) Then
1770                          ms_MensajeError = ms_MensajeError & "El profesional de Salud : " & rsReporte.Fields!dEmpleado & " deberá tener 'Colegio Médico' y 'DNI'" & Chr(13)
1780                          lcCodif = ""
1790                      Else
1800                          Set oRsTmp = mo_ReglasAdmision.TiposDocIdentidadXfiltro("idDocIdentidad=" & rsReporte.Fields!idTipoDocumento, oConexion)
1810                          lcTipoDocumentoHisDoctor = "1"
1820                          If oRsTmp.RecordCount > 0 Then
1830                             lcTipoDocumentoHisDoctor = oRsTmp.Fields!CodigoHIS
1840                          End If
1850                          oRsTmp.Close
1860                          lcCodif = lcTipoDocumentoHisDoctor & Left(rsReporte.Fields!dni, 8) & rsReporte.Fields!idColegioHIS
1870                          lnDNImedico1 = Left(rsReporte.Fields!dni, 8)
1880                      End If
                          '
1890                      lcPlaza = lcCodif
                          '
                          '
1900                      If lbAumentoPagina = True Then
1910                         lnNumPag = lnNumPag + 1
1920                      Else
1930                         lnNumPag = 1
1940                      End If
1950                      lnTot_reg = 0: lnReg = 0
1960                      Do While Not rsReporte.EOF And lnIdEmpleado = rsReporte.Fields!IdEmpleado And lcCodigoServicioHIS = rsReporte.Fields!codigoServicioHIS And ldFecha = rsReporte.Fields!FechaIngreso
1970                          lcFuenteFinanciamientoPaciente = IIf(IsNull(rsReporte.Fields!HISfuenteFinanciamiento), "", rsReporte.Fields!HISfuenteFinanciamiento)
1980                          lcEtniaPaciente = IIf(IsNull(rsReporte.Fields!IdEtnia), "", rsReporte.Fields!IdEtnia)
                              '1112333333334455
1990                          lcDNIpaciente = Space$(3): lcPaisPaciente = Space$(3)                     '1)Pais de Procedencia
2000                          If Not IsNull(rsReporte.Fields!idPaisProcedencia) Then
2010                              Set oRsTmp = mo_ReglasServGeograf.PaisesXfiltro("idPais=" & rsReporte.Fields!idPaisProcedencia)
2020                              If oRsTmp.RecordCount > 0 Then
2030                                 lcDNIpaciente = Right(Space$(3) & oRsTmp.Fields!codigo, 3)
2040                                 lcPaisPaciente = Right(Space$(3) & oRsTmp.Fields!codigo, 3)
2050                              End If
2060                              oRsTmp.Close
2070                          End If
2080                          lcTipoDocumentoPaciente = Space$(1)
2090                          If Not IsNull(rsReporte.Fields!HIStipoDocumento) Then    '2)Tipo de Documento de Identidad
2100                             lcDNIpaciente = lcDNIpaciente & rsReporte.Fields!HIStipoDocumento
2110                             lcTipoDocumentoPaciente = rsReporte.Fields!HIStipoDocumento
2120                          Else
2130                             lcDNIpaciente = lcDNIpaciente & IIf(IsNull(rsReporte.Fields!nroDocumento), Space$(1), "1")
2140                          End If
2150                          lcDNIpaciente1 = "00000000"
2160                          If Not IsNull(rsReporte.Fields!nroDocumento) Then          '3)Numero de Documento de identidad
2170                             lcDNIpaciente = lcDNIpaciente & Right(Trim(rsReporte.Fields!nroDocumento), 8)
                                    If Val(lcTipoDocumentoPaciente) > 0 Then 'solo novafis
2180                                  lcDNIpaciente1 = Left(rsReporte.Fields!nroDocumento, 8)
                                    End If
2190                          Else
2200                             lcDNIpaciente = lcDNIpaciente & Space$(8)
2210                          End If
2220                          lcDNIpaciente = lcDNIpaciente & Space$(4)                  '4)CUATRO espacios en BLANCO
                                                                                        '5)Numero de Hijos
2230                          Set oRsTmp = mo_ReglasAdmision.atencionesDatosAdicionalesXfiltro("dbo.AtencionesDatosAdicionales.idAtencion=" & rsReporte.Fields!idAtencion, oConexion)
2240                          If oRsTmp.RecordCount > 0 Then
2250                             If Not IsNull(oRsTmp.Fields!NumeroDeHijos) Then
2260                                 lcDNIpaciente = lcDNIpaciente & Right("0" & Trim(str(oRsTmp.Fields!NumeroDeHijos)), 2)
2270                             Else
2280                                 lcDNIpaciente = lcDNIpaciente & "00"
2290                             End If
2300                          Else
2310                             lcDNIpaciente = lcDNIpaciente & "00"
2320                          End If
2330                          oRsTmp.Close
                              
                              '
2340                          If rsReporte.Fields!NroHistoriaClinica = 511302 Then
2350                          lcFichaOhistoria = ""
2360                          End If
                              
2370                          If IsNull(rsReporte.Fields!NroHistoriaClinica) And IsNull(rsReporte.Fields!FichaFamiliar) Then
2380                             MsgBox "El Paciente No existe para la CUENTA N° " & rsReporte.Fields!idCuentaAtencion & Chr(13) & "Revise la tabla ATENCIONES", vbInformation, "Mensaje"
2390                             Exit Sub
2400                          End If
2410                          lcFichaOhistoria = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(str(rsReporte.Fields!NroHistoriaClinica)), False)
                               
2420                          If Not IsNull(rsReporte.Fields!FichaFamiliar) And lbEsCentroSalud = True Then
2430                             If Len(Trim(rsReporte.Fields!FichaFamiliar)) > 2 Then
2440                                lcFichaOhistoria = Trim(rsReporte.Fields!FichaFamiliar)
2450                             End If
2460                          End If
                              '
2470                          lnIndiceDxCPT = 0
2480                          If lnoRsImpHIStodos = 0 Then
2490                              LimpiarTablaDxCpt
2500                              If oRsTmp.State = 1 Then oRsTmp.Close
2510                              Set oRsTmp = mo_ReglasAdmision.AtencionesDiagnosticosXidAtencion(rsReporte.Fields!idAtencion, oConexion)
2520                              If oRsTmp.RecordCount > 0 Then
2530                                 oRsTmp.MoveFirst
2540                                 Do While Not oRsTmp.EOF
2550                                    If IsNull(oRsTmp.Fields!DescripcionTipoDx) Then
2560                                       MsgBox "No existe DIAGNOSTICO para la CUENTA N° " & rsReporte.Fields!idCuentaAtencion & Chr(13) & "Revise la tabla ATENCIONES", vbInformation, "Mensaje"
2570                                       Exit Sub
2580                                    End If
2590                                    lnIndiceDxCPT = lnIndiceDxCPT + 1
2600                                    oRsDxCpt.AddNew
2610                                    oRsDxCpt.Fields!IdDxCpt = lnIndiceDxCPT
2620                                    oRsDxCpt.Fields!dx = sighentidades.DevuelveCodigoDxSinPUNTO(Left(oRsTmp.Fields!CodigoCie2004, 6))
2630                                    oRsDxCpt.Fields!TDx = IIf(IsNull(oRsTmp.Fields!DescripcionTipoDx), "D", Left(oRsTmp.Fields!DescripcionTipoDx, 1))
2640                                    oRsDxCpt.Fields!CDx = IIf(IsNull(oRsTmp.Fields!ClaseDxHIS), "1", Left(oRsTmp.Fields!ClaseDxHIS, 1))
2650                                    oRsDxCpt.Fields!LDx = IIf(IsNull(oRsTmp.Fields!labConfHIS), "", oRsTmp.Fields!labConfHIS)
2660                                    oRsDxCpt.Update
2670                                    oRsTmp.MoveNext
2680                                 Loop
2690                              End If
2700                              oRsTmp.Close
                                  'CPT
2710                              If lbAgregaOrdenesMedicas = True Or lnChkExportaCPT = 1 Then
                                       lcFilter = "idPuntoCArga=1"
                                       If sighentidades.Parametro561 = "S" Then
                                             lcFilter = lcFilter + " or idPuntoCarga=99"
                                       End If
2720                                   If oRsTmp.State = 1 Then oRsTmp.Close
2730                                   Set oRsTmp = mo_ReglasAdmision.BuscaOtrosCPTyExamenesparaFormatoHIS(rsReporte.Fields!idCuentaAtencion, 1)
                                       If lbAgregaOrdenesMedicas = False Then
                                          oRsTmp.Filter = lcFilter
                                       End If
2740                                   If oRsTmp.RecordCount > 0 Then
2750                                        oRsTmp.MoveFirst
2760                                        Do While Not oRsTmp.EOF
2770                                          lnIndiceDxCPT = lnIndiceDxCPT + 1
2780                                          oRsDxCpt.AddNew
2790                                          oRsDxCpt.Fields!IdDxCpt = lnIndiceDxCPT
2800                                          oRsDxCpt.Fields!dx = oRsTmp.Fields("codigo").Value
2810                                          oRsDxCpt.Fields!TDx = "D"
2820                                          oRsDxCpt.Fields!CDx = ""
2830                                          oRsDxCpt.Fields!LDx = oRsTmp.Fields("labConfHIS").Value
2840                                          oRsDxCpt.Update
2850                                          oRsTmp.MoveNext
2860                                        Loop
2870                                   End If
2880                                   oRsTmp.Close
2890                              End If
2900                          Else

2910                              If rsDx1.State = 1 Then rsDx1.Close
2920                              Set rsDx1 = mo_ReglasAdmision.BuscaAtencionesDxCEparaFormatoHIS(rsReporte.Fields!idAtencion)
2930                              If lbAgregaOrdenesMedicas = True Or lnChkExportaCPT = 1 Then
                                       'CPT
                                       lcFilter = "idPuntoCArga=1"
                                       If sighentidades.Parametro561 = "S" Then
                                            lcFilter = lcFilter + " or idPuntoCarga=99"
                                       End If
2940                                   If oRsTmp.State = 1 Then oRsTmp.Close
2950                                   Set oRsTmp = mo_ReglasAdmision.BuscaOtrosCPTyExamenesparaFormatoHIS(rsReporte.Fields!idCuentaAtencion, 1)
                                       If lbAgregaOrdenesMedicas = False Then
                                          oRsTmp.Filter = lcFilter
                                       End If
2960                              End If
2970                              mo_ReglasAdmision.CreaTemporalYcargaDxYcptXgrupo oRsDxCpt, rsDx1, oRsTmp, _
                                                                                  rsReporte!idAtencion, False, lcAD040
2980                              lnIndiceDxCPT = oRsDxCpt.RecordCount
2990                          End If
                              '
3000                          lbContinuar = True
3010                          If oRsDxCpt.RecordCount = 0 Then
3020                              lbContinuar = False
3030                          End If
                              '
3040                          If lbContinuar = True Then
3050                              lcDistritoDomicilio = Space(6)
3060                              If Not IsNull(rsReporte.Fields!IdDistrito) Then
3070                                 lcDistritoDomicilio = Right("0" & Trim(str(rsReporte.Fields!IdDistrito)), 6)
3080                              End If
                                  '
3090                              lnTot_reg = lnTot_reg + 1
3100                              lnReg = lnReg + 1
3110                              lnEdadEnAtencion = rsReporte.Fields!Edad
3120                              If lnEdadEnAtencion > 99 Then
3130                                 lnEdadEnAtencion = 99
3140                              End If
3150                              lnGrupo6DxCPT = IIf(Round(oRsDxCpt.RecordCount / 6) < oRsDxCpt.RecordCount / 6, Round(oRsDxCpt.RecordCount / 6) + 1, Round(oRsDxCpt.RecordCount / 6))
3160                              lnRestoDxCPT = IIf(oRsDxCpt.RecordCount Mod 6 = 0, 6, oRsDxCpt.RecordCount Mod 6)

3170                              For lnContador = 0 To lnGrupo6DxCPT - 1
3180                                  lnIndiceDxCPT = lnContador * 6
3190                                  If lbExportaAnovafis = True Then
                                          '***********exporta al NOVAFIS ****************
3200                                      lnEdadEnTablaEdades = 0
3210                                      oRsNovafis_edades.Filter = "tipo_edad='" & UCase(rsReporte.Fields!EdadCodigo) & "' and " & _
                                                                     "edad=" & lnEdadEnAtencion
3220                                      If oRsNovafis_edades.RecordCount > 0 Then
3230                                         lnEdadEnTablaEdades = oRsNovafis_edades!ed
3240                                      End If
                                          '
3250                                      oRsFox1.AddNew
3260                                      oRsFox1.Fields!cod_estab = lcCod_2000
3270                                      oRsFox1.Fields!ano = Val(lcANIO)
3280                                      oRsFox1.Fields!Mes = Val(lcMes)
3290                                      oRsFox1.Fields!nom_lote = lcNomLote
3300                                      oRsFox1.Fields!Num_pag = lnNumPag
3310                                      oRsFox1.Fields!num_reg = lnReg
3320                                      oRsFox1.Fields!dia = Day(rsReporte.Fields!FechaIngreso)
3330                                      oRsFox1.Fields!fichafam = lcFichaOhistoria
3340                                      oRsFox1.Fields!Ubigeo = lcDistritoDomicilio
3350                                      oRsFox1.Fields!Edad = lnEdadEnAtencion
3360                                      oRsFox1.Fields!tip_edad = rsReporte.Fields!EdadCodigo
3370                                      oRsFox1.Fields!Sexo = IIf(rsReporte.Fields!idTipoSexo = 1, "M", "F")
3380                                      oRsFox1.Fields!establec = IIf(rsReporte.Fields!IdTipoCondicionALEstab = 1, "N", IIf(rsReporte.Fields!IdTipoCondicionALEstab = 2, "R", "C"))
3390                                      oRsFox1.Fields!Servicio = IIf(rsReporte.Fields!IdTipoCondicionAlServicio = 1, "N", IIf(rsReporte.Fields!IdTipoCondicionAlServicio = 2, "R", "C"))
                                          'Coloca por defecto los Dx vacios
3400                                      oRsFox1.Fields!dx1 = ""
3410                                      oRsFox1.Fields!lc1 = ""
3420                                      oRsFox1.Fields!td1 = ""
3430                                      oRsFox1.Fields!dx2 = ""
3440                                      oRsFox1.Fields!lc2 = ""
3450                                      oRsFox1.Fields!td2 = ""
3460                                      oRsFox1.Fields!dx3 = ""
3470                                      oRsFox1.Fields!lc3 = ""
3480                                      oRsFox1.Fields!td3 = ""
3490                                      oRsFox1.Fields!dx4 = ""
3500                                      oRsFox1.Fields!lc4 = ""
3510                                      oRsFox1.Fields!td4 = ""
3520                                      oRsFox1.Fields!dx5 = ""
3530                                      oRsFox1.Fields!lc5 = ""
3540                                      oRsFox1.Fields!td5 = ""
3550                                      oRsFox1.Fields!dx6 = ""
3560                                      oRsFox1.Fields!lc6 = ""
3570                                      oRsFox1.Fields!td6 = ""
3580                                      lcGestante = "0"
3590                                      For lnContDxCpt = 1 To 6
3600                                          oRsDxCpt.MoveFirst
3610                                          oRsDxCpt.Find "IdDxCpt=" & lnIndiceDxCPT + lnContDxCpt
3620                                          If Not oRsDxCpt.EOF Then
3630                                              Select Case lnContDxCpt
                                                  Case 1
3640                                                  oRsFox1.Fields!td1 = oRsDxCpt.Fields!TDx
3650                                                  oRsFox1.Fields!lc1 = oRsDxCpt.Fields!LDx
3660                                                  oRsFox1.Fields!dx1 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
3670                                                  If mo_ReglasComunes.novafis_ElDxEsDeEmbarazada(oRsDxCpt!dx, rsReporte.Fields!idTipoSexo) = True Then
3680                                                     lcGestante = "G"
3690                                                  End If
3700                                              Case 2
3710                                                  oRsFox1.Fields!td2 = oRsDxCpt.Fields!TDx
3720                                                  oRsFox1.Fields!lc2 = oRsDxCpt.Fields!LDx
3730                                                  oRsFox1.Fields!dx2 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
3740                                                  If mo_ReglasComunes.novafis_ElDxEsDeEmbarazada(oRsDxCpt!dx, rsReporte.Fields!idTipoSexo) = True Then
3750                                                     lcGestante = "G"
3760                                                  End If
3770                                              Case 3
3780                                                  oRsFox1.Fields!td3 = oRsDxCpt.Fields!TDx
3790                                                  oRsFox1.Fields!lc3 = oRsDxCpt.Fields!LDx
3800                                                  oRsFox1.Fields!dx3 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
3810                                                  If mo_ReglasComunes.novafis_ElDxEsDeEmbarazada(oRsDxCpt!dx, rsReporte.Fields!idTipoSexo) = True Then
3820                                                     lcGestante = "G"
3830                                                  End If
3840                                              Case 4
3850                                                  oRsFox1.Fields!td4 = oRsDxCpt.Fields!TDx
3860                                                  oRsFox1.Fields!lc4 = oRsDxCpt.Fields!LDx
3870                                                  oRsFox1.Fields!dx4 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
3880                                                  If mo_ReglasComunes.novafis_ElDxEsDeEmbarazada(oRsDxCpt!dx, rsReporte.Fields!idTipoSexo) = True Then
3890                                                     lcGestante = "G"
3900                                                  End If
3910                                              Case 5
3920                                                  oRsFox1.Fields!td5 = oRsDxCpt.Fields!TDx
3930                                                  oRsFox1.Fields!lc5 = oRsDxCpt.Fields!LDx
3940                                                  oRsFox1.Fields!dx5 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
3950                                                  If mo_ReglasComunes.novafis_ElDxEsDeEmbarazada(oRsDxCpt!dx, rsReporte.Fields!idTipoSexo) = True Then
3960                                                     lcGestante = "G"
3970                                                  End If
3980                                              Case 6
3990                                                  oRsFox1.Fields!td6 = oRsDxCpt.Fields!TDx
4000                                                  oRsFox1.Fields!lc6 = oRsDxCpt.Fields!LDx
4010                                                  oRsFox1.Fields!dx6 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
4020                                                  If mo_ReglasComunes.novafis_ElDxEsDeEmbarazada(oRsDxCpt!dx, rsReporte.Fields!idTipoSexo) = True Then
4030                                                     lcGestante = "G"
4040                                                  End If
4050                                              End Select
4060                                          End If
4070                                      Next
4080                                      oRsFox1.Fields!codpsal = lnDNImedico1
4090                                      oRsFox1.Fields!digita = Right("0" & lcCod_dig, 8)
4100                                      oRsFox1.Fields!cod_servsa = lcCod_servsa
4110                                      oRsFox1.Fields!mt = lcMt
4120                                      oRsFox1.Fields!cf = Left(lcFuenteFinanciamientoPaciente & " ", 2)
4130                                      oRsFox1.Fields!et = lcEtniaPaciente
4140                                      oRsFox1.Fields!pais = lcPaisPaciente
4150                                      oRsFox1.Fields!cod_doc = lcTipoDocumentoPaciente
4160                                      oRsFox1.Fields!familia = "00"
4170                                      oRsFox1.Fields!dni = lcDNIpaciente1
4180                                      oRsFox1.Fields!est = lcGestante
4190                                      oRsFox1.Fields!ed = lnEdadEnTablaEdades
4200                                      oRsFox1.Fields!ID = Space(30)
                                          oRsFox1.Fields!IdCuenta = rsReporte!idCuentaAtencion
4210                                      lnEdadEnTablaEdades = 0
4220                                  Else
                                          '***********exporta al HIS ****************
4230                                      oRsFox1.AddNew
4240                                      oRsFox1.Fields!cod_2000 = lcCod_2000
4250                                      oRsFox1.Fields!ano = Val(lcANIO)
4260                                      oRsFox1.Fields!Mes = Val(lcMes)
4270                                      oRsFox1.Fields!nom_lote = lcNomLote
4280                                      oRsFox1.Fields!Num_pag = lnNumPag
4290                                      oRsFox1.Fields!num_reg = lnReg
4300                                      oRsFox1.Fields!dia = Day(rsReporte.Fields!FechaIngreso)
4310                                      oRsFox1.Fields!fichafam = lcFichaOhistoria
4320                                      oRsFox1.Fields!Cod_Dpto = Left(lcDistritoDomicilio, 2)
4330                                      oRsFox1.Fields!Cod_Prov = Mid(lcDistritoDomicilio, 3, 2)
4340                                      oRsFox1.Fields!Cod_Dist = Right(lcDistritoDomicilio, 2)
4350                                      oRsFox1.Fields!Edad = lnEdadEnAtencion
4360                                      oRsFox1.Fields!tip_edad = rsReporte.Fields!EdadCodigo
4370                                      oRsFox1.Fields!Sexo = IIf(rsReporte.Fields!idTipoSexo = 1, "M", "F")
4380                                      oRsFox1.Fields!establec = IIf(rsReporte.Fields!IdTipoCondicionALEstab = 1, "N", IIf(rsReporte.Fields!IdTipoCondicionALEstab = 2, "R", "C"))
4390                                      oRsFox1.Fields!Servicio = IIf(rsReporte.Fields!IdTipoCondicionAlServicio = 1, "N", IIf(rsReporte.Fields!IdTipoCondicionAlServicio = 2, "R", "C"))
                                          
                                          'Coloca por defecto los Dx vacios
4400                                      oRsFox1.Fields!diagnost1 = ""
4410                                      oRsFox1.Fields!labconf1 = ""
4420                                      oRsFox1.Fields!codigo1 = ""
4430                                      oRsFox1.Fields!diagnost2 = ""
4440                                      oRsFox1.Fields!labconf2 = ""
4450                                      oRsFox1.Fields!codigo2 = ""
4460                                      oRsFox1.Fields!diagnost3 = ""
4470                                      oRsFox1.Fields!labconf3 = ""
4480                                      oRsFox1.Fields!codigo3 = ""
4490                                      oRsFox1.Fields!diagnost4 = ""
4500                                      oRsFox1.Fields!labconf4 = ""
4510                                      oRsFox1.Fields!codigo4 = ""
4520                                      oRsFox1.Fields!diagnost5 = ""
4530                                      oRsFox1.Fields!labconf5 = ""
4540                                      oRsFox1.Fields!codigo5 = ""
4550                                      oRsFox1.Fields!diagnost6 = ""
4560                                      oRsFox1.Fields!labconf6 = ""
4570                                      oRsFox1.Fields!codigo6 = ""
                                          
4580                                      For lnContDxCpt = 1 To 6
4590                                          oRsDxCpt.MoveFirst
4600                                          oRsDxCpt.Find "IdDxCpt=" & lnIndiceDxCPT + lnContDxCpt
4610                                          If Not oRsDxCpt.EOF Then
4620                                              Select Case lnContDxCpt
                                                  Case 1
4630                                                  oRsFox1.Fields!diagnost1 = oRsDxCpt.Fields!TDx
4640                                                  oRsFox1.Fields!labconf1 = oRsDxCpt.Fields!LDx
                                                      If sighentidades.Lx_LabVacio = oRsDxCpt.Fields!dx Then
                                                         oRsFox1.Fields!codigo1 = " "
                                                      Else
4650                                                     oRsFox1.Fields!codigo1 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
                                                      End If
4660                                              Case 2
4670                                                  oRsFox1.Fields!diagnost2 = oRsDxCpt.Fields!TDx
4680                                                  oRsFox1.Fields!labconf2 = oRsDxCpt.Fields!LDx
                                                      If sighentidades.Lx_LabVacio = oRsDxCpt.Fields!dx Then
                                                         oRsFox1.Fields!codigo2 = " "
                                                      Else
4690                                                     oRsFox1.Fields!codigo2 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
                                                      End If
4700                                              Case 3
4710                                                  oRsFox1.Fields!diagnost3 = oRsDxCpt.Fields!TDx
4720                                                  oRsFox1.Fields!labconf3 = oRsDxCpt.Fields!LDx
                                                      If sighentidades.Lx_LabVacio = oRsDxCpt.Fields!dx Then
                                                         oRsFox1.Fields!codigo3 = " "
                                                      Else
4730                                                     oRsFox1.Fields!codigo3 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
                                                      End If
4740                                              Case 4
4750                                                  oRsFox1.Fields!diagnost4 = oRsDxCpt.Fields!TDx
4760                                                  oRsFox1.Fields!labconf4 = oRsDxCpt.Fields!LDx
                                                      If sighentidades.Lx_LabVacio = oRsDxCpt.Fields!dx Then
                                                         oRsFox1.Fields!codigo4 = " "
                                                      Else
                                                         oRsFox1.Fields!codigo4 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
                                                      End If
4780                                              Case 5
4790                                                  oRsFox1.Fields!diagnost5 = oRsDxCpt.Fields!TDx
4800                                                  oRsFox1.Fields!labconf5 = oRsDxCpt.Fields!LDx
                                                      If sighentidades.Lx_LabVacio = oRsDxCpt.Fields!dx Then
                                                         oRsFox1.Fields!codigo5 = " "
                                                      Else
                                                         oRsFox1.Fields!codigo5 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
                                                      End If
4820                                              Case 6
4830                                                  oRsFox1.Fields!diagnost6 = oRsDxCpt.Fields!TDx
4840                                                  oRsFox1.Fields!labconf6 = oRsDxCpt.Fields!LDx
                                                      If sighentidades.Lx_LabVacio = oRsDxCpt.Fields!dx Then
                                                         oRsFox1.Fields!codigo6 = " "
                                                      Else
                                                         oRsFox1.Fields!codigo6 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
                                                      End If
4860                                              End Select
4870                                          End If
4880                                      Next
4890                                      oRsFox1.Fields!esta_reg = "2"
4900                                      oRsFox1.Fields!mt = lcMt
4910                                      oRsFox1.Fields!dni = lcDNIpaciente
4920                                      oRsFox1.Fields!FI = lcFuenteFinanciamientoPaciente
4930                                      oRsFox1.Fields!et = lcEtniaPaciente
4940                                      oRsFox1.Fields!st = "1"
                                          oRsFox1.Fields!IdCuenta = rsReporte!idCuentaAtencion
4950                                      oRsFox1.Update
4960                                  End If
4970                              Next
                                  '
4980                          End If
4990                          lnTotal = lnTotal + 1
                              oRsFox5.Open "update parametros set valorint=" & Trim(str(lnTotal)) & " where idparametro=2", oConexion, adOpenKeyset, adLockOptimistic
                              ml_ValorActual = Trim(str(lnTotal))
5000                          'mo_ProgressRpt1.Value = lnTotal
5010                          'DoEvents
If lnTotal = 1719 Then
ml_ValorActual = 0
End If
                              
5020                          rsReporte.MoveNext
5030                          If rsReporte.EOF Then
5040                             Exit Do
5050                          End If
5060                          If lnReg >= lnMaxPacientesXhoja And lnIdEmpleado = rsReporte.Fields!IdEmpleado And lcCodigoServicioHIS = rsReporte.Fields!codigoServicioHIS And ldFecha = rsReporte.Fields!FechaIngreso Then
5070                             lbAumentoPagina = True
5080                             lnReg = 0
5090                             Exit Do
5100                          Else
5110                             lbAumentoPagina = False
5120                          End If
5130                      Loop
5140                      If lnTot_reg > 0 Then
5150                          If lbExportaAnovafis = True Then
                                  '***********exporta al NOVAFIS ****************
5160                              lbNuevoRegistro = True
5170                              If oRsFox2.RecordCount > 0 Then
5180                                 oRsFox2.MoveFirst
5190                                 Do While Not oRsFox2.EOF
5200                                    If oRsFox2!cod_estab = lcCod_2000 And oRsFox2!ano = Val(lcANIO) And _
                                                          oRsFox2!Mes = Val(lcMes) And oRsFox2!nom_lote = lcNomLote Then
5210                                       lbNuevoRegistro = False
5220                                       Exit Do
5230                                    End If
5240                                    oRsFox2.MoveNext
5250                                 Loop
5260                              End If
5270                              If lbNuevoRegistro = True Then
5280                                  oRsFox2.AddNew
5290                                  oRsFox2.Fields!cod_estab = lcCod_2000
5300                                  oRsFox2.Fields!ano = Val(lcANIO)
5310                                  oRsFox2.Fields!Mes = Val(lcMes)
5320                                  oRsFox2.Fields!nom_lote = lcNomLote
5330                                  oRsFox2.Fields!tot_pag = lnNumPag
5340                              Else
5350                                  oRsFox2.Fields!tot_pag = oRsFox2.Fields!tot_pag + lnNumPag
5360                              End If
5370                          Else
                                  '***********exporta al HIS ****************
5380                              oRsFox2.AddNew
5390                              oRsFox2.Fields!cod_2000 = lcCod_2000
5400                              oRsFox2.Fields!ano = Val(lcANIO)
5410                              oRsFox2.Fields!Mes = Val(lcMes)
5420                              oRsFox2.Fields!nom_lote = lcNomLote
5430                              oRsFox2.Fields!Num_pag = lnNumPag
5440                              oRsFox2.Fields!Codif = lcCodif
5450                              oRsFox2.Fields!cod_servsa = lcCod_servsa
5460                              oRsFox2.Fields!plaza = lcCodif
5470                              oRsFox2.Fields!Esta_pag = "2"
5480                              oRsFox2.Fields!Tot_reg = lnTot_reg
5490                              oRsFox2.Fields!FlagEnvio = Space(1)
5500                              oRsFox2.Fields!mt = lcMt
5510                              oRsFox2.Fields!st = "1"
5520                              oRsFox2.Update
5530                          End If
                              '
5540                          lnTTot_pag = lnTTot_pag + lnNumPag
5550                          lnTTot_reg = lnTTot_reg + lnTot_reg
                              '
5560                          oRsFox3.AddNew
5570                          oRsFox3.Fields!cod_2000 = lcCod_2000
5580                          oRsFox3.Fields!ano = Val(lcANIO)
5590                          oRsFox3.Fields!Mes = Val(lcMes)
5600                          oRsFox3.Fields!nom_lote = lcNomLote
5610                          If lbExportaAnovafis = True Then
                                  '***********exporta al NOVAFIS ****************
5620                              oRsFox3.Fields!cod_dig = "1"
5630                          Else
                                  '***********exporta al HIS ****************
5640                              oRsFox3.Fields!cod_dig = lcCod_dig
5650                          End If
5660                          oRsFox3.Fields!esta_lote = "LC"
5670                          oRsFox3.Fields!tot_pag = lnTTot_pag
5680                          oRsFox3.Fields!Tot_reg = lnTTot_reg
5690                          oRsFox3.Fields!Tot_rmalos = 0
5700                          oRsFox3.Fields!tot_pgcarg = 0
5710                          oRsFox3.Fields!esta_local = Space(1)
5720                          oRsFox3.Fields!nomfile = Space(14)
5730                          oRsFox3.Fields!flagReport = "3"
5740                          oRsFox3.Fields!mt = lcMt
5750                          oRsFox3.Update
5760                      End If
5770                      If rsReporte.EOF Then
5780                         Exit Do
5790                      End If
5800                  Loop
5810             Loop
                 '
5820             If lbExportaAnovafis = True Then
                     '***********exporta al NOVAFIS ****************
5830                 lnTTot_pag = oRsFox1.RecordCount
5840                 If lnTTot_pag > 0 Then
5850                      oRsFox1.MoveFirst
5860                      Do While Not oRsFox1.EOF
5870                         lnNumPag = 0
5880                         oRsFox2.MoveFirst
5890                         Do While Not oRsFox2.EOF
5900                              If oRsFox2!cod_estab = oRsFox1!cod_estab And oRsFox2!ano = oRsFox1!ano And _
                                                    oRsFox2!Mes = oRsFox1!Mes And oRsFox2!nom_lote = oRsFox1!nom_lote Then
5910                                 lnNumPag = oRsFox2!tot_pag
5920                                 Exit Do
5930                              End If
5940                              oRsFox2.MoveNext
5950                         Loop
                             
5960                         lcMt = oRsFox1!cod_estab & Trim(str(oRsFox1!ano)) & Right("0" & Trim(str(oRsFox1!Mes)), 2) & _
                                   oRsFox1!nom_lote & Right("00" & Trim(str(lnNumPag)), 3) & Right("0" & oRsFox1!Num_pag, 2) & _
                                   Right("0" & Trim(str(oRsFox1!num_reg)), 2)
5970                         oRsFox1!ID = lcMt
5980                         oRsFox1.Update
5990                         oRsFox1.MoveNext
6000                      Loop
                          '
6010                      If ms_MensajeError = "" Then
6020                          GeneraTXT_una_linea oRsFox1, lcRutaExportar & "envio.txt", lnTTot_pag, 1, Chr(9), Chr(10)
6030                          Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "A", 0, "Atenciones", oConexion, 500 + 141, mo_lcNombrePc, "Exportar al NOVAFIS: " & lcFec1 & " al " & lcFec2)     '500+ ListBarReporte.idReporte
6040                          MsgBox "Se exportó correctamente el archivo: " & lcRutaExportar & "envio.txt"
6050                      End If
6060                 End If
6070             Else
                     '***********exporta al HIS ****************
6080                 Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "A", 0, "Atenciones", oConexion, 500 + 141, mo_lcNombrePc, "Exportar al HIS: " & lcFec1 & " al " & lcFec2)     '500+ ListBarReporte.idReporte
6090             End If
                 '
                 If lbExportaAlHisMinsa = True Then
                    ExportarDatosAlHisMINSA oRsFox1, lnIdMesElegido, lcANIO, rsReporte, oConexion, oConexionExterna
                 End If
6100          End If
6110          If oRsTmp.State = 1 Then oRsTmp.Close
6120          oRsFox1.Close
6130          oRsFox2.Close
6140          oRsFox3.Close
6150          oConexionFox.Close
6160          oConexion.Close
6170          oConexionExterna.Close
              If lbExportaAnovafis = True Then
6180             oRsNovafis_edades.Close
              End If
6190          If ms_MensajeError <> "" Then
6200             MsgBox ms_MensajeError, vbInformation, "Mensajes"
6210          End If
              
6220          Set mo_ReglasFarmacia = Nothing
6230          Set mo_ReglasComunes = Nothing
6240          Set mo_ReglasSeguridad = Nothing
6250          Set mo_ReglasAdmision = Nothing
6260          Set mo_ReglasServGeograf = Nothing
6270          Set oConexion = Nothing
6280          Set oConexionFox = Nothing
6290          Set rsReporte = Nothing
6300          Set oRsTmp = Nothing
6310          Set oRsFox1 = Nothing
6320          Set oRsFox2 = Nothing
6330          Set oRsFox3 = Nothing
6340          Set oRsFox4 = Nothing
6350          Set oRsEdadHis = Nothing
6360          Set oRsImpHIStodos = Nothing
6370          Set rsDx1 = Nothing
6380          Set oConexionExterna = Nothing
6390          Set oRsNovafis_edades = Nothing
6400      Exit Sub
          
ErrorProceso:
6410      If Err.Number = -2147217900 Or Err.Number = -2147217865 Then
6420         Resume Next
6430      Else
6440         MsgBox Err.Description & sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                                                             "Sub ExportaDAtosAlHISv4", "procesos.cls")  'debb-02/05/2016
6450         Resume
6460      End If
              
End Sub

Sub ExportarDatosAlHisMINSA(oRsFox1 As Recordset, lnMes As Integer, lcANIO As String, rsReporte As Recordset, _
                            oConexion As Connection, oConexionExterna As Connection)
    Dim lcMensajeLicencia As String, lbTieneLicencia As Boolean, lnTotalRegistros As Long
    lnTotalRegistros = oRsFox1.RecordCount
    lbTieneLicencia = True
    If lbTieneLicencia = False Or lnTotalRegistros = 0 Then
       Exit Sub
    End If
    
    On Error GoTo ErrJson
    Dim p As Object
    Dim sInputJson As String
    Dim fso
    Dim act
    Dim lnFor As Long
    Dim lcArchivoJSON As String, lbUltimoRegistro As Boolean
    Dim lnIdCuenta As Long, ldHoy As Date, lcSisAfiliacion As String, lcFechaEmision As String
    Dim lcHoraEmision As String, lcCitaFecha As String, lcCitaHora  As String, lcItems As String
    Dim lcErroresHallados As String, lcLabValor As String, lcLabCodigo As String, lcItemTipoDiagnostico As String
    Dim lcItemCodigo As String, lcItemTipoItem As String, lcPeso As String, lcTalla As String, lcHemoglobina As String
    Dim oEdad As Edad, ldFechaHemoglobina As Date, lcCitaParte1 As String, lcCitaParte2 As String
    Dim lcCitaParte3 As String, lcUsuario As String, lcMedico As String, lcPaciente As String, lcSaltoDeLinea As String
    Dim lcCabecera As String, lcIdCondicion As String, lcCita As String, lcErroresGraves As String, lbContinua As Boolean
    Dim lcRenaesReferenciaOrigen As String, lcPaisProcedencia As String, lcPacDocumento As String
    Dim lcPacTipoDoc As String, lcIdProfesion As String, lcFechaNacimiento As String
    Const lcxItems As String = "items: ["
    Dim oRsItems As New Recordset
    Dim oRsLab As New Recordset
    Dim oRsCondicionTrabajo As New Recordset
    Dim oRsTmp_11 As New Recordset
    Dim oRsProfesion As New Recordset
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
    
    
    lcErroresHallados = ""
    lcErroresGraves = ""
    Set oRsProfesion = mo_CabeceraReportes.TiposEmpleadoSeleccionarTodos
    Set oRsCondicionTrabajo = mo_CabeceraReportes.TiposCondicionTrabajoSeleccionarTodos
    Set oRsLab = mo_CabeceraReportes.DevuelveHIS_SITUACIOporDescripcion()
    ldHoy = CDate(lcBuscaParametro.RetornaFechaHoraServidorSQL)
    lcSaltoDeLinea = IIf(sighentidades.Parametro573valorInt = "1", vbCrLf, "")
    lcArchivoJSON = sighentidades.DevuelveRutaConSlashInvertida(lcBuscaParametro.SeleccionaFilaParametro(313)) & "JSON" & _
                   lcANIO & sighentidades.DevuelveNombreMes(lnMes)
    Set fso = CreateObject("scripting.filesystemobject")
    Set act = fso.CreateTextFile(lcArchivoJSON, True)
    act.Write ("[")
    lbUltimoRegistro = False
    lnFor = 1
    oRsFox1.Sort = "idCuenta"
    oRsFox1.MoveFirst
    Do While Not oRsFox1.EOF
          lnIdCuenta = oRsFox1!IdCuenta
          rsReporte.Filter = "idCuentaAtencion=" & lnIdCuenta
          lcSisAfiliacion = IIf(IsNull(rsReporte!sisAfiliacion), "", Trim(rsReporte!sisAfiliacion))
          lcFechaEmision = Trim(str(Year(ldHoy))) & Right("0" & Trim(str(Month(ldHoy))), 2) & Right("0" & Trim(str(Day(ldHoy))), 2)
          lcHoraEmision = Right("00" & str(Hour(ldHoy)), 2) & Right("00" & Trim(str(Minute(ldHoy))), 2) & "00"
          lcCitaFecha = Trim(str(oRsFox1!ano)) & Right("0" & Trim(str(oRsFox1!Mes)), 2) & Right("0" & Trim(str(oRsFox1!dia)), 2)
          lcCitaHora = Left(rsReporte!HoraIngreso, 2) & Right(rsReporte!HoraIngreso, 2) & "00"
          '************* procesa CABECERA
          lcCabecera = "fechaemision: " & lcFechaEmision & ", " & _
                       "horaemision: " & lcHoraEmision & ", " & _
                       "idenvio: '" & lcFechaEmision & lcHoraEmision & Right(oRsFox1!cod_estab, 5) & Mid(oRsFox1!dni, 4, 1) & Mid(oRsFox1!dni, 5, 8) & "', " & _
                       "idinstitucion: '', " & _
                       "tipooperacion: '01'" & lcSaltoDeLinea
          '************procesa CITA parte uno
          lcCitaParte1 = "numeroafiliacion: '" & lcSisAfiliacion & "', " & _
                         "fechaatencion: " & lcCitaFecha & ", " & _
                         "estadoregistro: 'A'" & lcSaltoDeLinea
          '************procesa CITA tercera parte (despues de Items dx, cpt, labs)
          lcPeso = "": lcTalla = ""
          Set oRsItems = mo_ReglasAdmision.AtencionesCEXSeleccionarPorId(rsReporte!idAtencion, oConexionExterna)
          If oRsItems.RecordCount > 0 Then
           If Not IsNull(oRsItems!triajePeso) Then
              lcPeso = oRsItems!triajePeso
           End If
           If Not IsNull(oRsItems!triajeTalla) Then
              lcTalla = Format(CCur(oRsItems!triajeTalla) / 100, "#0.0000")
           End If
          End If
          oRsItems.Close
          mo_ReglasLaboratorio.DevuelveHemoglobinaDeLaAtencion ldFechaHemoglobina, lcHemoglobina, rsReporte!idPaciente, _
                             rsReporte!idAtencion, rsReporte!FechaIngreso, oConexion
          oEdad = sighentidades.CalcularEdad(rsReporte!FechaNacimiento, rsReporte!FechaIngreso)
          lcCitaParte3 = "idups: '" & oRsFox1!cod_servsa & "', " & _
                         "idestablecimiento: '" & Right(oRsFox1!cod_estab, 5) & "', " & _
                         "diaedad: " & oEdad.EdadDia & ", " & _
                         "edadregistro: " & oEdad.EdadAnio & ", " & _
                         "idturno: '" & oRsFox1!mt & "', " & _
                         "idtipoedadregistro: 'A', " & _
                         "fgdiag: 11, " & _
                         "mesedad: " & oEdad.EdadMes & ", " & _
                         "componente: '', " & _
                         "idfinanciador: '" & oRsFox1!cf & ", " & _
                         "annioedad: " & oEdad.EdadAnio & ", " & _
                         "examenfisico: {" & _
                                        "peso: '" & lcPeso & "', " & _
                                        "talla: '" & lcTalla & "', " & _
                                        "hemoglobina: '" & lcHemoglobina & "'" & _
                                        "}" & lcSaltoDeLinea
              
          '************procesa USUARIO
          lbContinua = True
          Set oRsTmp_11 = mo_CabeceraReportes.EmpleadosSeleccionarPorIdEmpleado(sighentidades.Usuario, oConexion)
          lcIdCondicion = "09"
          oRsCondicionTrabajo.Filter = "idCondicionTrabajo=" & oRsTmp_11!idCondicionTrabajo
          If oRsCondicionTrabajo.RecordCount > 0 Then
             lcIdCondicion = oRsCondicionTrabajo!condicionTrabajoSEM
          End If
          oRsProfesion.Filter = "idTipoEmpleado=" & oRsTmp_11!idTipoEmpleado
          If oRsProfesion.RecordCount > 0 Then
             lcIdProfesion = oRsProfesion!tipoEmpleadoHIS
          Else
             lbContinua = False
             lcErroresGraves = lcErroresGraves & "Al EMPLEADO: " & oRsTmp_11!ApellidoPaterno & " " & _
                               oRsTmp_11!ApellidoMaterno & oRsTmp_11!Nombres & " no le ha elejido la PROFESION" & Chr(13)
          End If
          If IsNull(oRsTmp_11!idTipoSexo) Then
             lbContinua = False
             lcErroresGraves = lcErroresGraves & "Al EMPLEADO: " & oRsTmp_11!ApellidoPaterno & " " & _
                               oRsTmp_11!ApellidoMaterno & oRsTmp_11!Nombres & " no le ha elejido el SEXO" & Chr(13)
          End If
          If Not IsDate(oRsTmp_11!FechaNacimiento) Then
             lbContinua = False
             lcErroresGraves = lcErroresGraves & "Al EMPLEADO: " & oRsTmp_11!ApellidoPaterno & " " & _
                               oRsTmp_11!ApellidoMaterno & oRsTmp_11!Nombres & " no le ha registrado la FECHA DE NACIMIENTO" & Chr(13)
          Else
             lcFechaNacimiento = Format(oRsTmp_11!FechaNacimiento, "yyyymmdd")
          End If
          If lbContinua Then
             lcUsuario = "personal_registra: {" & _
                                           "nrodocumento: '" & Trim(oRsTmp_11!dni) & "', " & _
                                           "apematerno: '" & Trim(oRsTmp_11!ApellidoMaterno) & "', " & _
                                           "idpais: '51', " & _
                                           "idprofesion: '" & lcIdProfesion & "', " & _
                                           "fechanacimiento: " & lcFechaNacimiento & ", " & _
                                           "nombres: '" & Trim(oRsTmp_11!Nombres) & "', " & _
                                           "idtipodoc: '1', " & _
                                           "apepaterno: '" & Trim(oRsTmp_11!ApellidoPaterno) & "', " & _
                                           "idsexo: '" & IIf(oRsTmp_11!idTipoSexo = 1, "M", "F") & "', " & _
                                           "idcondicion: '" & lcIdCondicion & "'" & _
                                           "}" & lcSaltoDeLinea
          Else
              Exit Do
          End If
          oRsTmp_11.Close
          '************procesa MEDICO
          lbContinua = True
          Set oRsTmp_11 = mo_CabeceraReportes.EmpleadosSeleccionarPorIdEmpleado(rsReporte!IdEmpleado, oConexion)
          lcIdCondicion = "09"
          oRsCondicionTrabajo.Filter = "idCondicionTrabajo=" & oRsTmp_11!idCondicionTrabajo
          If oRsCondicionTrabajo.RecordCount > 0 Then
             lcIdCondicion = oRsCondicionTrabajo!condicionTrabajoSEM
          End If
          oRsProfesion.Filter = "idTipoEmpleado=" & oRsTmp_11!idTipoEmpleado
          If oRsProfesion.RecordCount > 0 Then
             lcIdProfesion = oRsProfesion!tipoEmpleadoHIS
          Else
             lbContinua = False
             lcErroresGraves = lcErroresGraves & "Al EMPLEADO: " & oRsTmp_11!ApellidoPaterno & " " & _
                               oRsTmp_11!ApellidoMaterno & oRsTmp_11!Nombres & " no le ha elejido la PROFESION" & Chr(13)
          End If
          If IsNull(oRsTmp_11!idTipoSexo) Then
             lbContinua = False
             lcErroresGraves = lcErroresGraves & "Al EMPLEADO: " & oRsTmp_11!ApellidoPaterno & " " & _
                               oRsTmp_11!ApellidoMaterno & oRsTmp_11!Nombres & " no le ha elejido el SEXO" & Chr(13)
          End If
          If Not IsDate(oRsTmp_11!FechaNacimiento) Then
             lbContinua = False
             lcErroresGraves = lcErroresGraves & "Al EMPLEADO: " & oRsTmp_11!ApellidoPaterno & " " & _
                               oRsTmp_11!ApellidoMaterno & oRsTmp_11!Nombres & " no le ha registrado la FECHA DE NACIMIENTO" & Chr(13)
          Else
             lcFechaNacimiento = Format(oRsTmp_11!FechaNacimiento, "yyyymmdd")
          End If
          If lbContinua Then
             lcMedico = "personal_atiende: {" & _
                                           "nrodocumento: '" & Trim(oRsTmp_11!dni) & "', " & _
                                           "apematerno: '" & Trim(oRsTmp_11!ApellidoMaterno) & "', " & _
                                           "idpais: '51', " & _
                                           "idprofesion: '" & lcIdProfesion & "', " & _
                                           "fechanacimiento: " & lcFechaNacimiento & ", " & _
                                           "nombres: '" & Trim(oRsTmp_11!Nombres) & "', " & _
                                           "idtipodoc: '1', " & _
                                           "apepaterno: '" & Trim(oRsTmp_11!ApellidoPaterno) & "', " & _
                                           "idsexo: '" & IIf(oRsTmp_11!idTipoSexo = 1, "M", "F") & "', " & _
                                           "idcondicion: '" & lcIdCondicion & "'" & _
                                           "}" & lcSaltoDeLinea
          Else
              Exit Do
          End If
          oRsTmp_11.Close
          '************procesa PACIENTE
          lcRenaesReferenciaOrigen = ""
          Set oRsTmp_11 = mo_ReglasAdmision.AtencionesDatosAdicionalesSeleccionarXidAtencion(rsReporte!idAtencion, oConexion)
          If oRsTmp_11.RecordCount > 0 Then
             If Not IsNull(oRsTmp_11!codigoOrigen) Then
                lcRenaesReferenciaOrigen = oRsTmp_11!codigOrigen
             End If
          End If
          oRsTmp_11.Close
          lcPaisProcedencia = "51"
          If Not IsNull(rsReporte!idPaisProcedencia) Then
             lcPaisProcedencia = mo_CabeceraReportes.PaisCodigoPostalPorIdentificador(rsReporte!idPaisProcedencia, oConexion)
             If lcPaisProcedencia = "" Then
                lcPaisProcedencia = "51"
             End If
          End If
          
          lcPacDocumento = ""
          lcPacTipoDoc = ""
          If Not IsNull(rsReporte!nroDocumento) Then
             lcPacDocumento = rsReporte!nroDocumento
             lcPacTipoDoc = rsReporte!HIStipoDocumento
          End If
          lcFechaNacimiento = Format(rsReporte!FechaNacimiento, "yyyymmdd")
          lcPaciente = "paciente: {" & _
                                           "nrodocumento: '" & lcPacDocumento & "', " & _
                                           "apematerno: '" & Trim(rsReporte!ApellidoMaterno) & "', " & _
                                           "idflag: 5, " & _
                                           "nombres: '" & Trim(rsReporte!PrimerNombre) & _
                                                        IIf(IsNull(rsReporte!SegundoNombre), "", _
                                                        " " & Trim(rsReporte!SegundoNombre)) & "', " & _
                                           "nrohistoriaclinica: '" & Trim(str(rsReporte!NroHistoriaClinica)) & "', " & _
                                           "idtipodoc: '" & lcPacTipoDoc & "', " & _
                                           "apepaterno: '" & Trim(rsReporte!ApellidoPaterno) & "', " & _
                                           "idetnia: '" & rsReporte!IdEtnia & "', " & _
                                           "fechanacimiento: " & lcFechaNacimiento & ", " & _
                                           "idestablecimiento: " & lcRenaesReferenciaOrigen & ", " & _
                                           "idpais: '" & lcPaisProcedencia & "'," & _
                                           "idsexo: '" & IIf(rsReporte!idTipoSexo = 1, "M", "F") & "'" & _
                                           "}" & lcSaltoDeLinea
          '************procesa CITA parte dos: procesa items de Dx, Cpt, Labs
          lcCitaParte2 = lcxItems
          Do While Not oRsFox1.EOF And lnIdCuenta = oRsFox1!IdCuenta
             'uno
             If Not IsNull(oRsFox1!dx1) Then
                lcLabValor = "***"
                lcLabCodigo = "***"
                If Not IsNull(oRsFox1!lc1) Then
                   lcLabValor = Space(3)
                   lcLabCodigo = "47"
                   oRsLab.Filter = "valores='" & oRsFox1!lc1 & "'"
                   If oRsLab.RecordCount = 0 Then
                       lcErroresHallados = lcErroresHallados & "La CUENTA " & Trim(str(lnIdCuenta)) & " Tiene LAB: " & _
                                           oRsFox1!lc1 & " que no está en la TABLA HIS_situacio" & Chr(13)
                   Else
                       lcLabValor = oRsLab!valores
                       lcLabCodigo = Trim(str(oRsLab!codigo))
                   End If
                End If
                lcItemTipoDiagnostico = oRsFox1!td1
                lcItemCodigo = oRsFox1!dx1
                lcItemTipoItem = "EX"
                Set oRsItems = mo_CabeceraReportes.DiagnosticosSeleccionarXcodigoCIEsinPto(oRsFox1!dx1, oConexion)
                If oRsItems.RecordCount > 0 Then
                   lcItemTipoItem = "CX"
                Else
                   oRsItems.Close
                   Set oRsItems = mo_CabeceraReportes.ProcedimientosSeleccionarPorCodigo(oRsFox1!dx1, oConexion)
                   If oRsItems.RecordCount > 0 Then
                      If oRsItems!IdServicioSubGrupo = 3 Then
                         lcItemTipoItem = "PL"
                      Else
                         lcItemTipoItem = "CP"
                      End If
                   End If
                End If
                oRsItems.Close
                lcCitaParte2 = lcCitaParte2 & "{"
                If lcLabValor <> "***" Then
                   lcCitaParte2 = lcCitaParte2 & "labs: [{codigo: " & lcLabCodigo & ", " & _
                                                          "valor: '" & lcLabValor & "'}], " & lcSaltoDeLinea
                Else
                   lcCitaParte2 = lcCitaParte2 & "labs: '', " & lcSaltoDeLinea
                End If
                lcCitaParte2 = lcCitaParte2 & "tipodiagnostico: '" & lcItemTipoDiagnostico & "', " & _
                                              "codigo: '" & lcItemCodigo & "', " & _
                                              "tipoitem: '" & lcItemTipoItem & "'"
                lcCitaParte2 = lcCitaParte2 & "}" & lcSaltoDeLinea
             End If
             'dos
             If Not IsNull(oRsFox1!dx2) Then
                lcLabValor = "***"
                lcLabCodigo = "***"
                If Not IsNull(oRsFox1!lc2) Then
                   lcLabValor = Space(3)
                   lcLabCodigo = "47"
                   oRsLab.Filter = "valores='" & oRsFox1!lc2 & "'"
                   If oRsLab.RecordCount = 0 Then
                       lcErroresHallados = lcErroresHallados & "La CUENTA " & Trim(str(lnIdCuenta)) & " Tiene LAB: " & _
                                           oRsFox1!lc2 & " que no está en la TABLA HIS_situacio" & Chr(13)
                   Else
                       lcLabValor = oRsLab!valores
                       lcLabCodigo = Trim(str(oRsLab!codigo))
                   End If
                End If
                lcItemTipoDiagnostico = oRsFox1!td2
                lcItemCodigo = oRsFox1!dx2
                lcItemTipoItem = "EX"
                Set oRsItems = mo_CabeceraReportes.DiagnosticosSeleccionarXcodigoCIEsinPto(oRsFox1!dx2, oConexion)
                If oRsItems.RecordCount > 0 Then
                   lcItemTipoItem = "CX"
                Else
                   oRsItems.Close
                   Set oRsItems = mo_CabeceraReportes.ProcedimientosSeleccionarPorCodigo(oRsFox1!dx2, oConexion)
                   If oRsItems.RecordCount > 0 Then
                      If oRsItems!IdServicioSubGrupo = 3 Then
                         lcItemTipoItem = "PL"
                      Else
                         lcItemTipoItem = "CP"
                      End If
                   End If
                End If
                oRsItems.Close
                If lcCitaParte2 = lcxItems Then
                   lcCitaParte2 = lcCitaParte2 & "{"
                Else
                   lcCitaParte2 = lcCitaParte2 & ",{"
                End If
                If lcLabValor <> "***" Then
                   lcCitaParte2 = lcCitaParte2 & "labs: [{codigo: " & lcLabCodigo & ", " & _
                                                          "valor: '" & lcLabValor & "'}], " & lcSaltoDeLinea
                Else
                   lcCitaParte2 = lcCitaParte2 & "labs: '', " & lcSaltoDeLinea
                End If
                lcCitaParte2 = lcCitaParte2 & "tipodiagnostico: '" & lcItemTipoDiagnostico & "', " & _
                                              "codigo: '" & lcItemCodigo & "', " & _
                                              "tipoitem: '" & lcItemTipoItem & "'"
                lcCitaParte2 = lcCitaParte2 & "}" & lcSaltoDeLinea
             End If
             'tres
             If Not IsNull(oRsFox1!dx3) Then
                lcLabValor = "***"
                lcLabCodigo = "***"
                If Not IsNull(oRsFox1!lc3) Then
                   lcLabValor = Space(3)
                   lcLabCodigo = "47"
                   oRsLab.Filter = "valores='" & oRsFox1!lc3 & "'"
                   If oRsLab.RecordCount = 0 Then
                       lcErroresHallados = lcErroresHallados & "La CUENTA " & Trim(str(lnIdCuenta)) & " Tiene LAB: " & _
                                           oRsFox1!lc3 & " que no está en la TABLA HIS_situacio" & Chr(13)
                   Else
                       lcLabValor = oRsLab!valores
                       lcLabCodigo = Trim(str(oRsLab!codigo))
                   End If
                End If
                lcItemTipoDiagnostico = oRsFox1!td3
                lcItemCodigo = oRsFox1!dx3
                lcItemTipoItem = "EX"
                Set oRsItems = mo_CabeceraReportes.DiagnosticosSeleccionarXcodigoCIEsinPto(oRsFox1!dx3, oConexion)
                If oRsItems.RecordCount > 0 Then
                   lcItemTipoItem = "CX"
                Else
                   oRsItems.Close
                   Set oRsItems = mo_CabeceraReportes.ProcedimientosSeleccionarPorCodigo(oRsFox1!dx3, oConexion)
                   If oRsItems.RecordCount > 0 Then
                      If oRsItems!IdServicioSubGrupo = 3 Then
                         lcItemTipoItem = "PL"
                      Else
                         lcItemTipoItem = "CP"
                      End If
                   End If
                End If
                oRsItems.Close
                If lcCitaParte2 = lcxItems Then
                   lcCitaParte2 = lcCitaParte2 & "{"
                Else
                   lcCitaParte2 = lcCitaParte2 & ",{"
                End If
                If lcLabValor <> "***" Then
                   lcCitaParte2 = lcCitaParte2 & "labs: [{codigo: " & lcLabCodigo & ", " & _
                                                          "valor: '" & lcLabValor & "'}], " & lcSaltoDeLinea
                Else
                   lcCitaParte2 = lcCitaParte2 & "labs: '', " & lcSaltoDeLinea
                End If
                lcCitaParte2 = lcCitaParte2 & "tipodiagnostico: '" & lcItemTipoDiagnostico & "', " & _
                                              "codigo: '" & lcItemCodigo & "', " & _
                                              "tipoitem: '" & lcItemTipoItem & "'"
                lcCitaParte2 = lcCitaParte2 & "}" & lcSaltoDeLinea
             End If
             'cuatro
             If Not IsNull(oRsFox1!dx4) Then
                lcLabValor = "***"
                lcLabCodigo = "***"
                If Not IsNull(oRsFox1!lc4) Then
                   lcLabValor = Space(3)
                   lcLabCodigo = "47"
                   oRsLab.Filter = "valores='" & oRsFox1!lc4 & "'"
                   If oRsLab.RecordCount = 0 Then
                       lcErroresHallados = lcErroresHallados & "La CUENTA " & Trim(str(lnIdCuenta)) & " Tiene LAB: " & _
                                           oRsFox1!lc4 & " que no está en la TABLA HIS_situacio" & Chr(13)
                   Else
                       lcLabValor = oRsLab!valores
                       lcLabCodigo = Trim(str(oRsLab!codigo))
                   End If
                End If
                lcItemTipoDiagnostico = oRsFox1!td4
                lcItemCodigo = oRsFox1!dx4
                lcItemTipoItem = "EX"
                Set oRsItems = mo_CabeceraReportes.DiagnosticosSeleccionarXcodigoCIEsinPto(oRsFox1!dx4, oConexion)
                If oRsItems.RecordCount > 0 Then
                   lcItemTipoItem = "CX"
                Else
                   oRsItems.Close
                   Set oRsItems = mo_CabeceraReportes.ProcedimientosSeleccionarPorCodigo(oRsFox1!dx4, oConexion)
                   If oRsItems.RecordCount > 0 Then
                      If oRsItems!IdServicioSubGrupo = 3 Then
                         lcItemTipoItem = "PL"
                      Else
                         lcItemTipoItem = "CP"
                      End If
                   End If
                End If
                oRsItems.Close
                If lcCitaParte2 = lcxItems Then
                   lcCitaParte2 = lcCitaParte2 & "{"
                Else
                   lcCitaParte2 = lcCitaParte2 & ",{"
                End If
                If lcLabValor <> "***" Then
                   lcCitaParte2 = lcCitaParte2 & "labs: [{codigo: " & lcLabCodigo & ", " & _
                                                          "valor: '" & lcLabValor & "'}], " & lcSaltoDeLinea
                Else
                   lcCitaParte2 = lcCitaParte2 & "labs: '', " & lcSaltoDeLinea
                End If
                lcCitaParte2 = lcCitaParte2 & "tipodiagnostico: '" & lcItemTipoDiagnostico & "', " & _
                                              "codigo: '" & lcItemCodigo & "', " & _
                                              "tipoitem: '" & lcItemTipoItem & "'"
                lcCitaParte2 = lcCitaParte2 & "}" & lcSaltoDeLinea
             End If
             'cinco
             If Not IsNull(oRsFox1!dx5) Then
                lcLabValor = "***"
                lcLabCodigo = "***"
                If Not IsNull(oRsFox1!lc5) Then
                   lcLabValor = Space(3)
                   lcLabCodigo = "47"
                   oRsLab.Filter = "valores='" & oRsFox1!lc5 & "'"
                   If oRsLab.RecordCount = 0 Then
                       lcErroresHallados = lcErroresHallados & "La CUENTA " & Trim(str(lnIdCuenta)) & " Tiene LAB: " & _
                                           oRsFox1!lc5 & " que no está en la TABLA HIS_situacio" & Chr(13)
                   Else
                       lcLabValor = oRsLab!valores
                       lcLabCodigo = Trim(str(oRsLab!codigo))
                   End If
                End If
                lcItemTipoDiagnostico = oRsFox1!td5
                lcItemCodigo = oRsFox1!dx5
                lcItemTipoItem = "EX"
                Set oRsItems = mo_CabeceraReportes.DiagnosticosSeleccionarXcodigoCIEsinPto(oRsFox1!dx5, oConexion)
                If oRsItems.RecordCount > 0 Then
                   lcItemTipoItem = "CX"
                Else
                   oRsItems.Close
                   Set oRsItems = mo_CabeceraReportes.ProcedimientosSeleccionarPorCodigo(oRsFox1!dx5, oConexion)
                   If oRsItems.RecordCount > 0 Then
                      If oRsItems!IdServicioSubGrupo = 3 Then
                         lcItemTipoItem = "PL"
                      Else
                         lcItemTipoItem = "CP"
                      End If
                   End If
                End If
                oRsItems.Close
                If lcCitaParte2 = lcxItems Then
                   lcCitaParte2 = lcCitaParte2 & "{"
                Else
                   lcCitaParte2 = lcCitaParte2 & ",{"
                End If
                If lcLabValor <> "***" Then
                   lcCitaParte2 = lcCitaParte2 & "labs: [{codigo: " & lcLabCodigo & ", " & _
                                                          "valor: '" & lcLabValor & "'}], " & lcSaltoDeLinea
                Else
                   lcCitaParte2 = lcCitaParte2 & "labs: '', " & lcSaltoDeLinea
                End If
                lcCitaParte2 = lcCitaParte2 & "tipodiagnostico: '" & lcItemTipoDiagnostico & "', " & _
                                              "codigo: '" & lcItemCodigo & "', " & _
                                              "tipoitem: '" & lcItemTipoItem & "'"
                lcCitaParte2 = lcCitaParte2 & "}" & lcSaltoDeLinea
             End If
             'seis
             If Not IsNull(oRsFox1!dx6) Then
                lcLabValor = "***"
                lcLabCodigo = "***"
                If Not IsNull(oRsFox1!lc6) Then
                   lcLabValor = Space(3)
                   lcLabCodigo = "47"
                   oRsLab.Filter = "valores='" & oRsFox1!lc6 & "'"
                   If oRsLab.RecordCount = 0 Then
                       lcErroresHallados = lcErroresHallados & "La CUENTA " & Trim(str(lnIdCuenta)) & " Tiene LAB: " & _
                                           oRsFox1!lc6 & " que no está en la TABLA HIS_situacio" & Chr(13)
                   Else
                       lcLabValor = oRsLab!valores
                       lcLabCodigo = Trim(str(oRsLab!codigo))
                   End If
                End If
                lcItemTipoDiagnostico = oRsFox1!td6
                lcItemCodigo = oRsFox1!dx6
                lcItemTipoItem = "EX"
                Set oRsItems = mo_CabeceraReportes.DiagnosticosSeleccionarXcodigoCIEsinPto(oRsFox1!dx6, oConexion)
                If oRsItems.RecordCount > 0 Then
                   lcItemTipoItem = "CX"
                Else
                   oRsItems.Close
                   Set oRsItems = mo_CabeceraReportes.ProcedimientosSeleccionarPorCodigo(oRsFox1!dx6, oConexion)
                   If oRsItems.RecordCount > 0 Then
                      If oRsItems!IdServicioSubGrupo = 3 Then
                         lcItemTipoItem = "PL"
                      Else
                         lcItemTipoItem = "CP"
                      End If
                   End If
                End If
                oRsItems.Close
                If lcCitaParte2 = lcxItems Then
                   lcCitaParte2 = lcCitaParte2 & "{"
                Else
                   lcCitaParte2 = lcCitaParte2 & ",{"
                End If
                If lcLabValor <> "***" Then
                   lcCitaParte2 = lcCitaParte2 & "labs: [{codigo: " & lcLabCodigo & ", " & _
                                                          "valor: '" & lcLabValor & "'}], " & lcSaltoDeLinea
                Else
                   lcCitaParte2 = lcCitaParte2 & "labs: '', " & lcSaltoDeLinea
                End If
                lcCitaParte2 = lcCitaParte2 & "tipodiagnostico: '" & lcItemTipoDiagnostico & "', " & _
                                              "codigo: '" & lcItemCodigo & "', " & _
                                              "tipoitem: '" & lcItemTipoItem & "'"
                lcCitaParte2 = lcCitaParte2 & "}" & lcSaltoDeLinea
             End If
             '
             If oRsFox1.EOF Then
                lbUltimoRegistro = True
                Exit Do
             End If
          Loop
          If lcCitaParte2 <> lcxItems Then
             lcCitaParte2 = lcCitaParte2 & "], " & lcSaltoDeLinea                    'Se agregó al menos un Dx/Cpt
          Else
             lcCitaParte2 = ""
          End If
          lcCita = "cita:{" & lcCitaParte1 & "," & lcCitaParte2 & "," & lcCitaParte3 & "}" & lcSaltoDeLinea
          '
          sInputJson = "{" & _
                       lcCabecera & "," & lcCita & "," & _
                       lcUsuario & "," & lcMedico & "," & lcPaciente & _
                       "}"
'          Select Case lnFor
'          Case 1
'                   sInputJson = "{fechaemision: 20170101, idEnvio: '123451',  labs: [{ codigo: 1, valor: '101'},{ codigo: 2,  valor: '1231'}] ,cita:{fecha:20170102,paciente:'juan peres', farm: [{ codigo: '1', cantidad: 101},{ codigo: '2',  cantidad: 1231}] }}"
'          Case 2
'                   sInputJson = "{fechaemision: 20180101, idEnvio: '34512',  labs: [{ codigo: 1, valor: '10'}] }"
'          Case 3
'                   sInputJson = "{fechaemision: 20190101, idEnvio: '9873',  labs: [{ codigo: 1, valor: '103'},{ codigo: 2,  valor: '1233'},{ codigo: 3,  valor: '144'}] }"
'          Case 4
'                   sInputJson = "{fechaemision: 20160101, idEnvio: '9870' }"
'          End Select
          

          Set p = JSON.parse(sInputJson)
          If lbUltimoRegistro = True Then                                         'ultima linea
              act.Write (JSON.toString(p) & lcSaltoDeLinea)
          Else
              act.Write (JSON.toString(p) & "," & lcSaltoDeLinea)   'act.Write (JSON.toString(p) & "," & "\n")   'salto de linea en JSON
          End If
          '
    Loop
    act.Write ("]")
    act.Close
    If lcErroresGraves = "" Then
        MsgBox "Se creó el archivo: " & lcArchivoJSON & Chr(13) & lcErroresHallados & Chr(13) & Chr(13) & _
                lcMensajeLicencia, vbInformation, ""
    Else
        MsgBox "(JSON) Tiene que corregir los ERRORES y luego procesar de nuevo" & Chr(13) & lcErroresGraves & _
               Chr(13) & Chr(13) & lcMensajeLicencia, vbInformation, ""
    End If
    
    Set oRsItems = Nothing
    Set oRsCondicionTrabajo = Nothing
    Set oRsTmp_11 = Nothing
    Set oRsProfesion = Nothing
    Set mo_ReglasAdmision = Nothing
    Set mo_ReglasLaboratorio = Nothing
    Set oRsLab = Nothing
    Exit Sub
ErrJson:
    MsgBox Err.Description & Chr(13) & "En la linea: " & Trim(str(lnFor))
    Exit Sub
    Resume
End Sub


Sub ExportarFacturasBoletas12(lcFechaInicial As String, lcFechaFinal As String, lcNroSerie As String, _
                            lcNroDocumento As String, _
                            ByRef lcCabecera10 As String, ByRef lcCabera11 As String, _
                            ByRef lcCabecera13 As String, ByRef lcCabecera14 As String, _
                            ByRef lcCabecera15 As String, lcNotaCredito As String, _
                            lbUsaResumenDiario As Boolean, Optional oRsAnuladosComoActivos As Recordset)
    Dim Ruta As String
    Dim lcNomComprobanteSunat As String
    Dim lcLinea As String
    Dim Ruc As String
    Dim TipoComprobante As String
    Dim CodigoCobro As String
    Dim oRsDetalleComprobante As New Recordset
    Dim oRsBusquedaRecibos As Recordset
    Dim oDOCajaNroDocumento As DOCajaNroDocumento
    Ruta = sighentidades.Parametro378
    Dim fso As Object 'crea objeto para manejar archivos
    Dim act As Object 'crea objeto para manejar archivos
    Dim NumeroComprobante As String, lnIGV As Integer
    Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim lcServicioInafecto As String, wxParametro549 As String
    Dim lbSeGeneraDesdeRangoFechas As Boolean
    Dim lcFechaResumen As String, ldHoy As Date, lcFechaEmision As String
    ldHoy = CDate(lcBuscaParametro.RetornaFechaServidorSQL)
    lcFechaResumen = Year(ldHoy) & "-" & Right("00" & Month(ldHoy), 2) & "-" & Right("00" & Day(ldHoy), 2)

    lcFechaEmision = "*"
    If lcNroSerie = "" And lcFechaInicial <> "" Then
       Set oRsBusquedaRecibos = mo_AdminCaja.CajaComprobantePagoFiltroPorNroSerieDocumentoOporRangoFemision("", "", CDate(lcFechaInicial), CDate(lcFechaFinal))
       lcFechaEmision = Year(CDate(lcFechaInicial)) & "-" & Right("00" & Month(CDate(lcFechaInicial)), 2) & "-" & Right("00" & Day(CDate(lcFechaInicial)), 2)
    ElseIf lcNroSerie <> "" Then
       Set oRsBusquedaRecibos = mo_AdminCaja.CajaComprobantePagoFiltroPorNroSerieDocumentoOporRangoFemision(lcNroSerie, lcNroDocumento, Date, Date)
    Else
       Set oRsBusquedaRecibos = oRsAnuladosComoActivos.Clone
    End If
    'solo ICA
    If Val(lcBuscaParametro.SeleccionaFilaParametro(280)) = 3359 Then
       oRsBusquedaRecibos.Filter = "idCaja<>2"
    End If
    '
    
    
    Ruc = Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(339)), 11)
    lnIGV = Val(lcBuscaParametro.SeleccionaFilaParametro(221))
    lcServicioInafecto = lcBuscaParametro.SeleccionaFilaParametro(537)
    wxParametro549 = lcBuscaParametro.SeleccionaFilaParametro(549)
    'DATOS CABECERA
    Dim TipoOperacion As String
    Dim FechaEmision As String
    Dim CodDomicilioFiscal As String
    Dim TipoDocumSunat As String
    Dim NroDocumSunat As String
    Dim NombreRazonSocial As String
    Dim TipoMoneda As String
    Dim DescGlobales As String
    Dim SumOtroCargos As String
    Dim TotalDescuentos As String
    Dim TotalVentasOper As String
    Dim TotalVentasInaf As String
    Dim TotalVentasExon As String
    Dim SumatoriaIGV As String
    Dim SumatoriaISC As String
    Dim SumatoriaOtrosTrib As String
    Dim ImporTotalVentas As String
    Dim lbProsigue As Boolean
    'Inicializando
    TipoOperacion = "0101"
    FechaEmision = ""
    CodDomicilioFiscal = "0000"
    TipoMoneda = "PEN"
    DescGlobales = "0.00"
    SumOtroCargos = "0.00"
    TotalDescuentos = "0.00"
    TotalVentasInaf = "0.00"
    TotalVentasExon = "0.00"
    SumatoriaISC = "0.00"
    SumatoriaOtrosTrib = "0.00"
    
    'DATOS DETALLE
    Dim UnidadMedida As String
    Dim CantidadUnids As String
    Dim CodigoProducto As String
    Dim CodigoProductoSunat As String
    Dim DescBienServicio As String
    Dim ValorUnitario As String
    Dim DescuItem As String
    Dim MontoIGV As String
    Dim AfectacionIGV As String
    Dim MontoISC As String
    Dim TipoISC As String
    Dim PrecioVentaUnitario As String
    Dim ValorVenta As String
    Dim DocumentoBaja As String, DescripcionBaja As String, CorrelativoBaja As String
    Dim lnFacturaSinIGV As Integer
    Dim lnPrecioUnitarioSinIgv As Double, lnMontoIGV As Double, lnTotalVentasOper As Double
    Dim oRsTmp876 As New Recordset, lcSql As String, lnImporteParcial As Double
    Dim oRsTmp877 As New Recordset
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    
    
    'Inicializando
    UnidadMedida = "NIU"
    TipoISC = "01"
    DescuItem = "0.00"
    AfectacionIGV = "10"
    MontoIGV = "0.00"
    MontoISC = "0.00"
    
    Dim lcHoraEmision As String, lcFechaVencimiento As String, lcSumatoriaDeTributos As String, lcTotalValorVenta As String
    Dim lcTotalAnticipos As String, lcVersionUBL As String, lcCustomizacionDocumento As String
    Dim lcDescuentos As String, lcISC As String, lcOtrosTributos As String, lcOtrosCargos As String
    Dim lcIGVgrilla As String, lcSubTotalGrilla As String, lnIGVgrilla As Double, lnSubTotalGrilla As Double
    Dim lnSumatoriaDeTributos As Double, lnTotalValorVenta As Double, lcTotalPrecioVenta As String
    Dim lcImporteTotalDeLaVenta As String, lnIGVbase As Double
    Dim oDoSunatResumenDia As New DoSunatResumenDia
    Dim oSunatParteDia As New SunatParteDia
    Dim oConexion As New Connection
    
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    Set oSunatParteDia.Conexion = oConexion
    
    lcHoraEmision = ""
    lcFechaVencimiento = "-"
    lcSumatoriaDeTributos = ""
    lcTotalValorVenta = ""
    lcTotalAnticipos = "0.00"
    lcVersionUBL = "2.1"
    lcCustomizacionDocumento = "2.0"
    lcDescuentos = "0.00"
    lcISC = "0.00"
    lcOtrosTributos = "0.00"
    lcOtrosCargos = "0.00"
    
    Dim lcCodigoProductoSunat As String, lcIGVnombre As String, lcIGVcodigo As String
    Dim lcIGVcodigoInt As String, lcSumatoriaProductosPorItem As String, lcBaseImponibleIgvPorItem As String
    Dim lcPorcentajeIGV As String, lbUsaCodigoSunatPorItem As Boolean, lbGrabaRD As Boolean
    Dim lcISCcodigoISC As String, lcISCmonto As String, lcISCbaseImponible As String, lcISCnombreTributo As String
    Dim lcISCcodigoXitem As String, lcISCtipoSistema As String, lcISCporcentaje As String, lcOTROcodigoISC As String
    Dim lcOTROmonto As String, lcOTRObaseImponible As String, lcOTROnombreTributo As String, lcOTROcodigoXitem As String
    Dim lcOTROporcentaje As String, lcValorReferencialUnitario As String, lcIGV As String, lcCantidadUnids As String
    Dim lcImporteSoles As String, lnImporteTotalDeLaVenta As Double, lnMontoIGVxITEM As Double, lcMontoIGVxITEM As String
    Dim lcValorVentaConIGV  As String, lcValorVentaSinIGV As String, lcTotalVentasExon As String, lbNoGeneraResumenDiario As Boolean
    Dim oRsTmpBoletas As New Recordset
    lbUsaCodigoSunatPorItem = IIf(lcBuscaParametro.SeleccionaFilaParametro(570) = "S", True, False)
    
    
    If oRsBusquedaRecibos.RecordCount > 0 Then
        oRsBusquedaRecibos.MoveFirst
        Do While Not oRsBusquedaRecibos.EOF
            AfectacionIGV = "10"
            TotalVentasInaf = "0.00"
            lnIGV = Val(lcBuscaParametro.SeleccionaFilaParametro(221))
            lnIGVbase = lnIGV
            lcPorcentajeIGV = Format(CCur(lnIGV), "#0.00")
            lcIGVnombre = "IGV"
            lcIGVcodigoInt = "VAT"
            lcIGVcodigo = "1000"
            lcCodigoProductoSunat = "-"
            lcISCcodigoISC = "-"
            lcISCmonto = "0.00"
            lcISCbaseImponible = "0.00"
            lcISCnombreTributo = ""
            lcISCcodigoXitem = ""
            lcISCtipoSistema = ""
            lcISCporcentaje = "0.00"
            lcOTROcodigoISC = "-"
            lcOTROmonto = "0.00"
            lcOTRObaseImponible = "0.00"
            lcOTROnombreTributo = ""
            lcOTROcodigoXitem = ""
            lcOTROporcentaje = "0.00"
            
            lcValorReferencialUnitario = "0.00"
            
            lbProsigue = True
            'No se considera BOLETAS/FACTURAS de SERVICIOS
            If lcBuscaParametro.SeleccionaFilaParametro(532) = "S" And oRsBusquedaRecibos!IdTipoOrden = sghCajaTipoOrden.sghCajaServicioCpt Then
                lbProsigue = False
            End If
            '
            If lbProsigue = True Then
                 lbNoGeneraResumenDiario = True
                 If lbUsaResumenDiario = True Then
                   If oRsBusquedaRecibos.Fields!idTipoComprobante <> sghCajaTipoComprobante.sghCajaFactura Then
                      lbNoGeneraResumenDiario = False
                   End If
                 End If
                 If oRsBusquedaRecibos!IdEstadoComprobante = 9 Then
                      If lbNoGeneraResumenDiario = False Then
                            '********************************Boleta/Factura anulada****************************
                            FechaEmision = Year(oRsBusquedaRecibos.Fields!fechaCobranza) & "-" & Right("00" & Month(oRsBusquedaRecibos.Fields!fechaCobranza), 2) & "-" & Right("00" & Day(oRsBusquedaRecibos.Fields!fechaCobranza), 2)
                            TipoComprobante = "RA"
                            If oRsBusquedaRecibos.Fields!idTipoComprobante = sghCajaTipoComprobante.sghCajaFactura Then '
                               TipoDocumSunat = "01" 'REG. UNICO DE CONTRIBUYENTES
                            Else
                               TipoDocumSunat = "03" 'boleta
                            End If
                            NumeroComprobante = Right("0000000" & Trim(oRsBusquedaRecibos!nroDocumento), 8)
                            DocumentoBaja = oRsBusquedaRecibos!NroSerie & "-" & NumeroComprobante
                            DescripcionBaja = "No se entregó al Cliente"
                            CorrelativoBaja = mo_AdminCaja.DevuelveCorrelativoDeFacturaBoletaAnulada(oRsBusquedaRecibos!NroSerie, oRsBusquedaRecibos!nroDocumento)
                            
                            lcNomComprobanteSunat = Ruta & "\" & Ruc & "-" & TipoComprobante & "-" & _
                                                    Format(oRsBusquedaRecibos!fechaCobranza, "yyyymmdd") & _
                                                    "-" & CorrelativoBaja
'                            Set fso = CreateObject("scripting.filesystemobject") 'creo objeto de manejo de archivos
'                            Set act = fso.CreateTextFile(lcNomComprobanteSunat & ".cba", True) 'creo el archivo con nombre asignado por parametro
'                            lcLinea = FechaEmision & "|" & FechaEmision & "|" & TipoDocumSunat & "|" & _
'                                      DocumentoBaja & "|" & DescripcionBaja
'                            act.Write (lcLinea) 'Graba la linea
                            '****resumen diario*****
                            If lcFechaEmision <> "*" Then
                                TipoDocumSunat = "0"
                                If oRsBusquedaRecibos.Fields!idTipoComprobante = sghCajaTipoComprobante.sghCajaFactura Then
                                    TipoComprobante = "01"
                                    TipoDocumSunat = "6" 'REG. UNICO DE CONTRIBUYENTES
                                    NroDocumSunat = oRsBusquedaRecibos.Fields!Ruc
                                Else
                                    TipoComprobante = "03"
                                    TipoDocumSunat = "1"
                                    If IsNull(oRsBusquedaRecibos.Fields!idPaciente) Then
                                        TipoDocumSunat = "1"                 'kike 2017
                                        NroDocumSunat = "00000000"
                                        If Not IsNull(oRsBusquedaRecibos!dni) Then
                                             If Len(Trim(oRsBusquedaRecibos!dni)) = 8 Then
                                                NroDocumSunat = oRsBusquedaRecibos!dni
                                             End If
                                        End If
                                    Else
                                        Dim oDOPaciente As New DOPaciente
                                        Set oDOPaciente = mo_ReglasAdmision.RetornaPacientesSeleccionarPorId(oRsBusquedaRecibos.Fields!idPaciente)
                                        Select Case oDOPaciente.IdDocIdentidad
                                        Case 1
                                            TipoDocumSunat = "1" 'DOC. NACIONAL DE IDENTIDAD
                                        Case 2
                                            TipoDocumSunat = "4" 'CARNET DE EXTRANJERIA
                                        Case 3
                                            TipoDocumSunat = "7" 'PASAPORTE
                                        End Select
                                        NroDocumSunat = Right("00000000" & oDOPaciente.nroDocumento, 8)
                                        Set oDOPaciente = Nothing
                                    End If
                                End If
                                TotalVentasInaf = "0.00"
                                SumatoriaIGV = "0.00"
                                Set oRsTmpBoletas = mo_AdminCaja.CajaComprobantesSeleccionarPorId(oRsBusquedaRecibos!IdComprobantePago, oConexion)
                                If oRsTmpBoletas.RecordCount > 0 Then
                                   If Not IsNull(oRsTmpBoletas!SunatTotalIgv) Then
                                        SumatoriaIGV = Format(oRsTmpBoletas!SunatTotalIgv, "#0.00")
                                   End If
                                   If Not IsNull(oRsTmpBoletas!SunatOpeInafectas) Then
                                        TotalVentasInaf = Format(oRsTmpBoletas!SunatOpeInafectas, "#0.00")
                                   End If
                                End If
                                oRsTmpBoletas.Close
                                lcTotalVentasExon = Format(CCur(oRsBusquedaRecibos.Fields!exoneraciones), "#0.00")
                                lcImporteTotalDeLaVenta = Format(CCur(oRsBusquedaRecibos.Fields!Total), "#0.00")
                                lbGrabaRD = True
                                If oRsBusquedaRecibos!exoneraciones > 0 Then
                                   If oRsBusquedaRecibos!Total = 0 Then
                                      TotalVentasInaf = "0.01"
                                      lcImporteTotalDeLaVenta = TotalVentasInaf
                                      SumatoriaIGV = "0.00"
                                      lbGrabaRD = False
                                   End If
                                End If
                                If lbGrabaRD = True Then
                                    'boleta activa
                                    oDoSunatResumenDia.FechaEmision = FechaEmision
                                    oDoSunatResumenDia.fechaResumen = FechaEmision  'lcFechaResumen
                                    oDoSunatResumenDia.tipoDcto = "03"  'tipo Boleta
                                    oDoSunatResumenDia.SerieDocumento = Trim(oRsBusquedaRecibos!NroSerie) & "-" & NumeroComprobante 'de la Boleta
                                    oDoSunatResumenDia.PacienteTipoDoc = TipoDocumSunat
                                    oDoSunatResumenDia.PacienteNumeroDoc = NroDocumSunat
                                    oDoSunatResumenDia.Moneda = TipoMoneda
                                    If Val(TotalVentasInaf) > 0 Then
                                       oDoSunatResumenDia.OpeGravadas = "0.00"
                                    Else
                                       oDoSunatResumenDia.OpeGravadas = Format(CCur(lcImporteTotalDeLaVenta) - CCur(SumatoriaIGV), "#0.00")
                                    End If
                                    oDoSunatResumenDia.OpeExoneradas = "0.00"    'lcTotalVentasExon
                                    oDoSunatResumenDia.OpeInafectas = TotalVentasInaf
                                    oDoSunatResumenDia.OpeGratuitas = "0.00"
                                    oDoSunatResumenDia.OtrosCargos = "0.00"
                                    oDoSunatResumenDia.TotalIsc = "0.00"
                                    oDoSunatResumenDia.TotalIgv = SumatoriaIGV
                                    oDoSunatResumenDia.TotalOtros = "0.00"
                                    oDoSunatResumenDia.ImporteVenta = lcImporteTotalDeLaVenta
                                    oDoSunatResumenDia.ModTipoDcto = " "
                                    oDoSunatResumenDia.ModSerie = " "
                                    oDoSunatResumenDia.ModDocumento = " "
                                    oDoSunatResumenDia.PercepRegimen = " "
                                    oDoSunatResumenDia.PercepPorcen = " "
                                    oDoSunatResumenDia.PercepBaseImp = "0"
                                    oDoSunatResumenDia.PercepMonto = "0"
                                    oDoSunatResumenDia.PercepCobrar = "0"
                                    oDoSunatResumenDia.Estado = "1"
                                    If oSunatParteDia.Insertar(oDoSunatResumenDia) Then
                                    End If
                                    'boleta anulada
                                    oDoSunatResumenDia.FechaEmision = FechaEmision
                                    oDoSunatResumenDia.fechaResumen = FechaEmision   'lcFechaResumen
                                    oDoSunatResumenDia.tipoDcto = "03"  'tipo Boleta
                                    oDoSunatResumenDia.SerieDocumento = Trim(oRsBusquedaRecibos!NroSerie) & "-" & NumeroComprobante 'de la Boleta
                                    oDoSunatResumenDia.PacienteTipoDoc = TipoDocumSunat
                                    oDoSunatResumenDia.PacienteNumeroDoc = NroDocumSunat
                                    oDoSunatResumenDia.Moneda = TipoMoneda
                                    If Val(TotalVentasInaf) > 0 Then
                                       oDoSunatResumenDia.OpeGravadas = "0.00"
                                    Else
                                       oDoSunatResumenDia.OpeGravadas = Format(CCur(lcImporteTotalDeLaVenta) - CCur(SumatoriaIGV), "#0.00")
                                    End If
                                    oDoSunatResumenDia.OpeExoneradas = lcTotalVentasExon
                                    oDoSunatResumenDia.OpeInafectas = TotalVentasInaf
                                    oDoSunatResumenDia.OpeGratuitas = "0.00"
                                    oDoSunatResumenDia.OtrosCargos = "0.00"
                                    oDoSunatResumenDia.TotalIsc = "0.00"
                                    oDoSunatResumenDia.TotalIgv = SumatoriaIGV
                                    oDoSunatResumenDia.TotalOtros = "0.00"
                                    oDoSunatResumenDia.ImporteVenta = lcImporteTotalDeLaVenta
                                    oDoSunatResumenDia.ModTipoDcto = " "
                                    oDoSunatResumenDia.ModSerie = " "
                                    oDoSunatResumenDia.ModDocumento = " "
                                    oDoSunatResumenDia.PercepRegimen = " "
                                    oDoSunatResumenDia.PercepPorcen = " "
                                    oDoSunatResumenDia.PercepBaseImp = "0"
                                    oDoSunatResumenDia.PercepMonto = "0"
                                    oDoSunatResumenDia.PercepCobrar = "0"
                                    oDoSunatResumenDia.Estado = "3"
                                    If oSunatParteDia.Insertar(oDoSunatResumenDia) Then
                                    End If
                                End If
                            End If
                      End If
                 Else
                      '*******************************Boleta/Factura activas****************************
                      'debb-24/07/2018-inicio-chimbote
                      lnFacturaSinIGV = 99
                      If oRsBusquedaRecibos.Fields!idTipoComprobante = sghCajaTipoComprobante.sghCajaFactura Then
                         lnFacturaSinIGV = 0
                         Set oDOCajaNroDocumento = mo_AdminCaja.CajaNroDocumentoSeleccionarPorId(oRsBusquedaRecibos!idCaja, 2)
                         If Not oDOCajaNroDocumento Is Nothing Then
                            If oDOCajaNroDocumento.FacturaSinIGV = True Then
                               lnFacturaSinIGV = 1
                            End If
                         End If
                      End If
                      'debb-24/07/2018-fin
                      
                      TipoDocumSunat = "0"
                      If oRsBusquedaRecibos.Fields!idTipoComprobante = sghCajaTipoComprobante.sghCajaFactura Then
                          TipoComprobante = "01"
                          TipoDocumSunat = "6" 'REG. UNICO DE CONTRIBUYENTES
                          NroDocumSunat = oRsBusquedaRecibos.Fields!Ruc
                      Else
                          TipoComprobante = "03"
                          TipoDocumSunat = "1"
                          If IsNull(oRsBusquedaRecibos.Fields!idPaciente) Then
                              TipoDocumSunat = "1"                 'kike 2017
                              NroDocumSunat = "00000000"
                              If Not IsNull(oRsBusquedaRecibos!dni) Then
                                   If Len(Trim(oRsBusquedaRecibos!dni)) = 8 Then
                                      NroDocumSunat = oRsBusquedaRecibos!dni
                                   End If
                              End If
                          Else
                              Dim oDOPaciente1 As New DOPaciente
                              Set oDOPaciente1 = mo_ReglasAdmision.RetornaPacientesSeleccionarPorId(oRsBusquedaRecibos.Fields!idPaciente)
                              Select Case oDOPaciente1.IdDocIdentidad
                              Case 1
                                  TipoDocumSunat = "1" 'DOC. NACIONAL DE IDENTIDAD
                              Case 2
                                  TipoDocumSunat = "4" 'CARNET DE EXTRANJERIA
                              Case 3
                                  TipoDocumSunat = "7" 'PASAPORTE
                              End Select
                              NroDocumSunat = Right("00000000" & oDOPaciente1.nroDocumento, 8)
                              Set oDOPaciente1 = Nothing
                          End If
                      End If
                      NombreRazonSocial = Trim(Left(oRsBusquedaRecibos.Fields!razonSocial, 100))
                      NumeroComprobante = Right("00000000" & Trim(oRsBusquedaRecibos.Fields!nroDocumento), 8)
If NumeroComprobante = "00000008" Then
lcNomComprobanteSunat = ""
End If
                      lcNomComprobanteSunat = Ruta & "\" & Ruc & "-" & TipoComprobante & "-" & _
                                             Trim(oRsBusquedaRecibos.Fields!NroSerie) & "-" & NumeroComprobante
                      
                      'CREANDO ARCHIVO CABECERA
                      If (lcServicioInafecto = "S" Or lcServicioInafecto = "SS") And oRsBusquedaRecibos!IdTipoOrden = sghCajaTipoOrden.sghCajaServicioCpt Then
                         lnIGV = 0
                      ElseIf lnFacturaSinIGV <> 99 Then
                         If lnFacturaSinIGV = 1 Then
                            lnIGV = 0
                         End If
                      End If
                      'debb-24/07/2018-fin
                      If oRsBusquedaRecibos.Fields!exoneraciones > 0 Then
                          
                            lnTotalVentasOper = Round(oRsBusquedaRecibos.Fields!exoneraciones / ((100 + lnIGV) / 100), 2)
                            FechaEmision = Year(oRsBusquedaRecibos.Fields!fechaCobranza) & "-" & Right("00" & Month(oRsBusquedaRecibos.Fields!fechaCobranza), 2) & "-" & Right("00" & Day(oRsBusquedaRecibos.Fields!fechaCobranza), 2)
                            TotalVentasOper = Format(lnTotalVentasOper, "#0.00")
                            TotalVentasExon = Format(CCur(oRsBusquedaRecibos!exoneraciones), "#0.00")
                            SumatoriaIGV = Format(Round(lnTotalVentasOper * lnIGV / 100, 2), "#0.00")
                            ImporTotalVentas = Format(oRsBusquedaRecibos!Total, "#0.00")
                            lcTotalVentasExon = TotalVentasExon
                      Else
                            lnTotalVentasOper = Round(oRsBusquedaRecibos.Fields!Total / ((100 + lnIGV) / 100), 2)
                            FechaEmision = Year(oRsBusquedaRecibos.Fields!fechaCobranza) & "-" & Right("00" & Month(oRsBusquedaRecibos.Fields!fechaCobranza), 2) & "-" & Right("00" & Day(oRsBusquedaRecibos.Fields!fechaCobranza), 2)
                            TotalVentasOper = Format(lnTotalVentasOper, "#0.00")
                            TotalVentasExon = Format(CCur(oRsBusquedaRecibos.Fields!exoneraciones), "#0.00")
                            SumatoriaIGV = Format(Round(lnTotalVentasOper * lnIGV / 100, 2), "#0.00")
                            ImporTotalVentas = Format(lnTotalVentasOper + CCur(SumatoriaIGV), "#0.00")
                            lcTotalVentasExon = ""
                      End If
                      If (lcServicioInafecto = "S" Or lcServicioInafecto = "SS") And oRsBusquedaRecibos!IdTipoOrden = sghCajaTipoOrden.sghCajaServicioCpt Then
                         TotalVentasOper = Format(0, "#0.00")
                         TotalVentasInaf = Format(oRsBusquedaRecibos!Total, "#0.00")
                      End If
                      'KIKE2019
                      lcHoraEmision = Format(oRsBusquedaRecibos.Fields!fechaCobranza, sighentidades.DevuelveHoraSoloFormato_HMS)
                      '
                      lnIGVgrilla = 0
                      lnSubTotalGrilla = 0
                      If oRsBusquedaRecibos.Fields!IdTipoOrden = sghCajaTipoOrden.sghCajaFarmacia2 Or _
                                                    oRsBusquedaRecibos.Fields!IdTipoOrden = sghCajaTipoOrden.sghCajaFarmacia3 Then  'Bien-Farmacia
                          Set oRsDetalleComprobante = mo_AdminCaja.CajaComprobantePagoProductosPorIdComprobante(oRsBusquedaRecibos.Fields!IdComprobantePago)
                          If oRsDetalleComprobante.RecordCount > 0 Then
                              oRsDetalleComprobante.MoveFirst
                              Do While Not oRsDetalleComprobante.EOF
                                  lnImporteParcial = oRsDetalleComprobante!PrecioUnitario * oRsDetalleComprobante!Cantidad
                                  lnPrecioUnitarioSinIgv = Round((oRsDetalleComprobante!PrecioUnitario / ((100 + lnIGV) / 100)), 10)
                                  lnMontoIGV = Round(lnPrecioUnitarioSinIgv * lnIGV / 100, 2)
                                  lnMontoIGVxITEM = lnImporteParcial - (oRsDetalleComprobante!Cantidad * lnPrecioUnitarioSinIgv)
                                  lnIGVgrilla = lnIGVgrilla + lnMontoIGVxITEM
                                  lnSubTotalGrilla = lnSubTotalGrilla + (oRsDetalleComprobante!Cantidad * lnPrecioUnitarioSinIgv)
                                  oRsDetalleComprobante.MoveNext
                              Loop
                          End If
                          oRsDetalleComprobante.Close
                      Else '1 Servicio
                          Set oRsDetalleComprobante = mo_AdminCaja.CajaComprobantePagoServiciosPorIdComprobante(oRsBusquedaRecibos.Fields!IdComprobantePago)
                          If oRsDetalleComprobante.RecordCount > 0 Then
                              oRsDetalleComprobante.MoveFirst
                              Do While Not oRsDetalleComprobante.EOF
                                  'Boletas/Facturas de SERVICIO - INAFECTAS - JAMO
                                  If lcServicioInafecto = "S" And oRsBusquedaRecibos!IdTipoOrden = sghCajaTipoOrden.sghCajaServicioCpt Then
                                     lnIGV = 0
                                  End If
                                  ValorVenta = Format(CCur(oRsDetalleComprobante!TotalPorPagar), "#0.00")
                                  PrecioVentaUnitario = Format(Round(oRsDetalleComprobante.Fields!TotalPorPagar * 100 / (lnIGV + 100), 2), "#0.00")
                                  lnPrecioUnitarioSinIgv = Round((oRsDetalleComprobante!PrecioUnitario / ((100 + lnIGV) / 100)), 10)
                                  MontoIGV = Format((oRsDetalleComprobante.Fields!TotalPorPagar - CCur(PrecioVentaUnitario)), "#0.00")
                                  lnMontoIGVxITEM = oRsDetalleComprobante!TotalPorPagar - (oRsDetalleComprobante!Cantidad * lnPrecioUnitarioSinIgv)
                                  lnIGVgrilla = lnIGVgrilla + lnMontoIGVxITEM
                                  lnSubTotalGrilla = lnSubTotalGrilla + (oRsDetalleComprobante!Cantidad * lnPrecioUnitarioSinIgv)
                                  oRsDetalleComprobante.MoveNext
                              Loop
                          End If
                          oRsDetalleComprobante.Close
                      End If
                      lcIGVgrilla = Format(CCur(lnIGVgrilla), "#0.00")
                      lcSubTotalGrilla = Format(CCur(lnSubTotalGrilla), "#0.00")
                      '
                      lnSumatoriaDeTributos = lnIGVgrilla + CCur(lcISC) + CCur(lcOtrosTributos) + CCur(lcOtrosCargos)
                      lcSumatoriaDeTributos = Format(CCur(lnSumatoriaDeTributos), "#0.00")
                      '
                      lnTotalValorVenta = lnSubTotalGrilla - CCur(lcDescuentos)
                      lcTotalValorVenta = Format(CCur(lnTotalValorVenta), "#0.00")
                      '
                      lcTotalPrecioVenta = Format(CCur(lnSumatoriaDeTributos + lnTotalValorVenta), "#0.00")
                      lcImporteTotalDeLaVenta = Format(CCur(lcTotalPrecioVenta) + CCur(lcDescuentos) + CCur(lcOtrosCargos) - CCur(lcTotalAnticipos), "#0.00")
                      '
                      If lbNoGeneraResumenDiario = True Then
                            If lcNotaCredito = "" Then
                                  Set fso = CreateObject("scripting.filesystemobject") 'creo objeto de manejo de archivos
                                  Set act = fso.CreateTextFile(lcNomComprobanteSunat & ".cab", True) 'creo el archivo con nombre asignado por parametro
                                  lcLinea = TipoOperacion & "|" & FechaEmision & "|" & _
                                            lcHoraEmision & "|" & lcFechaVencimiento & "|" & _
                                            CodDomicilioFiscal & "|" & TipoDocumSunat & "|" & _
                                            NroDocumSunat & "|" & NombreRazonSocial & "|" & _
                                            TipoMoneda & "|" & lcSumatoriaDeTributos & "|" & _
                                            lcTotalValorVenta & "|" & lcTotalPrecioVenta & "|" & _
                                            lcDescuentos & "|" & lcOtrosCargos & "|" & _
                                            lcTotalAnticipos & "|" & lcImporteTotalDeLaVenta & "|" & _
                                            lcVersionUBL & "|" & lcCustomizacionDocumento & "|"
                                  act.Write (lcLinea) 'Graba la linea
                            Else
                                  Set fso = CreateObject("scripting.filesystemobject") 'creo objeto de manejo de archivos
                                  lcNomComprobanteSunat = lcNotaCredito
                            End If
                      End If
                      
                      mo_AdminCaja.CajaComprobantesPagoActualizaDatosSunat oRsBusquedaRecibos!IdComprobantePago, _
                                                                           CCur(TotalVentasInaf), CCur(SumatoriaIGV), oConexion
                      lcCabecera10 = lcSumatoriaDeTributos
                      lcCabera11 = lcTotalValorVenta
                      lcCabecera13 = lcDescuentos
                      lcCabecera14 = lcOtrosCargos
                      lcCabecera15 = lcTotalAnticipos
                               
                      'CREANDO ARCHIVO DETALLE
                      If lbNoGeneraResumenDiario = True Then
                        Set act = fso.CreateTextFile(lcNomComprobanteSunat & ".det", True) 'creo el archivo con nombre asignado por parametro
                        If oRsBusquedaRecibos.Fields!IdTipoOrden = sghCajaTipoOrden.sghCajaFarmacia2 Or _
                                  oRsBusquedaRecibos.Fields!IdTipoOrden = sghCajaTipoOrden.sghCajaFarmacia3 Then  'Bien-Farmacia
                            Set oRsDetalleComprobante = mo_AdminCaja.CajaComprobantePagoProductosPorIdComprobante(oRsBusquedaRecibos.Fields!IdComprobantePago)
                            If oRsDetalleComprobante.RecordCount > 0 Then
                                oRsDetalleComprobante.MoveFirst
                                Do While Not oRsDetalleComprobante.EOF
                                    If lbUsaCodigoSunatPorItem = True Then
                                       If Not IsNull(oRsDetalleComprobante!codigoSunat) Then
                                          lcCodigoProductoSunat = Trim(oRsDetalleComprobante!codigoSunat)
                                       End If
                                    End If
                                    lnImporteParcial = oRsDetalleComprobante!PrecioUnitario * oRsDetalleComprobante!Cantidad
                                    lnPrecioUnitarioSinIgv = Round((oRsDetalleComprobante!PrecioUnitario / ((100 + lnIGV) / 100)), 10)
                                    lnMontoIGV = Round(lnPrecioUnitarioSinIgv * lnIGV / 100, 2)
                                    CantidadUnids = Format(CCur(oRsDetalleComprobante.Fields!Cantidad), "#0")
                                    CodigoProducto = Trim(oRsDetalleComprobante.Fields!codigo)
                                    DescBienServicio = oRsDetalleComprobante.Fields!nombreProducto
                                    ValorUnitario = Format(CCur(lnPrecioUnitarioSinIgv), "#0.0000000000")
                                    MontoIGV = Format(CCur(lnMontoIGV), "#0.00")
                                    PrecioVentaUnitario = Format(lnPrecioUnitarioSinIgv + lnMontoIGV, "#0.00")
                                    ValorVenta = Format(Round(oRsDetalleComprobante!Cantidad * lnPrecioUnitarioSinIgv, 2), "#0.00")
                                    'kike 2019
                                    lnMontoIGVxITEM = lnImporteParcial - (oRsDetalleComprobante!Cantidad * lnPrecioUnitarioSinIgv)
                                    lcMontoIGVxITEM = Format(lnMontoIGVxITEM, "#0.00")
                                    lcBaseImponibleIgvPorItem = Format(Round(oRsDetalleComprobante!Cantidad * lnPrecioUnitarioSinIgv, 2), "#0.00")
                                    lcIGV = Format(lnIGV, "#0.00")
                                    lcSumatoriaProductosPorItem = Format(CCur(lcMontoIGVxITEM) + CCur(lcISCmonto) + CCur(lcOTROmonto), "#0.00")
                                    lcCantidadUnids = CantidadUnids + ".0000"
                                    lcValorVentaConIGV = Format(lnImporteParcial, "#0.00")
                                    lcValorVentaSinIGV = lcBaseImponibleIgvPorItem
                                    '
                                    lcLinea = UnidadMedida & "|" & lcCantidadUnids & "|" & _
                                    CodigoProducto & "|" & lcCodigoProductoSunat & "|" & _
                                    DescBienServicio & "|" & ValorUnitario & "|" & _
                                    lcSumatoriaProductosPorItem & "|" & lcIGVcodigo & "|" & _
                                    lcMontoIGVxITEM & "|" & lcBaseImponibleIgvPorItem & "|" & _
                                    lcIGVnombre & "|" & lcIGVcodigoInt & "|" & _
                                    AfectacionIGV & "|" & lcIGV & "|" & _
                                    lcISCcodigoISC & "|" & lcISCmonto & "|" & _
                                    lcISCbaseImponible & "|" & lcISCnombreTributo & "|" & _
                                    lcISCcodigoXitem & "|" & lcISCtipoSistema & "|" & _
                                    lcISCporcentaje & "|" & lcOTROcodigoISC & "|" & _
                                    lcOTROmonto & "|" & lcOTRObaseImponible & "|" & _
                                    lcOTROnombreTributo & "|" & lcOTROcodigoXitem & "|" & _
                                    lcOTROporcentaje & "|" & lcValorVentaConIGV & "|" & _
                                    lcValorVentaSinIGV & "|" & lcValorReferencialUnitario & "|"
                                    '
                                    act.Write (lcLinea & vbCrLf)    'Graba la linea
                                    oRsDetalleComprobante.MoveNext
                                Loop
                            End If
                        Else '1 Servicio
                            Set oRsDetalleComprobante = mo_AdminCaja.CajaComprobantePagoServiciosPorIdComprobante(oRsBusquedaRecibos.Fields!IdComprobantePago)
                            If oRsDetalleComprobante.RecordCount > 0 Then
                                oRsDetalleComprobante.MoveFirst
                                Do While Not oRsDetalleComprobante.EOF
                                    'Boletas/Facturas de SERVICIO - INAFECTAS - JAMO
                                    If lcServicioInafecto = "S" And oRsBusquedaRecibos!IdTipoOrden = sghCajaTipoOrden.sghCajaServicioCpt Then
                                       AfectacionIGV = "30"
                                       lnIGV = 0
                                       lcIGVnombre = "INA"
                                       lcIGVcodigoInt = "FRE"
                                       lcIGVcodigo = "9998"
                                    End If
                                    '
                                    If lbUsaCodigoSunatPorItem = True Then
                                       If Not IsNull(oRsDetalleComprobante!codigoSunat) Then
                                          lcCodigoProductoSunat = Trim(oRsDetalleComprobante!codigoSunat)
                                       End If
                                    End If
                                    lnImporteParcial = oRsDetalleComprobante.Fields!TotalPorPagar
                                    CantidadUnids = Format(CCur(oRsDetalleComprobante.Fields!Cantidad), "#0")
                                    CodigoProducto = Trim(oRsDetalleComprobante.Fields!codigo)
                                    If oRsDetalleComprobante!idProducto = Val(wxParametro549) Then
                                       DescBienServicio = sighentidades.RetornaTextoSinChr13Chr10(oRsDetalleComprobante!Observaciones)
                                    Else
                                       DescBienServicio = oRsDetalleComprobante.Fields!nombreProducto
                                    End If
                                    ValorUnitario = Format(CCur(oRsDetalleComprobante.Fields!PrecioUnitario), "#0.00")
                                    MontoIGV = "0.00"
                                    ValorVenta = Format(CCur(oRsDetalleComprobante.Fields!TotalPorPagar), "#0.00")
                                    '
                                    PrecioVentaUnitario = Format(Round(oRsDetalleComprobante.Fields!TotalPorPagar * 100 / (lnIGV + 100), 2), "#0.00")
                                    MontoIGV = Format((oRsDetalleComprobante.Fields!TotalPorPagar - CCur(PrecioVentaUnitario)), "#0.00")
                                    'kike 2019
                                    lnPrecioUnitarioSinIgv = Round((oRsDetalleComprobante!PrecioUnitario / ((100 + lnIGV) / 100)), 10)
                                    lnMontoIGVxITEM = oRsDetalleComprobante!TotalPorPagar - (oRsDetalleComprobante!Cantidad * lnPrecioUnitarioSinIgv)
                                    lcMontoIGVxITEM = Format(lnMontoIGVxITEM, "#0.00")
                                    lcSumatoriaProductosPorItem = Format(CCur(lcMontoIGVxITEM) + CCur(lcISCmonto) + CCur(lcOTROmonto), "#0.00")
                                    lcCantidadUnids = CantidadUnids + ".0000"
                                    If lcServicioInafecto = "S" And oRsBusquedaRecibos!IdTipoOrden = sghCajaTipoOrden.sghCajaServicioCpt Then
                                       lcBaseImponibleIgvPorItem = "0.00"
                                       lcIGV = Format(lnIGVbase, "#0.00")
                                       lcValorVentaConIGV = Format(CCur(lnImporteParcial), "#0.00")
                                       lcValorVentaSinIGV = Format(CCur(lnImporteParcial), "#0.00")
                                    Else
                                       lcBaseImponibleIgvPorItem = Format(Round(oRsDetalleComprobante!Cantidad * lnPrecioUnitarioSinIgv, 2), "#0.00")
                                       lcIGV = Format(lnIGV, "#0.00")
                                       lcValorVentaConIGV = Format(lnImporteParcial, "#0.00")
                                       lcValorVentaSinIGV = lcBaseImponibleIgvPorItem
                                    End If
                                    
                                    '
                                    lcLinea = UnidadMedida & "|" & lcCantidadUnids & "|" & _
                                    CodigoProducto & "|" & lcCodigoProductoSunat & "|" & _
                                    DescBienServicio & "|" & ValorUnitario & "|" & _
                                    lcSumatoriaProductosPorItem & "|" & lcIGVcodigo & "|" & _
                                    lcMontoIGVxITEM & "|" & lcBaseImponibleIgvPorItem & "|" & _
                                    lcIGVnombre & "|" & lcIGVcodigoInt & "|" & _
                                    AfectacionIGV & "|" & lcIGV & "|" & _
                                    lcISCcodigoISC & "|" & lcISCmonto & "|" & _
                                    lcISCbaseImponible & "|" & lcISCnombreTributo & "|" & _
                                    lcISCcodigoXitem & "|" & lcISCtipoSistema & "|" & _
                                    lcISCporcentaje & "|" & lcOTROcodigoISC & "|" & _
                                    lcOTROmonto & "|" & lcOTRObaseImponible & "|" & _
                                    lcOTROnombreTributo & "|" & lcOTROcodigoXitem & "|" & _
                                    lcOTROporcentaje & "|" & lcValorVentaConIGV & "|" & _
                                    lcValorVentaSinIGV & "|" & lcValorReferencialUnitario & "|"
                                    '
                                    act.Write (lcLinea & vbCrLf)   'Graba la linea
                                    
                                    oRsDetalleComprobante.MoveNext
                                Loop
                            End If
                         End If
                         If lcNotaCredito = "" Then
                            'CREANDO ARCHIVO LEY
                            lnImporteTotalDeLaVenta = CCur(lcImporteTotalDeLaVenta)
                            lcImporteSoles = Trim(sighentidades.Numlet(sighentidades.DevuelveNumeroSinDecimales(lnImporteTotalDeLaVenta)) & " SOLES " & _
                                             sighentidades.DevuelveSoloDecimales(lnImporteTotalDeLaVenta) & "/100 MN;")
                            If oRsBusquedaRecibos.Fields!idTipoComprobante = sghCajaTipoComprobante.sghCajaFactura Then  'FActura Sunat
                                Set oRsTmp877 = mo_ReglasFacturacion.ProveedoresSeleccionarPorRUC(oRsBusquedaRecibos!Ruc)
                                If oRsTmp877.RecordCount > 0 Then
                                   If Not IsNull(oRsTmp877!Direccion) Then
                                     lcImporteSoles = lcImporteSoles & UCase(oRsTmp877!Direccion)
                                   End If
                                End If
                                oRsTmp877.Close
                            End If
                            Set act = fso.CreateTextFile(lcNomComprobanteSunat & ".ley", True)
                            lcLinea = "1000|" & lcImporteSoles & "|"
                            act.Write (lcLinea)
                            'CREANDO ARCHIVO TRI
                            Set act = fso.CreateTextFile(lcNomComprobanteSunat & ".tri", True)
                            lcLinea = "1000|IGV|VAT|" & _
                                      IIf(lcIGVnombre = "IGV", lcTotalValorVenta & "|" & lcSumatoriaDeTributos, "0.00|0.00") & "|"
                            act.Write (lcLinea & vbCrLf)
                            lcLinea = "9997|EXO|VAT|" & _
                                       IIf(lcTotalVentasExon <> "", lcTotalVentasExon & "|0.00|", "0.00|0.00|")
                            act.Write (lcLinea & vbCrLf)
                            lcLinea = "9998|INA|FRE|" & _
                                       IIf(lcIGVnombre = "INA", lcTotalValorVenta & "|" & lcSumatoriaDeTributos, "0.00|0.00") & "|"
                            act.Write (lcLinea & vbCrLf)
                            If oRsBusquedaRecibos.Fields!idTipoComprobante <> sghCajaTipoComprobante.sghCajaFactura Then ' Factura Sunat
                                  lcLinea = "9996|GRA|FRE|" & _
                                            "0.00|0.00|"
                                            
                                  act.Write (lcLinea & vbCrLf)
                            End If
                            
                            'CREANDO ARCHIVO REL -  solo para boletas
                            If oRsBusquedaRecibos.Fields!idTipoComprobante <> sghCajaTipoComprobante.sghCajaFactura Then
                                  Set act = fso.CreateTextFile(lcNomComprobanteSunat & ".rel", True)
                                  lcLinea = ""
                                  act.Write (lcLinea)
                            End If
                          End If
                      Else
                          '****resumen diario BOLETA (solo genera RECORDSET y lo devuelve)*****
                          If lcFechaEmision <> "*" Then
                                TotalVentasInaf = "0.00"
                                SumatoriaIGV = "0.00"
                                Set oRsTmpBoletas = mo_AdminCaja.CajaComprobantesSeleccionarPorId(oRsBusquedaRecibos!IdComprobantePago, oConexion)
                                If oRsTmpBoletas.RecordCount > 0 Then
                                   If Not IsNull(oRsTmpBoletas!SunatTotalIgv) Then
                                        SumatoriaIGV = Format(oRsTmpBoletas!SunatTotalIgv, "#0.00")
                                   End If
                                   If Not IsNull(oRsTmpBoletas!SunatOpeInafectas) Then
                                        TotalVentasInaf = Format(oRsTmpBoletas!SunatOpeInafectas, "#0.00")
                                   End If
                                End If
                                oRsTmpBoletas.Close
                                lcTotalVentasExon = Format(CCur(oRsBusquedaRecibos.Fields!exoneraciones), "#0.00")
                                lcImporteTotalDeLaVenta = Format(CCur(oRsBusquedaRecibos.Fields!Total), "#0.00")
                                lbGrabaRD = True
                                If oRsBusquedaRecibos!exoneraciones > 0 Then
                                   If oRsBusquedaRecibos!Total = 0 Then
                                      TotalVentasInaf = "0.01"
                                      lcImporteTotalDeLaVenta = TotalVentasInaf
                                      SumatoriaIGV = "0.00"
                                      lbGrabaRD = False
                                   End If
                                End If
                                If lbGrabaRD = True Then
                                    oDoSunatResumenDia.FechaEmision = FechaEmision
                                    oDoSunatResumenDia.fechaResumen = FechaEmision  'lcFechaResumen
                                    oDoSunatResumenDia.tipoDcto = "03"  'tipo Boleta
                                    oDoSunatResumenDia.SerieDocumento = Trim(oRsBusquedaRecibos!NroSerie) & "-" & NumeroComprobante 'de la Boleta
                                    oDoSunatResumenDia.PacienteTipoDoc = TipoDocumSunat
                                    oDoSunatResumenDia.PacienteNumeroDoc = NroDocumSunat
                                    oDoSunatResumenDia.Moneda = TipoMoneda
                                    If Val(TotalVentasInaf) > 0 Then
                                       oDoSunatResumenDia.OpeGravadas = "0.00"
                                    Else
                                       oDoSunatResumenDia.OpeGravadas = Format(CCur(lcImporteTotalDeLaVenta) - CCur(SumatoriaIGV), "#0.00")
                                    End If
                                    oDoSunatResumenDia.OpeExoneradas = "0.00"   ' lcTotalVentasExon
                                    oDoSunatResumenDia.OpeInafectas = TotalVentasInaf
                                    oDoSunatResumenDia.OpeGratuitas = "0.00"
                                    oDoSunatResumenDia.OtrosCargos = "0.00"
                                    oDoSunatResumenDia.TotalIsc = "0.00"
                                    oDoSunatResumenDia.TotalIgv = SumatoriaIGV
                                    oDoSunatResumenDia.TotalOtros = "0.00"
                                    oDoSunatResumenDia.ImporteVenta = lcImporteTotalDeLaVenta
                                    oDoSunatResumenDia.ModTipoDcto = " "
                                    oDoSunatResumenDia.ModSerie = " "
                                    oDoSunatResumenDia.ModDocumento = " "
                                    oDoSunatResumenDia.PercepRegimen = " "
                                    oDoSunatResumenDia.PercepPorcen = " "
                                    oDoSunatResumenDia.PercepBaseImp = "0"
                                    oDoSunatResumenDia.PercepMonto = "0"
                                    oDoSunatResumenDia.PercepCobrar = "0"
                                    oDoSunatResumenDia.Estado = "1"
                                    If oSunatParteDia.Insertar(oDoSunatResumenDia) Then
                                    End If
                                End If
                          End If
                      End If
                End If
           End If
           oRsBusquedaRecibos.MoveNext
        Loop
    End If
    Set mo_AdminCaja = Nothing
    Set mo_ReglasAdmision = Nothing
    Set oDOCajaNroDocumento = Nothing
    Set oRsTmp877 = Nothing
    Set oRsTmp876 = Nothing
    Set oRsDetalleComprobante = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set oDoSunatResumenDia = Nothing
    Set oSunatParteDia = Nothing
    Set oConexion = Nothing
    Set oRsTmpBoletas = Nothing
End Sub



Sub ExportarFacturasBoletas(lcFechaInicial As String, lcFechaFinal As String, lcNroSerie As String, _
                            lcNroDocumento As String, lbUsaResumenDiario As Boolean, _
                            Optional oRsAnuladosComoActivos As Recordset)
    Dim lcVersionSunat As String
    lcVersionSunat = Trim(lcBuscaParametro.SeleccionaFilaParametro(573))
    If lcVersionSunat = "1.2" Then
       Dim lcCabecera10 As String, lcCabecera11 As String, lcCabecera13 As String
       Dim lcCabecera14 As String, lcCabecera15 As String
       ExportarFacturasBoletas12 lcFechaInicial, lcFechaFinal, lcNroSerie, _
                            lcNroDocumento, lcCabecera10, lcCabecera11, lcCabecera13, lcCabecera14, lcCabecera15, "", _
                            lbUsaResumenDiario, oRsAnuladosComoActivos
       Exit Sub
    ElseIf lcVersionSunat <> "1.05" Then
       Exit Sub
    End If
    '
    Dim Ruta As String
    Dim lcNomComprobanteSunat As String
    Dim lcLinea As String
    Dim Ruc As String
    Dim TipoComprobante As String
    Dim CodigoCobro As String
    Dim oRsDetalleComprobante As New Recordset
    Dim oRsBusquedaRecibos As Recordset
    Dim oDOCajaNroDocumento As DOCajaNroDocumento
    Ruta = lcBuscaParametro.SeleccionaFilaParametro(378)
    Dim fso As Object 'crea objeto para manejar archivos
    Dim act As Object 'crea objeto para manejar archivos
    Dim NumeroComprobante As String, lnIGV As Integer
    Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim lcServicioInafecto As String, wxParametro549 As String


    If lcNroSerie = "" And lcFechaInicial <> "" Then
       Set oRsBusquedaRecibos = mo_AdminCaja.CajaComprobantePagoFiltroPorNroSerieDocumentoOporRangoFemision("", "", CDate(lcFechaInicial), CDate(lcFechaFinal))
    ElseIf lcNroSerie <> "" Then
       Set oRsBusquedaRecibos = mo_AdminCaja.CajaComprobantePagoFiltroPorNroSerieDocumentoOporRangoFemision(lcNroSerie, lcNroDocumento, Date, Date)
    Else
       Set oRsBusquedaRecibos = oRsAnuladosComoActivos.Clone
    End If
    
    
    
    Ruc = Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(339)), 11)
    lnIGV = Val(lcBuscaParametro.SeleccionaFilaParametro(221))
    lcServicioInafecto = lcBuscaParametro.SeleccionaFilaParametro(537)
    wxParametro549 = lcBuscaParametro.SeleccionaFilaParametro(549)
    'DATOS CABECERA
    Dim TipoOperacion As String
    Dim FechaEmision As String
    Dim CodDomicilioFiscal As String
    Dim TipoDocumSunat As String
    Dim NroDocumSunat As String
    Dim NombreRazonSocial As String
    Dim TipoMoneda As String
    Dim DescGlobales As String
    Dim SumOtroCargos As String
    Dim TotalDescuentos As String
    Dim TotalVentasOper As String
    Dim TotalVentasInaf As String
    Dim TotalVentasExon As String
    Dim SumatoriaIGV As String
    Dim SumatoriaISC As String
    Dim SumatoriaOtrosTrib As String
    Dim ImporTotalVentas As String
    Dim lbProsigue As Boolean
    'Inicializando
    TipoOperacion = "01"
    FechaEmision = ""
    CodDomicilioFiscal = "0"
    TipoMoneda = "PEN"
    DescGlobales = "0.00"
    SumOtroCargos = "0.00"
    TotalDescuentos = "0.00"
    TotalVentasInaf = "0.00"
    TotalVentasExon = "0.00"
    SumatoriaISC = "0.00"
    SumatoriaOtrosTrib = "0.00"
    
    'DATOS DETALLE
    Dim UnidadMedida As String
    Dim CantidadUnids As String
    Dim CodigoProducto As String
    Dim CodigoProductoSunat As String
    Dim DescBienServicio As String
    Dim ValorUnitario As String
    Dim DescuItem As String
    Dim MontoIGV As String
    Dim AfectacionIGV As String
    Dim MontoISC As String
    Dim TipoISC As String
    Dim PrecioVentaUnitario As String
    Dim ValorVenta As String
    Dim DocumentoBaja As String, DescripcionBaja As String, CorrelativoBaja As String
    Dim lnFacturaSinIGV As Integer
    Dim lnPrecioUnitarioSinIgv As Double, lnMontoIGV As Double, lnTotalVentasOper As Double
    Dim oRsTmp876 As New Recordset, lcSql As String, lnImporteParcial As Double
    
    
    'Inicializando
    UnidadMedida = "NIU"
    TipoISC = "01"
    DescuItem = "0.00"
    
    MontoIGV = "0.00"
    MontoISC = "0.00"
    
    If oRsBusquedaRecibos.RecordCount > 0 Then
        oRsBusquedaRecibos.MoveFirst
        Do While Not oRsBusquedaRecibos.EOF
            AfectacionIGV = "10"
            lnIGV = Val(lcBuscaParametro.SeleccionaFilaParametro(221))
            lbProsigue = True
            'No se considera BOLETAS/FACTURAS de SERVICIOS
            If lcBuscaParametro.SeleccionaFilaParametro(532) = "S" And oRsBusquedaRecibos!IdTipoOrden = 1 Then
                lbProsigue = False
            End If
            '
            If lbProsigue = True Then
                 If oRsBusquedaRecibos!IdEstadoComprobante = 9 Then
                      '********************************Boleta/Factura anulada****************************
                      FechaEmision = Year(oRsBusquedaRecibos.Fields!fechaCobranza) & "-" & Right("00" & Month(oRsBusquedaRecibos.Fields!fechaCobranza), 2) & "-" & Right("00" & Day(oRsBusquedaRecibos.Fields!fechaCobranza), 2)
                      TipoComprobante = "RA"
                      If oRsBusquedaRecibos.Fields!idTipoComprobante = 2 Then ' Factura Sunat
                         TipoDocumSunat = "01" 'REG. UNICO DE CONTRIBUYENTES
                      Else
                         TipoDocumSunat = "03" 'boleta
                      End If
                      
                      DocumentoBaja = oRsBusquedaRecibos!NroSerie & "-" & _
                                      Right("0000000" & oRsBusquedaRecibos!nroDocumento, 8)
                      DescripcionBaja = "No se entregó al Cliente"
                      CorrelativoBaja = mo_AdminCaja.DevuelveCorrelativoDeFacturaBoletaAnulada(oRsBusquedaRecibos!NroSerie, oRsBusquedaRecibos!nroDocumento)
                      
                      lcNomComprobanteSunat = Ruta & "\" & Ruc & "-" & TipoComprobante & "-" & _
                                              Format(oRsBusquedaRecibos!fechaCobranza, "yyyymmdd") & _
                                              "-" & CorrelativoBaja
                      Set fso = CreateObject("scripting.filesystemobject") 'creo objeto de manejo de archivos
                      Set act = fso.CreateTextFile(lcNomComprobanteSunat & ".cba", True) 'creo el archivo con nombre asignado por parametro
                      lcLinea = FechaEmision & "|" & FechaEmision & "|" & TipoDocumSunat & "|" & _
                                DocumentoBaja & "|" & DescripcionBaja
                      act.Write (lcLinea) 'Graba la linea
                 
                 Else
                      '*******************************Boleta/Factura activas****************************
                      'debb-24/07/2018-inicio-chimbote
                      lnFacturaSinIGV = 99
                      If oRsBusquedaRecibos.Fields!idTipoComprobante = 2 Then  ' Factura
                         lnFacturaSinIGV = 0
                         Set oDOCajaNroDocumento = mo_AdminCaja.CajaNroDocumentoSeleccionarPorId(oRsBusquedaRecibos!idCaja, 2)
                         If Not oDOCajaNroDocumento Is Nothing Then
                            If oDOCajaNroDocumento.FacturaSinIGV = True Then
                               lnFacturaSinIGV = 1
                            End If
                         End If
                      End If
                      'debb-24/07/2018-fin
                      
                      
                      If oRsBusquedaRecibos.Fields!idTipoComprobante = 2 Then ' Factura Sunat
                          TipoComprobante = "01"
                          TipoDocumSunat = "6" 'REG. UNICO DE CONTRIBUYENTES
                          NroDocumSunat = oRsBusquedaRecibos.Fields!Ruc
                      Else
                          'oRsBusquedaRecibos.Fields!IdTipoComprobante = 3 Then 'Boleta Sunat
                          TipoComprobante = "03"
                          TipoDocumSunat = "1"
                          If IsNull(oRsBusquedaRecibos.Fields!idPaciente) Then
                              TipoDocumSunat = "1"                 'kike 2017
                              NroDocumSunat = "00000000"
                              If Not IsNull(oRsBusquedaRecibos!dni) Then
                                   If Len(Trim(oRsBusquedaRecibos!dni)) = 8 Then
                                      NroDocumSunat = oRsBusquedaRecibos!dni
                                   End If
                              End If
                          Else
                              Dim oDOPaciente As New DOPaciente
                              Set oDOPaciente = mo_ReglasAdmision.RetornaPacientesSeleccionarPorId(oRsBusquedaRecibos.Fields!idPaciente)
                              Select Case oDOPaciente.IdDocIdentidad
                              Case 1
                                  TipoDocumSunat = "1" 'DOC. NACIONAL DE IDENTIDAD
                              Case 2
                                  TipoDocumSunat = "4" 'CARNET DE EXTRANJERIA
                              Case 3
                                  TipoDocumSunat = "7" 'PASAPORTE
                              End Select
                              NroDocumSunat = Right("00000000" & oDOPaciente.nroDocumento, 8)
                              Set oDOPaciente = Nothing
                          End If
                      End If
                      NombreRazonSocial = Trim(Left(oRsBusquedaRecibos.Fields!razonSocial, 100))
                      NumeroComprobante = Right("00000000" & Trim(oRsBusquedaRecibos.Fields!nroDocumento), 8)
                      lcNomComprobanteSunat = Ruta & "\" & Ruc & "-" & TipoComprobante & "-" & _
                                             Trim(oRsBusquedaRecibos.Fields!NroSerie) & "-" & NumeroComprobante
                      
                      'CREANDO ARCHIVO CABECERA
                      Set fso = CreateObject("scripting.filesystemobject") 'creo objeto de manejo de archivos
                      Set act = fso.CreateTextFile(lcNomComprobanteSunat & ".cab", True) 'creo el archivo con nombre asignado por parametro
                      'debb-24/07/2018-inicio chimbote
                      'Boletas/Facturas de SERVICIO - INAFECTAS - JAMO
'                      If (lcServicioInafecto = "S" Or lcServicioInafecto = "SS") And oRsBusquedaRecibos!IdTipoOrden = 1 Then
'                         lnIGV = 0
'                      End If
                      If (lcServicioInafecto = "S" Or lcServicioInafecto = "SS") And oRsBusquedaRecibos!IdTipoOrden = 1 Then
                         lnIGV = 0
                      ElseIf lnFacturaSinIGV <> 99 Then
                         If lnFacturaSinIGV = 1 Then
                            lnIGV = 0
                         End If
                      End If
                      'debb-24/07/2018-fin
                      If oRsBusquedaRecibos.Fields!exoneraciones > 0 Then
                            lnTotalVentasOper = Round(oRsBusquedaRecibos.Fields!exoneraciones / ((100 + lnIGV) / 100), 2)
                            FechaEmision = Year(oRsBusquedaRecibos.Fields!fechaCobranza) & "-" & Right("00" & Month(oRsBusquedaRecibos.Fields!fechaCobranza), 2) & "-" & Right("00" & Day(oRsBusquedaRecibos.Fields!fechaCobranza), 2)
                            TotalVentasOper = Format(lnTotalVentasOper, "#0.00")
                            TotalVentasExon = Format(CCur(oRsBusquedaRecibos!exoneraciones), "#0.00")
                            SumatoriaIGV = Format(Round(lnTotalVentasOper * lnIGV / 100, 2), "#0.00")
                            ImporTotalVentas = Format(oRsBusquedaRecibos!Total, "#0.00")
                      Else
                            lnTotalVentasOper = Round(oRsBusquedaRecibos.Fields!Total / ((100 + lnIGV) / 100), 2)
                            FechaEmision = Year(oRsBusquedaRecibos.Fields!fechaCobranza) & "-" & Right("00" & Month(oRsBusquedaRecibos.Fields!fechaCobranza), 2) & "-" & Right("00" & Day(oRsBusquedaRecibos.Fields!fechaCobranza), 2)
                            TotalVentasOper = Format(lnTotalVentasOper, "#0.00")
                            TotalVentasExon = Format(CCur(oRsBusquedaRecibos.Fields!exoneraciones), "#0.00")
                            SumatoriaIGV = Format(Round(lnTotalVentasOper * lnIGV / 100, 2), "#0.00")
                            ImporTotalVentas = Format(lnTotalVentasOper + CCur(SumatoriaIGV), "#0.00")
                      End If
                      If (lcServicioInafecto = "S" Or lcServicioInafecto = "SS") And oRsBusquedaRecibos!IdTipoOrden = 1 Then
                         TotalVentasOper = Format(0, "#0.00")
                         TotalVentasInaf = Format(oRsBusquedaRecibos!Total, "#0.00")
                      End If
                      '
                      lcLinea = TipoOperacion & "|" & FechaEmision & "|" & _
                                CodDomicilioFiscal & "|" & TipoDocumSunat & "|" & _
                                NroDocumSunat & "|" & NombreRazonSocial & "|" & _
                                TipoMoneda & "|" & DescGlobales & "|" & _
                                SumOtroCargos & "|" & TotalDescuentos & "|" & _
                                TotalVentasOper & "|" & TotalVentasInaf & "|" & _
                                TotalVentasExon & "|" & SumatoriaIGV & "|" & _
                                SumatoriaISC & "|" & SumatoriaOtrosTrib & "|" & ImporTotalVentas & "|"
                      act.Write (lcLinea) 'Graba la linea
                               
                      'CREANDO ARCHIVO DETALLE
                      Set act = fso.CreateTextFile(lcNomComprobanteSunat & ".det", True) 'creo el archivo con nombre asignado por parametro
                      
                      If oRsBusquedaRecibos.Fields!IdTipoOrden = 2 Or oRsBusquedaRecibos.Fields!IdTipoOrden = 3 Then 'Bien-Farmacia
                          Set oRsDetalleComprobante = mo_AdminCaja.CajaComprobantePagoProductosPorIdComprobante(oRsBusquedaRecibos.Fields!IdComprobantePago)
                          If oRsDetalleComprobante.RecordCount > 0 Then
                              oRsDetalleComprobante.MoveFirst
                              Do While Not oRsDetalleComprobante.EOF
                                  lnImporteParcial = oRsDetalleComprobante!PrecioUnitario * oRsDetalleComprobante!Cantidad
                                  lnPrecioUnitarioSinIgv = Round((oRsDetalleComprobante!PrecioUnitario / ((100 + lnIGV) / 100)), 10)
                                  lnMontoIGV = Round(lnPrecioUnitarioSinIgv * lnIGV / 100, 2)
                                  CantidadUnids = Format(CCur(oRsDetalleComprobante.Fields!Cantidad), "#0")
                                  CodigoProducto = oRsDetalleComprobante.Fields!codigo
                                  DescBienServicio = oRsDetalleComprobante.Fields!nombreProducto
                                  ValorUnitario = Format(CCur(lnPrecioUnitarioSinIgv), "#0.0000000000")
                                  MontoIGV = Format(CCur(lnMontoIGV), "#0.00")
                                  PrecioVentaUnitario = Format(lnPrecioUnitarioSinIgv + lnMontoIGV, "#0.00")
                                  ValorVenta = Format(Round(oRsDetalleComprobante!Cantidad * lnPrecioUnitarioSinIgv, 2), "#0.00")
                                  'kike 2017
                                  '
                                  lcLinea = UnidadMedida & "|" & CantidadUnids & "|" & _
                                  CodigoProducto & "|" & CodigoProductoSunat & "|" & _
                                  DescBienServicio & "|" & ValorUnitario & "|" & _
                                  DescuItem & "|" & MontoIGV & "|" & _
                                  AfectacionIGV & "|" & MontoISC & "|" & _
                                  TipoISC & "|" & PrecioVentaUnitario & "|" & ValorVenta & "|"
                                  act.Write (lcLinea & vbCrLf)    'Graba la linea
                                  
                                  oRsDetalleComprobante.MoveNext
                              Loop
                          End If
                      Else '1 Servicio
                          Set oRsDetalleComprobante = mo_AdminCaja.CajaComprobantePagoServiciosPorIdComprobante(oRsBusquedaRecibos.Fields!IdComprobantePago)
                          If oRsDetalleComprobante.RecordCount > 0 Then
                              oRsDetalleComprobante.MoveFirst
                              Do While Not oRsDetalleComprobante.EOF
                                  'Boletas/Facturas de SERVICIO - INAFECTAS - JAMO
                                  If lcServicioInafecto = "S" And oRsBusquedaRecibos!IdTipoOrden = 1 Then
                                     AfectacionIGV = "30"
                                     lnIGV = 0
                                  End If
                                  '
                                  CantidadUnids = Format(CCur(oRsDetalleComprobante.Fields!Cantidad), "#0")
                                  CodigoProducto = oRsDetalleComprobante.Fields!codigo
                                  If oRsDetalleComprobante!idProducto = Val(wxParametro549) Then
                                     DescBienServicio = oRsDetalleComprobante!Observaciones
                                  Else
                                     DescBienServicio = oRsDetalleComprobante.Fields!nombreProducto
                                  End If
                                  ValorUnitario = Format(CCur(oRsDetalleComprobante.Fields!PrecioUnitario), "#0.00")
                                  MontoIGV = "0.00"
                                  PrecioVentaUnitario = Format(CCur(oRsDetalleComprobante.Fields!TotalPorPagar), "#0.00")
                                  ValorVenta = Format(CCur(oRsDetalleComprobante.Fields!TotalPorPagar), "#0.00")
                                  '
                                  PrecioVentaUnitario = Format(Round(oRsDetalleComprobante.Fields!TotalPorPagar * 100 / (lnIGV + 100), 2), "#0.00")
                                  MontoIGV = Format((oRsDetalleComprobante.Fields!TotalPorPagar - CCur(PrecioVentaUnitario)), "#0.00")

                                  lcLinea = UnidadMedida & "|" & CantidadUnids & "|" & _
                                  CodigoProducto & "|" & CodigoProductoSunat & "|" & _
                                  DescBienServicio & "|" & ValorUnitario & "|" & _
                                  DescuItem & "|" & MontoIGV & "|" & _
                                  AfectacionIGV & "|" & MontoISC & "|" & _
                                  TipoISC & "|" & PrecioVentaUnitario & "|" & ValorVenta & "|"
                                  act.Write (lcLinea & vbCrLf)   'Graba la linea
                                  
                                  oRsDetalleComprobante.MoveNext
                              Loop
                          End If
                     End If
                End If
           End If
           oRsBusquedaRecibos.MoveNext
        Loop
    End If
    Set mo_AdminCaja = Nothing
    Set mo_ReglasAdmision = Nothing
    Set oDOCajaNroDocumento = Nothing
End Sub

Function DevuelveTextoParaQRsullana(oRsBusquedaRecibos As Recordset) As String
On Error Resume Next
Dim TipoComprobante  As String, TipoDocumSunat As String, NroDocumSunat As String, Ruc As String
Dim NombreRazonSocial  As String, NumeroComprobante As String, lcServicioInafecto As String
Dim lnTotalVentasOper As Double, FechaEmision As String, TotalVentasOper As String, TotalVentasExon As String
Dim SumatoriaIGV As String, ImporTotalVentas As String, TotalVentasInaf As String
Dim lcNroSerie As String, lcNroDocumento As String, lnIGV As Integer
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
        lcNroSerie = oRsBusquedaRecibos!NroSerie
        lcNroDocumento = oRsBusquedaRecibos!nroDocumento
        Ruc = Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(339)), 11)
        lcServicioInafecto = lcBuscaParametro.SeleccionaFilaParametro(537)
        lnIGV = Val(lcBuscaParametro.SeleccionaFilaParametro(221))
        If oRsBusquedaRecibos.Fields!idTipoComprobante = 2 Then ' Factura Sunat
            TipoComprobante = "01"
            TipoDocumSunat = "6" 'REG. UNICO DE CONTRIBUYENTES
            NroDocumSunat = oRsBusquedaRecibos.Fields!Ruc
        Else
            'oRsBusquedaRecibos.Fields!IdTipoComprobante = 3 Then 'Boleta Sunat
            TipoComprobante = "03"
            TipoDocumSunat = "1"
            If IsNull(oRsBusquedaRecibos.Fields!idPaciente) Then
                TipoDocumSunat = "1"                 'kike 2017
                NroDocumSunat = "00000000"
                If Not IsNull(oRsBusquedaRecibos!dni) Then
                     If Len(Trim(oRsBusquedaRecibos!dni)) = 8 Then
                        NroDocumSunat = oRsBusquedaRecibos!dni
                     End If
                End If
            Else
                Dim oDOPaciente As New DOPaciente
                Set oDOPaciente = mo_ReglasAdmision.RetornaPacientesSeleccionarPorId(oRsBusquedaRecibos.Fields!idPaciente)
                Select Case oDOPaciente.IdDocIdentidad
                Case 1
                    TipoDocumSunat = "1" 'DOC. NACIONAL DE IDENTIDAD
                Case 2
                    TipoDocumSunat = "4" 'CARNET DE EXTRANJERIA
                Case 3
                    TipoDocumSunat = "7" 'PASAPORTE
                End Select
                NroDocumSunat = Right("00000000" & oDOPaciente.nroDocumento, 8)
                Set oDOPaciente = Nothing
            End If
        End If
        NombreRazonSocial = Trim(Left(oRsBusquedaRecibos.Fields!razonSocial, 100))
        NumeroComprobante = Right("00000000" & Trim(oRsBusquedaRecibos.Fields!nroDocumento), 8)
'        lcNomComprobanteSunat = Ruta & "\" & Ruc & "-" & TipoComprobante & "-" & _
'                               Trim(oRsBusquedaRecibos.Fields!nroSerie) & "-" & NumeroComprobante
        
        
        'Boletas/Facturas de SERVICIO - INAFECTAS - JAMO
        If (lcServicioInafecto = "S" Or lcServicioInafecto = "SS") And oRsBusquedaRecibos!IdTipoOrden = 1 Then
           lnIGV = 0
        End If
        
        If oRsBusquedaRecibos.Fields!exoneraciones > 0 Then
              lnTotalVentasOper = Round(oRsBusquedaRecibos.Fields!exoneraciones / ((100 + lnIGV) / 100), 2)
              FechaEmision = Year(oRsBusquedaRecibos.Fields!fechaCobranza) & "-" & Right("00" & Month(oRsBusquedaRecibos.Fields!fechaCobranza), 2) & "-" & Right("00" & Day(oRsBusquedaRecibos.Fields!fechaCobranza), 2)
              TotalVentasOper = Format(lnTotalVentasOper, "#0.00")
              TotalVentasExon = Format(CCur(oRsBusquedaRecibos!exoneraciones), "#0.00")
              SumatoriaIGV = Format(Round(lnTotalVentasOper * lnIGV / 100, 2), "#0.00")
              ImporTotalVentas = Format(oRsBusquedaRecibos!TotalBoleta, "#0.00")
        Else
              lnTotalVentasOper = Round(oRsBusquedaRecibos!TotalBoleta / ((100 + lnIGV) / 100), 2)
              FechaEmision = Year(oRsBusquedaRecibos.Fields!fechaCobranza) & "-" & Right("00" & Month(oRsBusquedaRecibos.Fields!fechaCobranza), 2) & "-" & Right("00" & Day(oRsBusquedaRecibos.Fields!fechaCobranza), 2)
              TotalVentasOper = Format(lnTotalVentasOper, "#0.00")
              TotalVentasExon = Format(CCur(oRsBusquedaRecibos.Fields!exoneraciones), "#0.00")
              SumatoriaIGV = Format(Round(lnTotalVentasOper * lnIGV / 100, 2), "#0.00")
              ImporTotalVentas = Format(lnTotalVentasOper + CCur(SumatoriaIGV), "#0.00")
        End If
        If (lcServicioInafecto = "S" Or lcServicioInafecto = "SS") And oRsBusquedaRecibos!IdTipoOrden = 1 Then
        
           TotalVentasOper = Format(0, "#0.00")
           TotalVentasInaf = Format(oRsBusquedaRecibos!TotalBoleta, "#0.00")
        End If
        '
        DevuelveTextoParaQRsullana = Ruc & "|" & TipoComprobante & "|" & _
                                     lcNroSerie & "|" & lcNroDocumento & "|" & _
                                     SumatoriaIGV & "|" & ImporTotalVentas & "|" & _
                                     FechaEmision & "|" & TipoDocumSunat & "|" & _
                                     NroDocumSunat & "|"
    Set mo_ReglasAdmision = Nothing
End Function


Function SeleccionarPorFechaEmision(lcFechaInicio As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
   oConexion.CommandTimeout = 900
   oConexion.CursorLocation = adUseClient
   oConexion.Open sighentidades.CadenaConexion
   Set SeleccionarPorFechaEmision = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = oConexion
     .CommandText = "Sunat_ResumenDiarioSeleccionarPorFechaEmision"
     Set oParameter = .CreateParameter("@FechaResumen", adDBTimeStamp, adParamInput, 8, CDate(lcFechaInicio)): .Parameters.Append oParameter
     Set oRecordset = .Execute
      Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorFechaEmision = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Sub ExportarResumenDiarioYanulaciones12(lbYaTieneHistoricoDeRD As Boolean, lcFechaInicio As String, _
                                        oRsBusquedaRecibos As Recordset, oRsTmpDctos As Recordset)
    Dim oDoSunatResumenDia As New DoSunatResumenDia
    Dim oSunatParteDia As New SunatParteDia
    Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
    Dim oConexion As New Connection
    Dim lcNroSerie As String, lcNroDocumento As String, FechaEmision As String, TipoComprobante As String
    Dim TipoDocumSunat As String, DocumentoBaja As String, DescripcionBaja As String, CorrelativoBaja  As String
    Dim fso As Object 'crea objeto para manejar archivos
    Dim act As Object 'crea objeto para manejar archivos
    Dim lcNomComprobanteSunat  As String, lcLinea  As String, Ruc As String, Ruta As String
    Dim CorrelativoAnualResumenDiario As String, ldHoy As Date, lcDctoRDI As String
    ldHoy = lcBuscaParametro.RetornaFechaHoraServidorSQL
    Ruta = lcBuscaParametro.SeleccionaFilaParametro(378)
    Ruc = Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(339)), 11)
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    Set oSunatParteDia.Conexion = oConexion
    If oRsTmpDctos.RecordCount > 0 Then
       'Genera archivo plano RESUMEN DIARIO
       If lbYaTieneHistoricoDeRD = False Then
          CorrelativoAnualResumenDiario = mo_AdminCaja.sunat_ResumenDiarioUltimoCorrelativo(oConexion)
          lcNomComprobanteSunat = Ruta & "\" & Ruc & "-RC-" & Format(CDate(lcFechaInicio), "yyyymmdd") & "-" & _
                                  CorrelativoAnualResumenDiario & ".RDI"
          lcDctoRDI = Ruc & "-RC-" & Format(CDate(lcFechaInicio), "yyyymmdd") & "-" & _
                      CorrelativoAnualResumenDiario
       Else
          lcNomComprobanteSunat = Ruta & "\" & oRsTmpDctos!DctoRDI & ".RDI"
       End If
       Set fso = CreateObject("scripting.filesystemobject") 'creo objeto de manejo de archivos
       Set act = fso.CreateTextFile(lcNomComprobanteSunat, True) 'creo el archivo con nombre asignado por parametro
       oRsTmpDctos.MoveFirst
       Do While Not oRsTmpDctos.EOF
          lcLinea = Trim(oRsTmpDctos!FechaEmision) & "|" & Trim(oRsTmpDctos!fechaResumen) & "|" & _
                  Trim(oRsTmpDctos!tipoDcto) & "|" & Trim(oRsTmpDctos!SerieDocumento) & "|" & _
                  Trim(oRsTmpDctos!PacienteTipoDoc) & "|" & Trim(oRsTmpDctos!PacienteNumeroDoc) & "|" & _
                  Trim(oRsTmpDctos!Moneda) & "|" & Trim(oRsTmpDctos!OpeGravadas) & "|" & _
                  Trim(oRsTmpDctos!OpeExoneradas) & "|" & Trim(oRsTmpDctos!OpeInafectas) & "|" & _
                  Trim(oRsTmpDctos!OpeGratuitas) & "|" & Trim(oRsTmpDctos!OtrosCargos) & "|" & _
                  Trim(oRsTmpDctos!TotalIsc) & "|" & Trim(oRsTmpDctos!TotalIgv) & "|" & _
                  Trim(oRsTmpDctos!TotalOtros) & "|" & Trim(oRsTmpDctos!ImporteVenta) & "|" & _
                  Trim(oRsTmpDctos!ModTipoDcto) & "|" & Trim(oRsTmpDctos!ModSerie) & "|" & _
                  Trim(oRsTmpDctos!ModDocumento) & "|" & Trim(oRsTmpDctos!PercepRegimen) & "|" & _
                  Trim(oRsTmpDctos!PercepPorcen) & "|" & Trim(oRsTmpDctos!PercepBaseImp) & "|" & _
                  Trim(oRsTmpDctos!PercepMonto) & "|" & Trim(oRsTmpDctos!PercepCobrar) & "|" & _
                  Trim(oRsTmpDctos!Estado) & "|"
          If lbYaTieneHistoricoDeRD = False Then
             oDoSunatResumenDia.ID = oRsTmpDctos!ID
             If oSunatParteDia.SeleccionarPorId(oDoSunatResumenDia) Then    'modifica por ID
                oDoSunatResumenDia.DctoRDI = lcDctoRDI
                If oSunatParteDia.Modificar(oDoSunatResumenDia) Then   'modifica por ID
                   act.Write (lcLinea & vbCrLf)    'Graba la linea
                End If
             End If
          Else
             act.Write (lcLinea & vbCrLf)    'Graba la linea
          End If
          oRsTmpDctos.MoveNext
       Loop
       'Genera archivo plano ANULADOS DEL DIA
'       oRsTmpDctos.Filter = "Estado='3'"
'       If oRsTmpDctos.RecordCount > 0 Then
'           If oRsBusquedaRecibos.State = 0 Then
'              Set oRsBusquedaRecibos = mo_AdminCaja.CajaComprobantePagoFiltroPorNroSerieDocumentoOporRangoFemision("", "", CDate(lcFechaInicio), CDate(Left(lcFechaInicio, 10) & " 23:59:59"))
'           End If
'           oRsTmpDctos.MoveFirst
'           Do While Not oRsTmpDctos.EOF
'                lcNroSerie = Trim(Left(oRsTmpDctos!SerieDocumento, InStr(oRsTmpDctos!SerieDocumento, "-") - 1))
'                lcNroDocumento = Trim(Mid(oRsTmpDctos!SerieDocumento, InStr(oRsTmpDctos!SerieDocumento, "-") + 1, 15))
'                oRsBusquedaRecibos.Filter = "nroSerie='" & lcNroSerie & "' and nroDocumento='" & lcNroDocumento & "'"
'                If oRsBusquedaRecibos.RecordCount > 0 Then
'                    FechaEmision = Year(oRsBusquedaRecibos.Fields!fechaCobranza) & "-" & _
'                                   Right("00" & Month(oRsBusquedaRecibos.Fields!fechaCobranza), 2) & "-" & _
'                                   Right("00" & Day(oRsBusquedaRecibos.Fields!fechaCobranza), 2)
'                    TipoComprobante = "RA"
'                    If oRsBusquedaRecibos.Fields!IdTipoComprobante = 2 Then ' Factura Sunat
'                       TipoDocumSunat = "01" 'REG. UNICO DE CONTRIBUYENTES
'                    Else
'                       TipoDocumSunat = "03" 'boleta
'                    End If
'                    DocumentoBaja = oRsBusquedaRecibos!NroSerie & "-" & _
'                                    Right("0000000" & oRsBusquedaRecibos!nroDocumento, 8)
'                    DescripcionBaja = "No se entregó al Cliente"
'                    CorrelativoBaja = mo_AdminCaja.DevuelveCorrelativoDeFacturaBoletaAnulada(oRsBusquedaRecibos!NroSerie, oRsBusquedaRecibos!nroDocumento)
'
'                    lcNomComprobanteSunat = Ruta & "\" & Ruc & "-" & TipoComprobante & "-" & _
'                                            Format(oRsBusquedaRecibos!fechaCobranza, "yyyymmdd") & _
'                                            "-" & CorrelativoBaja
'                    Set fso = CreateObject("scripting.filesystemobject") 'creo objeto de manejo de archivos
'                    Set act = fso.CreateTextFile(lcNomComprobanteSunat & ".cba", True) 'creo el archivo con nombre asignado por parametro
''                    lcLinea = FechaEmision & "|" & oRsTmpDctos!fechaResumen & "|" & TipoDocumSunat & "|" & _
''                              DocumentoBaja & "|" & DescripcionBaja
'                    lcLinea = FechaEmision & "|" & FechaEmision & "|" & TipoDocumSunat & "|" & _
'                              DocumentoBaja & "|" & DescripcionBaja
'                    act.Write (lcLinea) 'Graba la linea
'                End If
'                oRsTmpDctos.MoveNext
'           Loop
'       End If
    End If
    oConexion.Close
    Set oDoSunatResumenDia = Nothing
    Set oSunatParteDia = Nothing
    Set mo_AdminCaja = Nothing
    Set oConexion = Nothing
End Sub



Sub ExportarNotasCredito12(lcFechaInicial As String, lcFechaFinal As String, lcNroSerie As String, lcNroDocumento As String, _
                           lbUsaResumenDiario As Boolean)
    Dim Ruta As String
    Dim lcNomComprobanteSunat As String
    Dim lcLinea As String
    Dim Ruc As String
    Dim TipoComprobante As String
    Dim CodigoCobro As String
    Dim oRsBusquedaRecibos As Recordset
    Dim oRsDetalleComprobante As Recordset
    Ruta = lcBuscaParametro.SeleccionaFilaParametro(378)
    Dim fso As Object 'crea objeto para manejar archivos
    Dim act As Object 'crea objeto para manejar archivos
    Dim NumeroComprobante As String
    Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim oDoSunatResumenDia As New DoSunatResumenDia
    Dim oSunatParteDia As New SunatParteDia
    Dim oConexion As New Connection
    Dim oRsDetalleProductosConLotes As New Recordset
    Dim lcFechaEmision As String
    Dim lcFechaResumen As String, ldHoy As Date
    ldHoy = CDate(lcBuscaParametro.RetornaFechaServidorSQL)
    lcFechaResumen = Year(ldHoy) & "-" & Right("00" & Month(ldHoy), 2) & "-" & Right("00" & Day(ldHoy), 2)
    lcFechaEmision = "*"
    If lcNroSerie = "" Then
       Set oRsBusquedaRecibos = mo_AdminCaja.NotaCreditoParaSunatPorNumYFecha(CDate(lcFechaInicial), CDate(lcFechaFinal), "", "")
       lcFechaEmision = Year(CDate(lcFechaInicial)) & "-" & Right("00" & Month(CDate(lcFechaInicial)), 2) & "-" & Right("00" & Day(CDate(lcFechaInicial)), 2)
    Else
       Set oRsBusquedaRecibos = mo_AdminCaja.NotaCreditoParaSunatPorNumYFecha(Date, Date, lcNroSerie, lcNroDocumento)
    End If
    'solo ICA
    If Val(lcBuscaParametro.SeleccionaFilaParametro(280)) = 3359 Then
       oRsBusquedaRecibos.Filter = "idCaja<>2"
    End If
    '
    Ruc = Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(339)), 11)
    
    'DATOS CABECERA
    Dim FechaEmision As String
    Dim CodMotivoTipoNota As String 'Catalogo 10
    Dim DescMotivoTipoNota As String 'Motivo
    Dim TipoMoneda As String
    Dim TipoCompModif As String
    Dim NroComprModif As String
    Dim TipoDocumSunat As String
    Dim NroDocumSunat As String
    Dim NombreRazonSocial As String
    Dim SumOtroCargos As String
    Dim TotalVentasOper As String
    Dim TotalVentasInaf As String
    Dim TotalVentasExon As String
    Dim SumatoriaIGV As String
    Dim SumatoriaISC As String
    Dim SumatoriaOtrosTrib As String
    Dim ImporTotalVentas As String
    'Inicializando
    TipoComprobante = "07" 'NOTA DE CREDITO
    FechaEmision = ""
    CodMotivoTipoNota = "03"
    DescMotivoTipoNota = "Penalidades/ otros conceptos"
    TipoMoneda = "PEN"
    SumOtroCargos = "0.00"
    TotalVentasInaf = "0.00"
    TotalVentasExon = "0.00"
    SumatoriaISC = "0.00"
    SumatoriaOtrosTrib = "0.00"
    
    'DATOS DETALLE
    Dim UnidadMedida As String
    Dim CantidadUnids As String
    Dim CodigoProducto As String
    Dim CodigoProductoSunat As String
    Dim DescBienServicio As String
    Dim ValorUnitario As String
    Dim DescuItem As String
    Dim MontoIGV As String
    Dim AfectacionIGV As String
    Dim MontoISC As String
    Dim TipoISC As String
    Dim PrecioVentaUnitario As String
    Dim ValorVenta As String
    Dim lbProsigue As Boolean
    Dim lnImporteParcial As Double, lnIGV As Double, lnPrecioUnitarioSinIgv As Double, lnMontoIGV As Double
    'Inicializando
    UnidadMedida = "NIU"
    TipoISC = "01"
    DescuItem = "0.00"
    AfectacionIGV = "10"
    MontoIGV = "0.00"
    MontoISC = "0.00"
    lnIGV = Val(lcBuscaParametro.SeleccionaFilaParametro(221))
    Dim lcTotalVentasExon As String, lcImporteTotalDeLaVenta As String
    Dim TipoOperacion As String, CodDomicilioFiscal As String, lcImporteTotalVenta As String
    Dim lcCabecera10 As String, lcCabecera11 As String, lcCabecera13 As String
    Dim lcCabecera14 As String, lcCabecera15 As String, lcTotalPrecioVenta As String
    Dim lcVersionUBL As String, lcCustomizacionDocumento As String, lcHoraEmision As String, lbNoGeneraResumenDiario As Boolean
    Dim oRsTmpBoletas As New Recordset
    TipoOperacion = "0101"
    CodDomicilioFiscal = "0000"
    lcVersionUBL = "2.1"
    lcCustomizacionDocumento = "2.0"
    
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    Set oSunatParteDia.Conexion = oConexion
    
    If oRsBusquedaRecibos.RecordCount > 0 Then
        oRsBusquedaRecibos.MoveFirst
        Do While Not oRsBusquedaRecibos.EOF
            lnIGV = Val(lcBuscaParametro.SeleccionaFilaParametro(221))
            'No se considera BOLETAS/FACTURAS de SERVICIOS
            lbProsigue = True
            If lcBuscaParametro.SeleccionaFilaParametro(532) = "S" And IsNull(oRsBusquedaRecibos!IdFarmacia) Then
                lbProsigue = False
            End If
            '
            If lbProsigue = True Then
                If oRsBusquedaRecibos.Fields!idTipoComprobante = sghCajaTipoComprobante.sghCajaFactura Then
                    TipoCompModif = "01"
                    TipoDocumSunat = "6" 'REG. UNICO DE CONTRIBUYENTES
                    NroDocumSunat = oRsBusquedaRecibos.Fields!RucComprAfec
                Else
                    TipoCompModif = "03"
                    TipoDocumSunat = "1"
                    If IsNull(oRsBusquedaRecibos.Fields!IdPacienteComprAfec) Then
                        TipoDocumSunat = "1"
                        NroDocumSunat = "00000000"
                        'debb-05/12/2017 inicio
                        NombreRazonSocial = "SELECT dni from CajaComprobantesPago where  idComprobantePago=" & oRsBusquedaRecibos!IdComprobantePago
                        oRsDetalleProductosConLotes.Open NombreRazonSocial, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
                        If oRsDetalleProductosConLotes.RecordCount > 0 Then
                           If Not IsNull(oRsDetalleProductosConLotes!dni) Then
                              If Len(Trim(oRsDetalleProductosConLotes!dni)) = 8 And Val(oRsDetalleProductosConLotes!dni) > 0 Then
                                 TipoDocumSunat = "1"
                                 NroDocumSunat = Trim(oRsDetalleProductosConLotes!dni)
                              End If
                           End If
                        End If
                        oRsDetalleProductosConLotes.Close
                        'debb-05/12/2017  fin
                        
                    Else
                        Dim oDOPaciente As New DOPaciente
                        Set oDOPaciente = mo_ReglasAdmision.RetornaPacientesSeleccionarPorId(oRsBusquedaRecibos.Fields!IdPacienteComprAfec)
                        Select Case oDOPaciente.IdDocIdentidad
                        Case 1
                            TipoDocumSunat = "1" 'DOC. NACIONAL DE IDENTIDAD
                        Case 2
                            TipoDocumSunat = "4" 'CARNET DE EXTRANJERIA
                        Case 3
                            TipoDocumSunat = "7" 'PASAPORTE
                        End Select
                        NroDocumSunat = Right("00000000" & oDOPaciente.nroDocumento, 8)
                        Set oDOPaciente = Nothing
                    End If
                End If
                NroComprModif = Trim(oRsBusquedaRecibos.Fields!serieComprAfectado) & "-" & Right("00000000" & Trim(oRsBusquedaRecibos.Fields!NroComprAfectado), 8)
                NombreRazonSocial = Trim(Left(oRsBusquedaRecibos.Fields!RazonSocialComprAfec, 100))
                NumeroComprobante = Right("00000000" & Trim(oRsBusquedaRecibos.Fields!nroDocumento), 8)
                lcNomComprobanteSunat = Ruta & "\" & Ruc & "-" & TipoComprobante & "-" & _
                Trim(oRsBusquedaRecibos.Fields!NroSerie) & "-" & NumeroComprobante
                
                
                'KIKE2019
                lbNoGeneraResumenDiario = True
                If lbUsaResumenDiario = True Then
                     If oRsBusquedaRecibos.Fields!idTipoComprobante <> sghCajaTipoComprobante.sghCajaFactura Then
                        lbNoGeneraResumenDiario = False
                     End If
                End If
                If lbNoGeneraResumenDiario = True Then
                    'generando DET/TRI/REL/LEY
                    lcCabecera10 = ""
                    lcCabecera11 = ""
                    lcCabecera13 = ""
                    lcCabecera14 = ""
                    lcCabecera15 = ""
                    ExportarFacturasBoletas12 "", "", oRsBusquedaRecibos!serieComprAfectado, oRsBusquedaRecibos!NroComprAfectado, _
                                              lcCabecera10, lcCabecera11, lcCabecera13, lcCabecera14, lcCabecera15, _
                                              lcNomComprobanteSunat, False
                    lcHoraEmision = Format(oRsBusquedaRecibos.Fields!FechaPagado, sighentidades.DevuelveHoraSoloFormato_HMS)
                    lcTotalPrecioVenta = Format(CCur(lcCabecera10) + CCur(lcCabecera11), "#0.00")
                    lcImporteTotalVenta = Format(CCur(lcTotalPrecioVenta) - CCur(lcCabecera13) + CCur(lcCabecera14) - CCur(lcCabecera15), "#0.00")
                    
                    'CREANDO ARCHIVO CABECERA NOT
                    Set fso = CreateObject("scripting.filesystemobject") 'creo objeto de manejo de archivos
                    Set act = fso.CreateTextFile(lcNomComprobanteSunat & ".not", True) 'creo el archivo con nombre asignado por parametro
                    
                    FechaEmision = Year(oRsBusquedaRecibos.Fields!FechaPagado) & "-" & Right("00" & Month(oRsBusquedaRecibos.Fields!FechaPagado), 2) & "-" & Right("00" & Day(oRsBusquedaRecibos.Fields!FechaPagado), 2)
                    TotalVentasOper = Format(CCur(oRsBusquedaRecibos.Fields!Subtotal), "#0.00")
                    SumatoriaIGV = Format(CCur(oRsBusquedaRecibos.Fields!IGV), "#0.00")
                    ImporTotalVentas = Format(CCur(oRsBusquedaRecibos.Fields!Total), "#0.00")
                                            
                                                       
                    lcLinea = TipoOperacion & "|" & FechaEmision & "|" & _
                              lcHoraEmision & "|" & CodDomicilioFiscal & "|" & _
                              TipoDocumSunat & "|" & NroDocumSunat & "|" & _
                              NombreRazonSocial & "|" & TipoMoneda & "|" & _
                              CodMotivoTipoNota & "|" & DescMotivoTipoNota & "|" & _
                              TipoCompModif & "|" & NroComprModif & "|" & _
                              lcCabecera10 & "|" & lcCabecera11 & "|" & _
                              lcTotalPrecioVenta & "|" & lcCabecera13 & "|" & _
                              lcCabecera14 & "|" & lcCabecera15 & "|" & _
                              lcImporteTotalVenta & "|" & lcVersionUBL & "|" & _
                              lcCustomizacionDocumento
                    act.Write (lcLinea) 'Graba la linea
                Else
                    '****resumen diario NC (agregar)*****
                    If lcFechaEmision <> "*" Then
                        TotalVentasInaf = "0.00"
                        SumatoriaIGV = "0.00"
                        lcTotalVentasExon = "0.00"
                        Set oRsTmpBoletas = mo_AdminCaja.CajaComprobantesSeleccionarPorId(oRsBusquedaRecibos!IdComprobantePago, oConexion)
                        If oRsTmpBoletas.RecordCount > 0 Then
                            If Not IsNull(oRsTmpBoletas!SunatTotalIgv) Then
                               SumatoriaIGV = Format(oRsTmpBoletas!SunatTotalIgv, "#0.00")
                            End If
                            If Not IsNull(oRsTmpBoletas!SunatOpeInafectas) Then
                               TotalVentasInaf = Format(oRsTmpBoletas!SunatOpeInafectas, "#0.00")
                            End If
                            If Not IsNull(oRsTmpBoletas!exoneraciones) Then
                               lcTotalVentasExon = Format(CCur(oRsTmpBoletas!exoneraciones), "#0.00")
                            End If
                        End If
                        oRsTmpBoletas.Close
                        
                        lcImporteTotalDeLaVenta = Format(CCur(oRsBusquedaRecibos.Fields!Total), "#0.00")
                        FechaEmision = Year(oRsBusquedaRecibos.Fields!FechaPagado) & "-" & Right("00" & Month(oRsBusquedaRecibos.Fields!FechaPagado), 2) & "-" & Right("00" & Day(oRsBusquedaRecibos.Fields!FechaPagado), 2)
                        ImporTotalVentas = Format(CCur(oRsBusquedaRecibos.Fields!Total), "#0.00")
                        oDoSunatResumenDia.FechaEmision = FechaEmision
                        oDoSunatResumenDia.fechaResumen = FechaEmision 'lcFechaResumen
                        oDoSunatResumenDia.tipoDcto = "07"  'tipo nc
                        oDoSunatResumenDia.SerieDocumento = Trim(oRsBusquedaRecibos!NroSerie) & "-" & NumeroComprobante 'de la NC
                        oDoSunatResumenDia.PacienteTipoDoc = TipoDocumSunat
                        oDoSunatResumenDia.PacienteNumeroDoc = NroDocumSunat
                        oDoSunatResumenDia.Moneda = TipoMoneda
                        oDoSunatResumenDia.OpeGravadas = Format(CCur(lcImporteTotalDeLaVenta) - CCur(SumatoriaIGV), "#0.00")
                        oDoSunatResumenDia.OpeExoneradas = lcTotalVentasExon
                        oDoSunatResumenDia.OpeInafectas = TotalVentasInaf
                        oDoSunatResumenDia.OpeGratuitas = "0.00"
                        oDoSunatResumenDia.OtrosCargos = "0.00"
                        oDoSunatResumenDia.TotalIsc = "0.00"
                        oDoSunatResumenDia.TotalIgv = SumatoriaIGV
                        oDoSunatResumenDia.TotalOtros = "0.00"
                        oDoSunatResumenDia.ImporteVenta = lcImporteTotalDeLaVenta
                        oDoSunatResumenDia.ModTipoDcto = "03"
                        oDoSunatResumenDia.ModSerie = oRsBusquedaRecibos!serieComprAfectado
                        oDoSunatResumenDia.ModDocumento = oRsBusquedaRecibos!NroComprAfectado
                        oDoSunatResumenDia.PercepRegimen = " "
                        oDoSunatResumenDia.PercepPorcen = " "
                        oDoSunatResumenDia.PercepBaseImp = "0"
                        oDoSunatResumenDia.PercepMonto = "0"
                        oDoSunatResumenDia.PercepCobrar = "0"
                        oDoSunatResumenDia.Estado = "1"
                        If oSunatParteDia.Insertar(oDoSunatResumenDia) Then
                        End If
                    End If
                End If
           End If
           oRsBusquedaRecibos.MoveNext
        Loop
    End If
    oConexion.Close
    Set mo_AdminCaja = Nothing
    Set mo_ReglasAdmision = Nothing
    Set oRsDetalleProductosConLotes = Nothing
    Set oDoSunatResumenDia = Nothing
    Set oSunatParteDia = Nothing
    Set oConexion = Nothing
    Set oRsTmpBoletas = Nothing
End Sub
Sub ExportarNotasCredito(lcFechaInicial As String, lcFechaFinal As String, lcNroSerie As String, lcNroDocumento As String, _
                         lbUsaResumenDiario As Boolean)
    Dim lcVersionSunat As String
    lcVersionSunat = Trim(lcBuscaParametro.SeleccionaFilaParametro(573))
    If lcVersionSunat = "1.2" Then
       ExportarNotasCredito12 lcFechaInicial, lcFechaFinal, lcNroSerie, lcNroDocumento, lbUsaResumenDiario
       Exit Sub
    ElseIf lcVersionSunat <> "1.05" Then
       Exit Sub
    End If
    '
    Dim Ruta As String
    Dim lcNomComprobanteSunat As String
    Dim lcLinea As String
    Dim Ruc As String
    Dim TipoComprobante As String
    Dim CodigoCobro As String
    Dim oRsBusquedaRecibos As Recordset
    Dim oRsDetalleComprobante As Recordset
    Ruta = lcBuscaParametro.SeleccionaFilaParametro(378)
    Dim fso As Object 'crea objeto para manejar archivos
    Dim act As Object 'crea objeto para manejar archivos
    Dim NumeroComprobante As String
    Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim oRsDetalleProductosConLotes As New Recordset
'    Dim fs, f
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    Set f = fs.CreateFolder("c:\ruta_carpeta\")
    If lcNroSerie = "" Then
       Set oRsBusquedaRecibos = mo_AdminCaja.NotaCreditoParaSunatPorNumYFecha(CDate(lcFechaInicial), CDate(lcFechaFinal), "", "")
    Else
       Set oRsBusquedaRecibos = mo_AdminCaja.NotaCreditoParaSunatPorNumYFecha(Date, Date, lcNroSerie, lcNroDocumento)
    End If
    Ruc = Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(339)), 11)
    
    'DATOS CABECERA
    Dim FechaEmision As String
    Dim CodMotivoTipoNota As String 'Catalogo 10
    Dim DescMotivoTipoNota As String 'Motivo
    Dim TipoMoneda As String
    Dim TipoCompModif As String
    Dim NroComprModif As String
    Dim TipoDocumSunat As String
    Dim NroDocumSunat As String
    Dim NombreRazonSocial As String
    Dim SumOtroCargos As String
    Dim TotalVentasOper As String
    Dim TotalVentasInaf As String
    Dim TotalVentasExon As String
    Dim SumatoriaIGV As String
    Dim SumatoriaISC As String
    Dim SumatoriaOtrosTrib As String
    Dim ImporTotalVentas As String
    'Inicializando
    TipoComprobante = "07" 'NOTA DE CREDITO
    FechaEmision = ""
    CodMotivoTipoNota = "01"
    DescMotivoTipoNota = "Anulación de la operación"
    TipoMoneda = "PEN"
    SumOtroCargos = "0.00"
    TotalVentasInaf = "0.00"
    TotalVentasExon = "0.00"
    SumatoriaISC = "0.00"
    SumatoriaOtrosTrib = "0.00"
    
    'DATOS DETALLE
    Dim UnidadMedida As String
    Dim CantidadUnids As String
    Dim CodigoProducto As String
    Dim CodigoProductoSunat As String
    Dim DescBienServicio As String
    Dim ValorUnitario As String
    Dim DescuItem As String
    Dim MontoIGV As String
    Dim AfectacionIGV As String
    Dim MontoISC As String
    Dim TipoISC As String
    Dim PrecioVentaUnitario As String
    Dim ValorVenta As String
    Dim lbProsigue As Boolean
    Dim lnImporteParcial As Double, lnIGV As Double, lnPrecioUnitarioSinIgv As Double, lnMontoIGV As Double
    'Inicializando
    UnidadMedida = "NIU"
    TipoISC = "01"
    DescuItem = "0.00"
    AfectacionIGV = "10"
    MontoIGV = "0.00"
    MontoISC = "0.00"
    lnIGV = Val(lcBuscaParametro.SeleccionaFilaParametro(221))
    If oRsBusquedaRecibos.RecordCount > 0 Then
        oRsBusquedaRecibos.MoveFirst
        Do While Not oRsBusquedaRecibos.EOF
            lnIGV = Val(lcBuscaParametro.SeleccionaFilaParametro(221))
            'No se considera BOLETAS/FACTURAS de SERVICIOS
            lbProsigue = True
            If lcBuscaParametro.SeleccionaFilaParametro(532) = "S" And IsNull(oRsBusquedaRecibos!IdFarmacia) Then
                lbProsigue = False
            End If
            '
            If lbProsigue = True Then
                If oRsBusquedaRecibos.Fields!idTipoComprobante = 2 Then ' Factura Sunat
                    TipoCompModif = "01"
                    TipoDocumSunat = "6" 'REG. UNICO DE CONTRIBUYENTES
                    NroDocumSunat = oRsBusquedaRecibos.Fields!RucComprAfec
                Else
                    'oRsBusquedaRecibos.Fields!IdTipoComprobante = 3 Then 'Boleta Sunat
                    TipoCompModif = "03"
                    TipoDocumSunat = "1"
                    If IsNull(oRsBusquedaRecibos.Fields!IdPacienteComprAfec) Then
                        TipoDocumSunat = "1"
                        NroDocumSunat = "00000000"
                        'debb-05/12/2017 inicio
                        NombreRazonSocial = "SELECT dni from CajaComprobantesPago where  idComprobantePago=" & oRsBusquedaRecibos!IdComprobantePago
                        oRsDetalleProductosConLotes.Open NombreRazonSocial, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
                        If oRsDetalleProductosConLotes.RecordCount > 0 Then
                           If Not IsNull(oRsDetalleProductosConLotes!dni) Then
                              If Len(Trim(oRsDetalleProductosConLotes!dni)) = 8 And Val(oRsDetalleProductosConLotes!dni) > 0 Then
                                 TipoDocumSunat = "1"
                                 NroDocumSunat = Trim(oRsDetalleProductosConLotes!dni)
                              End If
                           End If
                        End If
                        oRsDetalleProductosConLotes.Close
                        'debb-05/12/2017  fin
                        
                    Else
                        Dim oDOPaciente As New DOPaciente
                        Set oDOPaciente = mo_ReglasAdmision.RetornaPacientesSeleccionarPorId(oRsBusquedaRecibos.Fields!IdPacienteComprAfec)
                        Select Case oDOPaciente.IdDocIdentidad
                        Case 1
                            TipoDocumSunat = "1" 'DOC. NACIONAL DE IDENTIDAD
                        Case 2
                            TipoDocumSunat = "4" 'CARNET DE EXTRANJERIA
                        Case 3
                            TipoDocumSunat = "7" 'PASAPORTE
                        End Select
                        NroDocumSunat = Right("00000000" & oDOPaciente.nroDocumento, 8)
                        Set oDOPaciente = Nothing
                    End If
                End If
                NroComprModif = Trim(oRsBusquedaRecibos.Fields!serieComprAfectado) & "-" & Right("00000000" & Trim(oRsBusquedaRecibos.Fields!NroComprAfectado), 8)
                NombreRazonSocial = Trim(Left(oRsBusquedaRecibos.Fields!RazonSocialComprAfec, 100))
                NumeroComprobante = Right("00000000" & Trim(oRsBusquedaRecibos.Fields!nroDocumento), 8)
                lcNomComprobanteSunat = Ruta & "\" & Ruc & "-" & TipoComprobante & "-" & _
                Trim(oRsBusquedaRecibos.Fields!NroSerie) & "-" & NumeroComprobante
                
                'CREANDO ARCHIVO CABECERA NOT
                Set fso = CreateObject("scripting.filesystemobject") 'creo objeto de manejo de archivos
                Set act = fso.CreateTextFile(lcNomComprobanteSunat & ".not", True) 'creo el archivo con nombre asignado por parametro
                
                FechaEmision = Year(oRsBusquedaRecibos.Fields!FechaPagado) & "-" & Right("00" & Month(oRsBusquedaRecibos.Fields!FechaPagado), 2) & "-" & Right("00" & Day(oRsBusquedaRecibos.Fields!FechaPagado), 2)
                TotalVentasOper = Format(CCur(oRsBusquedaRecibos.Fields!Subtotal), "#0.00")
                SumatoriaIGV = Format(CCur(oRsBusquedaRecibos.Fields!IGV), "#0.00")
                ImporTotalVentas = Format(CCur(oRsBusquedaRecibos.Fields!Total), "#0.00")
                                        
                lcLinea = FechaEmision & "|" & CodMotivoTipoNota & "|" & _
                          DescMotivoTipoNota & "|" & TipoCompModif & "|" & _
                          NroComprModif & "|" & TipoDocumSunat & "|" & _
                          NroDocumSunat & "|" & NombreRazonSocial & "|" & TipoMoneda & "|" & _
                          SumOtroCargos & "|" & ImporTotalVentas & "|" & _
                          TotalVentasInaf & "|" & TotalVentasExon & "|" & _
                          SumatoriaIGV & "|" & SumatoriaISC & "|" & _
                          SumatoriaOtrosTrib & "|" & ImporTotalVentas & "|"
                act.Write (lcLinea) 'Graba la linea
                                                    
                'CREANDO ARCHIVO DETALLE
                Set act = fso.CreateTextFile(lcNomComprobanteSunat & ".det", True) 'creo el archivo con nombre asignado por parametro
    
                If oRsBusquedaRecibos.Fields!IdTipoOrden = 2 Or oRsBusquedaRecibos.Fields!IdTipoOrden = 3 Then 'Bien-Farmacia
                    Set oRsDetalleComprobante = mo_AdminCaja.CajaComprobantePagoProductosPorIdComprobante(oRsBusquedaRecibos.Fields!IdComprobantePago)
                    If oRsDetalleComprobante.RecordCount > 0 Then
                        oRsDetalleComprobante.MoveFirst
                        Do While Not oRsDetalleComprobante.EOF
                            lnImporteParcial = oRsDetalleComprobante!PrecioUnitario * oRsDetalleComprobante!Cantidad
                            lnPrecioUnitarioSinIgv = Round((oRsDetalleComprobante!PrecioUnitario / ((100 + lnIGV) / 100)), 10)
                            lnMontoIGV = Round(lnPrecioUnitarioSinIgv * lnIGV / 100, 2)
                            CantidadUnids = Format(CCur(oRsDetalleComprobante.Fields!Cantidad), "#0")
                            CodigoProducto = oRsDetalleComprobante.Fields!codigo
                            DescBienServicio = oRsDetalleComprobante.Fields!nombreProducto
                            ValorUnitario = Format(CCur(lnPrecioUnitarioSinIgv), "#0.0000000000")
                            MontoIGV = Format(CCur(lnMontoIGV), "#0.00")
                            PrecioVentaUnitario = Format(lnPrecioUnitarioSinIgv + lnMontoIGV, "#0.00")
                            ValorVenta = Format(Round(oRsDetalleComprobante!Cantidad * lnPrecioUnitarioSinIgv, 2), "#0.00")
'                            CantidadUnids = Format(CCur(ORsDetalleComprobante.Fields!Cantidad), "#0")
'                            CodigoProducto = ORsDetalleComprobante.Fields!Codigo
'                            DescBienServicio = ORsDetalleComprobante.Fields!nombreProducto
'                            ValorUnitario = Format(CCur(ORsDetalleComprobante.Fields!PrecioUnitario), "#0.00")
'                            MontoIGV = "0.00"
'                            PrecioVentaUnitario = Format(CCur(ORsDetalleComprobante.Fields!TotalPorPagar), "#0.00")
'                            ValorVenta = Format(CCur(ORsDetalleComprobante.Fields!TotalPorPagar), "#0.00")
                            lcLinea = UnidadMedida & "|" & CantidadUnids & "|" & _
                                       CodigoProducto & "|" & CodigoProductoSunat & "|" & _
                                       DescBienServicio & "|" & ValorUnitario & "|" & _
                                       DescuItem & "|" & MontoIGV & "|" & _
                                       AfectacionIGV & "|" & MontoISC & "|" & _
                                       TipoISC & "|" & PrecioVentaUnitario & "|" & ValorVenta & "|"
                            act.Write (lcLinea & vbCrLf) 'Graba la linea
                            oRsDetalleComprobante.MoveNext
                        Loop
                    End If
                Else '1 Servicio
                    Set oRsDetalleComprobante = mo_AdminCaja.CajaComprobantePagoServiciosPorIdComprobante(oRsBusquedaRecibos.Fields!IdComprobantePago)
                    If oRsDetalleComprobante.RecordCount > 0 Then
                        oRsDetalleComprobante.MoveFirst
                        Do While Not oRsDetalleComprobante.EOF
                            CantidadUnids = Format(CCur(oRsDetalleComprobante.Fields!Cantidad), "#0")
                            CodigoProducto = oRsDetalleComprobante.Fields!codigo
                            DescBienServicio = oRsDetalleComprobante.Fields!nombreProducto
                            ValorUnitario = Format(CCur(oRsDetalleComprobante.Fields!PrecioUnitario), "#0.00")
                            MontoIGV = "0.00"
                            PrecioVentaUnitario = Format(CCur(oRsDetalleComprobante.Fields!TotalPorPagar), "#0.00")
                            ValorVenta = Format(CCur(oRsDetalleComprobante.Fields!TotalPorPagar), "#0.00")
                            '
                            PrecioVentaUnitario = Format(Round(oRsDetalleComprobante.Fields!TotalPorPagar * 100 / (lnIGV + 100), 2), "#0.00")
                            MontoIGV = Format((oRsDetalleComprobante.Fields!TotalPorPagar - CCur(PrecioVentaUnitario)), "#0.00")
'                            CantidadUnids = Format(CCur(ORsDetalleComprobante.Fields!Cantidad), "#0")
'                            CodigoProducto = ORsDetalleComprobante.Fields!Codigo
'                            DescBienServicio = ORsDetalleComprobante.Fields!nombreProducto
'                            ValorUnitario = Format(CCur(ORsDetalleComprobante.Fields!PrecioUnitario), "#0.00")
'                            MontoIGV = "0.00"
'                            PrecioVentaUnitario = Format(CCur(ORsDetalleComprobante.Fields!TotalPorPagar), "#0.00")
'                            ValorVenta = Format(CCur(ORsDetalleComprobante.Fields!TotalPorPagar), "#0.00")
                            lcLinea = UnidadMedida & "|" & CantidadUnids & "|" & _
                                       CodigoProducto & "|" & CodigoProductoSunat & "|" & _
                                       DescBienServicio & "|" & ValorUnitario & "|" & _
                                       DescuItem & "|" & MontoIGV & "|" & _
                                       AfectacionIGV & "|" & MontoISC & "|" & _
                                       TipoISC & "|" & PrecioVentaUnitario & "|" & ValorVenta & "|"
                            act.Write (lcLinea & vbCrLf) 'Graba la linea
                            oRsDetalleComprobante.MoveNext
                        Loop
                    End If
                End If
                
               
           End If
           oRsBusquedaRecibos.MoveNext
        Loop
    End If
    Set mo_AdminCaja = Nothing
    Set mo_ReglasAdmision = Nothing
    Set oRsDetalleProductosConLotes = Nothing



    
End Sub


Sub HistoriaWebEliminar(lcDNI As String)
    On Error GoTo HWE
    Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_Pacientes As New DOPaciente
    Dim oPacientes As New Pacientes
    Dim oDoSunasaPacientesHistoricos As New DoSunasaPacientesHistoricos
    Dim oConexion As New Connection
    sighentidades.AbreConexionSIGH oConexion
    Set oPacientes.Conexion = oConexion
    If oPacientes.PacientesSeleccionarPorDNI(lcDNI, mo_Pacientes) Then
        If mo_Pacientes.Observacion = wxObservacionCrearPaciente Then
            If mo_AdminFacturacion.PacienteSePuedeEliminar(mo_Pacientes.idPaciente) Then
               Set oDoSunasaPacientesHistoricos = mo_AdminAdmision.SunasaPacientesHistoricosSeleccionarPorIdPaciente(mo_Pacientes.idPaciente, oConexion)
               mo_Pacientes.IdUsuarioAuditoria = sighentidades.Usuario
               If mo_AdminAdmision.PacientesEliminar(mo_Pacientes, (500 + 183), sighentidades.RetornaNombrePC, _
                  Trim(str(mo_Pacientes.NroHistoriaClinica)) & " : " & Trim(mo_Pacientes.ApellidoPaterno) & " " & _
                  Trim(mo_Pacientes.ApellidoMaterno) & " " & Trim(mo_Pacientes.PrimerNombre) & " " & _
                  mo_Pacientes.SegundoNombre, oDoSunasaPacientesHistoricos) Then
                  
               End If
            End If
        End If
    End If
HWE:
    oConexion.Close
    Set mo_AdminFacturacion = Nothing
    Set mo_AdminAdmision = Nothing
    Set mo_Pacientes = Nothing
    Set oPacientes = Nothing
    Set oDoSunasaPacientesHistoricos = Nothing
    Set oConexion = Nothing
End Sub
Function HistoriaWebCrearIgualAlDNIsoloSIS(lcDNI As String, oConexion As Connection, oRsBuscaPacientesSis As Recordset, _
                                ldHoy As Date, wxParametro351 As String) As Boolean
    On Error GoTo errCHC
    HistoriaWebCrearIgualAlDNIsoloSIS = False
    Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_Pacientes As New DOPaciente
    Dim oPacientes As New Pacientes
    Dim mo_Historia As New DOHistoriaClinica
    Dim oDoSunasaPacientesHistoricos As New DoSunasaPacientesHistoricos
    Dim rspacientes As New Recordset
    Dim lnDni1 As Long
    If wxParametro351 = "S" Then
'GLCC 02/11/20 CAMBIO36 INICIO
'ANTERIOR:        lnDni1 = Val(wxNueve & lcDNI)
            lnDni1 = Val(lcDNI)
'GLCC 02/11/20 CAMBIO36 FIN
    Else
        lnDni1 = Val(lcDNI)
    End If
    With mo_Pacientes
        .idPaciente = 0
        .ApellidoMaterno = IIf(IsNull(oRsBuscaPacientesSis!ApMaterno) Or oRsBuscaPacientesSis!ApMaterno = "", sighentidades.DevuelveSinApellido, oRsBuscaPacientesSis!ApMaterno)
        .ApellidoPaterno = IIf(IsNull(oRsBuscaPacientesSis!ApPaterno) Or oRsBuscaPacientesSis!ApPaterno = "", sighentidades.DevuelveSinApellido, oRsBuscaPacientesSis!ApPaterno)
        .FechaNacimiento = oRsBuscaPacientesSis!Fnacimiento
        .IdDistritoDomicilio = IIf(IsNull(oRsBuscaPacientesSis!distritoDomicilio), 0, Val(oRsBuscaPacientesSis!distritoDomicilio))
        .IdDocIdentidad = 1
        .idTipoNumeracion = 2  'manual
        .idTipoSexo = IIf(oRsBuscaPacientesSis!Sexo = 0, 2, 1)
        .IdUsuarioAuditoria = sighentidades.Usuario
        .nroDocumento = lcDNI
        .NroHistoriaClinica = lnDni1
        .Observacion = wxObservacionCrearPaciente
        .PrimerNombre = oRsBuscaPacientesSis!PNombre
        .SegundoNombre = IIf(IsNull(oRsBuscaPacientesSis!sNombre), "", oRsBuscaPacientesSis!sNombre)
        .Autogenerado = mo_AdminAdmision.PacienteCrearNroAutogenerado(mo_Pacientes)
        .IdEtnia = "58"     'mestizo
        .IdIdioma = 101   'español
    End With
    With mo_Historia
        .fechaCreacion = Format(ldHoy, sighentidades.DevuelveFechaSoloFormato_DMY)
        .FechaPasoAPasivo = 0
        .IdEstadoHistoria = 1
        .IdTipoHistoria = 1
        .idTipoNumeracion = 2   'manual
        .IdUsuarioAuditoria = sighentidades.Usuario
        .NroHistoriaClinica = lnDni1
    End With
    With oDoSunasaPacientesHistoricos
         .YaNoTieneSeguro = False
         .IdUsuarioAuditoria = sighentidades.Usuario
    End With
    
    Set oPacientes.Conexion = oConexion
    Set rspacientes = oPacientes.ObtenerConElMismoAutogenerado(mo_Pacientes)
    If rspacientes.RecordCount = 0 Then
        If mo_AdminAdmision.PacientesAgregarPacienteEHistoriaClinica(mo_Pacientes, mo_Historia, _
                       (500 + 183), sighentidades.RetornaNombrePC, _
                       Trim(mo_Pacientes.ApellidoPaterno) & " " & Trim(mo_Pacientes.ApellidoMaterno) & " " & Trim(mo_Pacientes.PrimerNombre) & " " & mo_Pacientes.SegundoNombre, _
                       oDoSunasaPacientesHistoricos) = True Then
            HistoriaWebCrearIgualAlDNIsoloSIS = True
        End If
    Else
        mo_Pacientes.idPaciente = rspacientes!idPaciente
        If oPacientes.SeleccionarPorId(mo_Pacientes) Then
           mo_Pacientes.IdDocIdentidad = 1
           mo_Pacientes.nroDocumento = lcDNI
           If oPacientes.Modificar(mo_Pacientes, False) Then
              HistoriaWebCrearIgualAlDNIsoloSIS = True
           End If
        End If
    End If
    rspacientes.Close
errCHC:
    Set mo_AdminAdmision = Nothing
    Set mo_Pacientes = Nothing
    Set mo_Historia = Nothing
    Set oDoSunasaPacientesHistoricos = Nothing
    Set rspacientes = Nothing
     
End Function




Sub EnviaEmail(lcParametro524 As String, lcParametro523 As String, lcAsunto As String, lcAdjunto As String, _
               lcEmailPara As String, lcMensaje As String)
        Dim oMail As New SIGHNegocios.clsCDOmail
        With oMail
             'datos para enviar
            .servidor = "smtp.gmail.com"
            .puerto = 465
            .UseAuntentificacion = True
            .ssl = True
            .Usuario = lcParametro524
            .Password = lcParametro523
            .Asunto = lcAsunto
            .Adjunto = lcAdjunto     '"G:\barrantes\libro1.pdf"
            .de = lcParametro524
            .para = lcEmailPara
            .mensaje = lcMensaje
            .Enviar_Backup 'manda el mail
        End With
        Set oMail = Nothing
End Sub

Sub MensajeCelularEnviar(mo_Pacientes As DOPaciente, lnIdCuentaAtencion As Long, lcMensaje As String, _
                        lcOpcion As String, oConexion As Connection, Optional lcCelular As Object)
     On Error GoTo ErrEnvMensaje
     Dim oDOsms_mensajes As New DOsms_mensajes
     Dim oSms_mensajes As New sms_mensajes
     Dim lcCelular1 As String
     lcCelular1 = ""
     If lcCelular Is Nothing Then
     Else
        lcCelular1 = lcCelular
     End If
     If (Len(mo_Pacientes.telefono) = 9 And Left(mo_Pacientes.telefono, 1) = "9") Or lcCelular1 <> "" Then
        Set oSms_mensajes.Conexion = oConexion
        oDOsms_mensajes.celular = IIf(lcCelular1 <> "", lcCelular1, mo_Pacientes.telefono)
        oDOsms_mensajes.fechaCreacion = lcBuscaParametro.RetornaFechaHoraServidorSQL
        'oDOsms_mensajes.FechaEnvio
        oDOsms_mensajes.idCuentaAtencion = lnIdCuentaAtencion
        oDOsms_mensajes.idEstadoEnvio = 1
        oDOsms_mensajes.idPaciente = mo_Pacientes.idPaciente
        oDOsms_mensajes.IdUsuarioAuditoria = sighentidades.Usuario
        oDOsms_mensajes.mensaje = Left(lcMensaje, 140)
        oDOsms_mensajes.Opcion = lcOpcion
        If oSms_mensajes.Insertar(oDOsms_mensajes) = False Then
           GoTo ErrEnvMensaje
        End If
     End If
ErrEnvMensaje:
     Set oDOsms_mensajes = Nothing
     Set oSms_mensajes = Nothing
     Exit Sub
     Resume
End Sub

Sub MensajeCelularEliminaMensajesHistoricos(oConexion As Connection)
     On Error GoTo ErrEnvMensaje1
     Dim oDOsms_mensajes As New DOsms_mensajes
     Dim oSms_mensajes As New sms_mensajes
     Set oSms_mensajes.Conexion = oConexion
     oDOsms_mensajes.fechaCreacion = CDate(lcBuscaParametro.RetornaFechaHoraServidorSQL) - 60   'menores a 1 meses
     If oSms_mensajes.EliminarMenoresAfecha(oDOsms_mensajes) = False Then
        GoTo ErrEnvMensaje1
     End If
ErrEnvMensaje1:
     Set oDOsms_mensajes = Nothing
     Set oSms_mensajes = Nothing
            
End Sub




Sub ImportaCitasWebMINSA(ml_idUsuario As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String)
     '
     On Error GoTo ErrImpCWEB
     Dim iFila As Long, lnIdWeb As Long
     Dim lineas() As String, lnRegistros As Long, Registro As String, Fichero As Integer, Campos() As String, Columna As Single
     Dim ldFecha As Date, lnIdServicio As Long, lnIdMedico As Long, lcHoraInicio As String
     Dim lcHoraFinal As String, lnIdEstadoCitaWeb As Long, lnIdCitaBloqueada As Long, lcDNI As String
     Dim lcApellidoPaterno As String, lcApellidoMaterno As String, lcPrimerNombre As String
     Dim ldFechaConfirmacion As Date, lcHoraConfirmacion As String, lnNroRegistro As Long
     Dim lcSegundoNombre As String, lnIdTipoSexo As Long, ldFechaNacimiento As Date, lnUbigeo As Long
     Dim lnDiasMaximoConfirmacion As Integer, lnDiasQfaltaPcita As Integer, lnIdFuenteFinanciamiento As Long
     Dim ldFechaHoy As Date, lcHoraHoy As String, lnHistoriaCreada As Long, lnIdAtencion As Long
     Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
     Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
     Dim mo_ReglasServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
     Dim oCitasWebCupos As New CitasWebCupos, oDOCitasWebCupos As New DOCitasWebCupos
     Dim oConexionExterna As New Connection, oConexion As New Connection
     Dim oRsTmp1 As New Recordset, lbImportaCitasWeb As Boolean
     Dim oRsEmailMensajes As New Recordset
     Dim lbContinua86 As Boolean
     Dim oRsTmp86 As New Recordset
     Dim mo_atenciones As New DOAtencion, oAtencion As New Atenciones
     Dim mo_Pacientes As New DOPaciente, oPacientes As New Pacientes
     Dim oMensajeCelular As New SIGHProxies.Procesos
     Dim lcMensajeCelular As String, wxParametro205 As String
     Dim lnIdEstablecimientoOrigen As Long, lcNroReferenciaOrigen As String, lcReferenciaMedicoOColeg As String
     Dim lnReferenciaOidDiagnostico As Long, lcReferenciaOservicio As String
     Dim lcParametro524 As String, lcParametro523 As String
     
     
     lbImportaCitasWeb = False
     '
     wxParametro205 = lcBuscaParametro.SeleccionaFilaParametro(205)
     lcParametro524 = lcBuscaParametro.SeleccionaFilaParametro(524)
     lcParametro523 = lcBuscaParametro.SeleccionaFilaParametro(523)
     '
     oConexion.CommandTimeout = 300
     oConexion.CursorLocation = adUseClient
     oConexion.Open sighentidades.CadenaConexion
     '
     oConexionExterna.CommandTimeout = 300
     oConexionExterna.CursorLocation = adUseClient
     oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
     oConexionExterna.BeginTrans
     Set oCitasWebCupos.Conexion = oConexionExterna
     '
     lnDiasMaximoConfirmacion = Val(lcBuscaParametro.SeleccionaFilaParametro(321))
     ldFechaHoy = CDate(lcBuscaParametro.RetornaFechaServidorSQL): lcHoraHoy = lcBuscaParametro.RetornaHoraServidorSQL
            
    lcSql = "select * from Reporte_cabecera where idUsuario=" & sighentidades.Usuario
    oRsEmailMensajes.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
            
            
    '*************************************** Citas Web MINSA ***********************************
    Set oRsTmp1 = mo_ReglasDeProgMedica.CitasWebCuposSeleccionarPorIdEstadoCita(sghCitaWebEstados.CupoDisponibleEnCitaWeb, oConexionExterna)
    If oRsTmp1.RecordCount > 0 Then
       Set oAtencion.Conexion = oConexion
       mo_atenciones.IdUsuarioAuditoria = sighentidades.Usuario
       Set oPacientes.Conexion = oConexion
       mo_Pacientes.IdUsuarioAuditoria = sighentidades.Usuario
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
            If oRsTmp1!idCitaBloqueada > 0 And ldFechaHoy = oRsTmp1!fecha Then
                    'elimina cupo WEB y libera CITA BLOQUEADA
                    oDOCitasWebCupos.idCitaBloqueada = oRsTmp1!idCitaBloqueada
                    oDOCitasWebCupos.IdUsuarioAuditoria = 0
                    If oCitasWebCupos.EliminarPorIdCitaBloqueada(oDOCitasWebCupos) = False Then
                       MsgBox oCitasWebCupos.MensajeError: GoTo ErrImpCWEB
                    End If
                    If mo_ReglasDeProgMedica.CitasBloqueadasEliminarPorId(oRsTmp1!idCitaBloqueada, oConexion) = False Then
                        If mo_ReglasDeProgMedica.MensajeError = "" Then
                           ms_MensajeError = "No se pudo ELIMINAR CITAS BLOQUEDAS WEB DEL DIA"
                        End If
                        MsgBox mo_ReglasDeProgMedica.MensajeError: GoTo ErrImpCWEB
                    End If
            End If
            oRsTmp1.MoveNext
       Loop
    End If
    oRsTmp1.Close
    '
    Set oRsTmp1 = mo_ReglasDeProgMedica.CitasWebCuposSeleccionarPorIdEstadoCita(sghCitaWebEstados.CupoConfirmadoEnCitaWeb, oConexionExterna)
    If oRsTmp1.RecordCount > 0 Then
       Set oAtencion.Conexion = oConexion
       mo_atenciones.IdUsuarioAuditoria = sighentidades.Usuario
       Set oPacientes.Conexion = oConexion
       mo_Pacientes.IdUsuarioAuditoria = sighentidades.Usuario
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
            If Not IsNull(oRsTmp1!dni) And oRsTmp1!idCitaBloqueada > 0 And ldFechaHoy <= oRsTmp1!fecha Then
                'chequea q el paciente no tenga cita ese DIA en el mismo CONSULTORIO
                'EN CASO haya sacado cita en el mismo EESS
                lbContinua86 = True
                Set oRsTmp86 = mo_ReglasAdmision.AtencionesSeleccionarPorIdPaciente(oRsTmp1!idPaciente, sghTipoServicio.sghConsultaExterna)
                If oRsTmp86.RecordCount > 0 Then
                   oRsTmp86.MoveFirst
                   Do While Not oRsTmp86.EOF
                      If Format(oRsTmp86!FechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY) = Format(oRsTmp1!fecha, sighentidades.DevuelveFechaSoloFormato_DMY) And _
                                                                 oRsTmp86!IdServicioIngreso = oRsTmp1!idServicio Then
                         lbContinua86 = False
                         Exit Do
                      End If
                      oRsTmp86.MoveNext
                   Loop
                End If
                oRsTmp86.Close
                '
                If lbContinua86 = True And oRsTmp1!idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoEnCitaWeb Then
                     lnDiasQfaltaPcita = DateDiff("d", CDate(Format(ldFechaHoy, sighentidades.DevuelveFechaSoloFormato_DMY) & _
                                                  " " & lcHoraHoy), _
                                                 CDate(Format(oRsTmp1!fecha, sighentidades.DevuelveFechaSoloFormato_DMY) & _
                                                 " " & oRsTmp1!HoraInicio))
                     Set oDOCitasWebCupos = oCitasWebCupos.SeleccionarPorIdCitaBloqueada(oRsTmp1!idCitaBloqueada)
                     If oDOCitasWebCupos.idCitaBloqueada > 0 Then
                            'asigna al cupo reservado la cita web, incluyendo el paciente web
                            '(si es necesario crea HISTORIA)
                            If Not IsNull(oRsTmp1!idWeb) Then
                               oDOCitasWebCupos.idWeb = oRsTmp1!idWeb
                            End If
                            oDOCitasWebCupos.idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoYconCitaEnGalenhos
                            oDOCitasWebCupos.dni = oRsTmp1!dni
                            oDOCitasWebCupos.ApellidoPaterno = oRsTmp1!ApellidoPaterno
                            oDOCitasWebCupos.ApellidoMaterno = oRsTmp1!ApellidoMaterno
                            oDOCitasWebCupos.PrimerNombre = oRsTmp1!PrimerNombre
                            oDOCitasWebCupos.SegundoNombre = IIf(IsNull(oRsTmp1!SegundoNombre), "", oRsTmp1!SegundoNombre)
                            oDOCitasWebCupos.idTipoSexo = IIf(IsNull(oRsTmp1!idTipoSexo), 0, oRsTmp1!idTipoSexo)
                            oDOCitasWebCupos.FechaNacimiento = IIf(IsNull(oRsTmp1!FechaNacimiento), 0, oRsTmp1!FechaNacimiento)
                            If Not IsNull(oRsTmp1!Ubigeo) Then
                               oDOCitasWebCupos.Ubigeo = oRsTmp1!Ubigeo
                            End If
                            If Not IsNull(oRsTmp1!FechaConfirmacion) Then
                               oDOCitasWebCupos.FechaConfirmacion = oRsTmp1!FechaConfirmacion
                            End If
                            If Not IsNull(oRsTmp1!HoraConfirmacion) Then
                               oDOCitasWebCupos.HoraConfirmacion = oRsTmp1!HoraConfirmacion
                            End If
                            oDOCitasWebCupos.idFuenteFinanciamiento = IIf(IsNull(oRsTmp1!idFuenteFinanciamiento), 0, oRsTmp1!idFuenteFinanciamiento)
                            If Not IsNull(oRsTmp1!IdTurno) Then
                               oDOCitasWebCupos.IdTurno = oRsTmp1!IdTurno
                            End If
                            If Not IsNull(oRsTmp1!email) Then
                               oDOCitasWebCupos.email = oRsTmp1!email
                            End If
                            If Not IsNull(oRsTmp1!telefono) Then
                               oDOCitasWebCupos.telefono = oRsTmp1!telefono
                            End If
                            
                            lnIdEstablecimientoOrigen = 0: lcNroReferenciaOrigen = "": lcReferenciaMedicoOColeg = ""
                            lnReferenciaOidDiagnostico = 0: lcReferenciaOservicio = ""
                            If Not IsNull(oRsTmp1!refNumero) Then
                                lnIdEstablecimientoOrigen = oRsTmp1!refIdEess
                                lcNroReferenciaOrigen = oRsTmp1!refNumero
                                If Not IsNull(oRsTmp1!refMedicoCOLEGIATURA) Then
                                   lcReferenciaMedicoOColeg = oRsTmp1!refMedicoCOLEGIATURA
                                End If
                                If Not IsNull(oRsTmp1!refidDiagnostico) Then
                                   lnReferenciaOidDiagnostico = oRsTmp1!refidDiagnostico
                                End If
                                lcReferenciaOservicio = oRsTmp1!refUPS
                            End If
                            
                            If mo_ReglasDeProgMedica.CrearCitaAutomatica(oRsTmp1!dni, oRsTmp1!fecha, oRsTmp1!HoraInicio, oRsTmp1!HoraFinal, _
                                                   oRsTmp1!idServicio, oRsTmp1!idMedico, oRsTmp1!ApellidoMaterno, _
                                                   oRsTmp1!ApellidoPaterno, oRsTmp1!FechaNacimiento, _
                                                   IIf(IsNull(oRsTmp1!Ubigeo), 0, oRsTmp1!Ubigeo), _
                                                   oRsTmp1!idTipoSexo, oRsTmp1!PrimerNombre, _
                                                   IIf(IsNull(oRsTmp1!SegundoNombre), "", oRsTmp1!SegundoNombre), _
                                                   oRsTmp1!idFuenteFinanciamiento, ldFechaHoy, lcHoraHoy, _
                                                   lnHistoriaCreada, oConexion, ml_idUsuario, mo_lnIdTablaLISTBARITEMS, _
                                                   mo_lcNombrePc, False, oDOCitasWebCupos.idPaciente, lnIdAtencion, _
                                                   lnIdEstablecimientoOrigen, lcNroReferenciaOrigen, lcReferenciaMedicoOColeg, _
                                                   lnReferenciaOidDiagnostico, lcReferenciaOservicio) = False Then
                               If mo_ReglasDeProgMedica.MensajeError <> "" Then
                                  ms_MensajeError = "No se pudo CREAR CITA AUTOMATICA"
                               End If
                               MsgBox "En Sisgalenplus: " & mo_ReglasDeProgMedica.MensajeError: GoTo ErrImpCWEB
                            End If
                            If mo_ReglasDeProgMedica.CitasBloqueadasEliminarPorId(oRsTmp1!idCitaBloqueada, oConexion) = False Then
                               If mo_ReglasDeProgMedica.MensajeError = "" Then
                                  ms_MensajeError = "No se pudo ELIMINAR CITA BLOQUEADA"
                               End If
                               MsgBox "En Sisgalenplus: " & mo_ReglasDeProgMedica.MensajeError: GoTo ErrImpCWEB
                            End If
                            '
                            mo_atenciones.idAtencion = lnIdAtencion
                            If oAtencion.SeleccionarPorId(mo_atenciones) = True Then
                               mo_Pacientes.idPaciente = mo_atenciones.idPaciente
                               If oPacientes.SeleccionarPorId(mo_Pacientes) = True Then
                                    lcMensajeCelular = "CITA WEB, para " & mo_atenciones.FechaIngreso & " " & mo_atenciones.HoraIngreso & _
                                                       " (N° Cuenta: " & mo_atenciones.idCuentaAtencion & ") " & _
                                                       "(" & Trim(oRsTmp1!ApellidoPaterno) & " " & Trim(oRsTmp1!ApellidoMaterno) & oRsTmp1!PrimerNombre & _
                                                       ") en " & wxParametro205
                                    oMensajeCelular.MensajeCelularEnviar mo_Pacientes, mo_atenciones.idCuentaAtencion, lcMensajeCelular, _
                                                    "CITAS WEB", oConexion
                                End If
                            End If
                            '
                            oDOCitasWebCupos.idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoYconCitaEnGalenhos
                            oDOCitasWebCupos.idPaciente = mo_atenciones.idPaciente
                            oDOCitasWebCupos.idWeb = mo_atenciones.idCuentaAtencion
                            If oCitasWebCupos.Modificar(oDOCitasWebCupos) = False Then
                               If oCitasWebCupos.MensajeError = "" Then
                                  ms_MensajeError = "No se pudo MODIFICAR CITASWEBCUPOS"
                               End If
                               MsgBox "En Sisgalenplus: " & oCitasWebCupos.MensajeError: GoTo ErrImpCWEB
                            End If
                            If oRsEmailMensajes.RecordCount > 0 Then
                                oRsEmailMensajes.MoveFirst
                                oRsEmailMensajes.Find "paciente='" & oDOCitasWebCupos.dni & "'"
                                If Not oRsEmailMensajes.EOF Then
                                   oRsEmailMensajes!nroCuenta = 1
                                   oRsEmailMensajes!idServicio = oDOCitasWebCupos.idWeb     'N cuenta
                                   oRsEmailMensajes.Update
                                End If
                                
                            End If
                     End If
                End If
           End If
           oRsTmp1.MoveNext
       Loop
    End If
    '
    oConexionExterna.CommitTrans
    oConexionExterna.Close
    'enviar EMAIL a CS/PS con las CITAS ya GENERADAS y sin problemas
    oRsEmailMensajes.Filter = "nroCuenta=1"
    If oRsEmailMensajes.RecordCount > 0 Then
       oRsEmailMensajes.MoveFirst
       Do While Not oRsEmailMensajes.EOF
          EnviaEmail lcParametro524, lcParametro523, oRsEmailMensajes!Estancia, _
                     "", _
                     oRsEmailMensajes!Motivo, oRsEmailMensajes!destino
          oRsEmailMensajes.MoveNext
       Loop
    End If
    '
    oConexion.Close
    '
    lbImportaCitasWeb = True
ErrImpCWEB:
    If lbImportaCitasWeb = False Then
        oConexionExterna.RollbackTrans
        oConexionExterna.Close
        oConexion.Close
    End If
    Set mo_ReglasDeProgMedica = Nothing
    Set mo_ReglasAdmision = Nothing
    Set mo_ReglasServiciosHosp = Nothing
    Set oConexionExterna = Nothing
    Set oConexion = Nothing
    Set oRsTmp1 = Nothing
    Set mo_atenciones = Nothing
    Set oAtencion = Nothing
    Set oMensajeCelular = Nothing
    Set mo_Pacientes = Nothing
    Set oPacientes = Nothing
    Set oRsEmailMensajes = Nothing
    Set oRsTmp86 = Nothing
    'Resume
End Sub







Sub MensajeCelularEnviarSegunDxGRAVE(mo_Pacientes As DOPaciente, oRsDiagnosticos As Recordset, _
                                        lcProviene As String, lnIdCuentaAtencion As Long, _
                                        Optional oConexion1 As Connection)
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim oMensajeCelular As New SIGHProxies.Procesos
Dim lcMensajeCelular As String
If oConexion1 Is Nothing Then
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
End If
If oRsDiagnosticos.RecordCount > 0 Then
oRsDiagnosticos.MoveFirst
Do While Not oRsDiagnosticos.EOF
   If InStr("/102/301/", Trim(str(oRsDiagnosticos!idTipoDiagnostico))) > 0 Then   '/ce=definitivo/emer=final/hosp=final/
         With oCommand
             .CommandType = adCmdStoredProc
             If oConexion1 Is Nothing Then
                Set .ActiveConnection = oConexion
             Else
                Set .ActiveConnection = oConexion1
             End If
             .CommandTimeout = 150
             .CommandText = "sms_diagnosticosgravesPORdx"
             Set oParameter = .CreateParameter("@dx", adVarChar, adParamInput, 7, oRsDiagnosticos!CodigoCie2004):  .Parameters.Append oParameter
             Set oRecordset = .Execute
             If oConexion1 Is Nothing Then
                Set oRecordset.ActiveConnection = Nothing
             End If
        End With
        Set oCommand = Nothing
        If oRecordset.RecordCount > 0 Then
           oRecordset.MoveFirst
           Do While Not oRecordset.EOF
            lcMensajeCelular = Trim(oRecordset!mensaje) & " (hc: " & mo_Pacientes.NroHistoriaClinica & " <>" & _
                               " Cta: " & Trim(str(lnIdCuentaAtencion)) & " <> " & _
                               mo_Pacientes.ApellidoPaterno & " " & mo_Pacientes.ApellidoMaterno & " " & _
                               mo_Pacientes.PrimerNombre & " <>cel: " & mo_Pacientes.telefono & " <>Dir: " & _
                               mo_Pacientes.DireccionDomicilio & ")"
            oMensajeCelular.MensajeCelularEnviar mo_Pacientes, lnIdCuentaAtencion, lcMensajeCelular, _
                            lcProviene, IIf(oConexion1 Is Nothing, oConexion, oConexion1), oRecordset!celularEspecialista
            oRecordset.MoveNext
           Loop
        End If
        oRecordset.Close
   End If
   oRsDiagnosticos.MoveNext
Loop
End If
If oConexion1 Is Nothing Then
    Set oParameter = Nothing
    Set oConexion = Nothing
End If
Set oRecordset = Nothing
Set oMensajeCelular = Nothing
Exit Sub
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Sub
End Sub



Function VerificaLicenciaMensajeTexto() As Boolean
    Dim lcMensajeLicencia As String, lbTieneLicencia As Boolean
    lbTieneLicencia = True 'licencia
    VerificaLicenciaMensajeTexto = lbTieneLicencia
End Function





Sub ExportaSEM2017medicos(lcFechaInicio As String, lcFechaFinal As String)
    On Error GoTo ErrorProceso
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_AdminAdmision As New ReglasAdmision
    Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_ReglasServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
    Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
    Dim mo_ReglasHoteleria As New SIGHNegocios.ReglasHoteleria
    Dim rsReporte As New Recordset
    Dim oRsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim rsOcupacion As New Recordset
    Dim oRsFox As New Recordset
    Dim rsServicio As New Recordset
    Dim oRsErrores As New Recordset
    Dim lbPrimerError As Boolean, lbNoExisteH As Boolean
    Dim oConexionFox As New Connection
    Dim oConexion As New Connection
    Dim lcDx1 As String, lcDx2 As String
    Dim lnNumMov As Long, lcSql As String, lnEstancia As Integer
    Dim lcFec1 As String, lcFec2 As String
    Dim lcMotivo As String, lcDestinoAtencion As String, lcReferido As String, lcUsuario As String
    Dim lcDireccionDomicilio As String, lcNombreAcompaniante As String
    Dim lnEsObservacionEmergencia As Integer, lcServicio As String, lcMedico As String, lcServicioObs As String
    Dim lcMedicoObs As String, ldFechaIngObs As Date, lcHoraIngObs As String, lcCamaObs As String, ldFechaSalObs As Date
    Dim lcHoraSalObs As String, lcDxObs As String, lcHoraEstanciaMax As String, lnTotal As Long
    Dim lcRenaes As String, lcUbigeoEESS As String, lcDisa As String, lcRed As String, lcMred As String
    Dim lcCodSal As String, lnRNmuertos As Long, lnRNvivos As Long, ldFechaNacimiento As Date, lcProfParto As String
    Dim lcCpt1 As String, lcCpt2 As String, lcCpt3 As String, lcCpt4 As String, lcDxTipo1 As String, lcDxTipo2 As String
    Dim lcDNIacompainante As String, lcNroEmergencia As String, lcUsuario9 As String, ldUsuarioFecha9 As Date
    Dim lcDx3 As String, lcDx4 As String, lcDxTipo3 As String, lcDxTipo4 As String, lcDiagn As String, lcProf As String
    Dim lcPrioridad As String, lcForm_ing As String, lcMensajes99 As String, lcUPS11 As String, ldFecha_ing As Date
    Dim lcCond As String, lcCod_2000 As String, lcNombre As String
    Const lcEtniaDefault As String = "80"

    FileCopy App.Path + "\archivos\sem_me1.dbf", App.Path + "\archivos\sem_med.dbf"
    
    lcRenaes = Right("00000000" & lcBuscaParametro.SeleccionaFilaParametro(208), 10)
    lcUbigeoEESS = lcBuscaParametro.SeleccionaFilaParametro(242)
    lcHoraEstanciaMax = lcBuscaParametro.SeleccionaFilaParametro(201)
    lcDisa = lcBuscaParametro.SeleccionaFilaParametro(239)
    lcRed = lcBuscaParametro.SeleccionaFilaParametro(240)
    lcMred = lcBuscaParametro.SeleccionaFilaParametro(241)
        '
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
        '
    oConexionFox.CommandTimeout = 300
    oConexionFox.Open "DSN=SEM"
    lcParaMsgbox = " (Eliminar Datos)"
    
    lcMensajes99 = ""
    
    
    '
    lcUsuario = " "
    Set oRsTmp1 = mo_ReglasComunes.EmpleadosSeleccionarPorIdEmpleado(ml_idUsuario)
    If oRsTmp1.RecordCount > 0 Then
       If Not IsNull(oRsTmp1.Fields!Usuario) Then
          lcUsuario = oRsTmp1.Fields!Usuario
       End If
    End If
    oRsTmp1.Close
    '
    Set rsReporte = mo_ReglasAdmision.AtencionesSeleccionarHospEmergPorFechasDeEgresoMedico(CDate(lcFechaInicio), CDate(lcFechaFinal), "FECHAEGRESO")
    rsReporte.Filter = "codigoCondicionAltaSEM<>null"
    lnTotal = rsReporte.RecordCount
    
    If lnTotal = 0 Then
      MsgBox "No existe Información  con esos Datos", vbInformation, "Mensaje"
    Else
      mo_ReglasComunes.AbreTablaFox oRsFox, oConexionFox, "sem_med", "", ""
      lcCod_2000 = Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9)
      '
      lnNumMov = 1
      rsReporte.MoveFirst
      Do While Not rsReporte.EOF

        '
        lcCodSal = "."
        If Not (IsNull(rsReporte!idTipoDocumento) Or IsNull(rsReporte!dni) Or IsNull(rsReporte!idColegioHIS)) Then
           lcCodSal = Trim(str(rsReporte!idTipoDocumento)) & Trim(rsReporte!dni) & rsReporte!idColegioHIS
        End If
        lcNombre = "": ldFecha_ing = CDate("01/01/2000")
        lcProf = Space(2): lcCond = Space(2)
        Set oRsTmp2 = mo_ReglasComunes.EmpleadosSeleccionarPorDNI(rsReporte!dni)
        If oRsTmp2.RecordCount > 0 Then
           lcNombre = oRsTmp2!ApellidoPaterno & " " & oRsTmp2!ApellidoMaterno & " " & oRsTmp2!Nombres
           ldFecha_ing = oRsTmp2!FechaIngreso
           lcProf = oRsTmp2!tipoEmpleadoSEM
           lcCond = oRsTmp2!condicionTrabajoSEM
        End If
        oRsTmp2.Close
        lbNoExisteH = True
        If oRsFox.RecordCount > 0 Then
           oRsFox.MoveFirst
           oRsFox.Find "codpsal='" & lcCodSal & "'"
           If Not oRsFox.EOF Then
              lbNoExisteH = False
           End If
        End If
        If lbNoExisteH = True Then
            oRsFox.AddNew
            oRsFox.Fields!codpsal = lcCodSal
            oRsFox.Fields!nombre = Left(lcNombre, 40)
            oRsFox.Fields!plaza = 0
            oRsFox.Fields!cod_2000 = lcCod_2000
            oRsFox.Fields!cod_prof = lcProf
            oRsFox.Fields!cod_cond = lcCond
            oRsFox.Fields!fecha_ing = ldFecha_ing
            oRsFox.Fields!fecha_baja = 0
            oRsFox.Update
        End If
        lnNumMov = lnNumMov + 1
        rsReporte.MoveNext
      Loop
      Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "A", 0, "Atenciones", oConexion, 500 + 140, mo_lcNombrePc, "Exportar al SEM_2017_medicos: " & lcFechaInicio & " al " & lcFechaFinal)       '500+ ListBarReporte.idReporte
      MsgBox "Se exportó la lista de MEDICOS en el archivo: c:\arc..\dig..\galenhos\archivos\sem_med.dbf", vbInformation, ""
    End If
    
    oRsFox.Close
    oConexionFox.Close
    oConexion.Close
    Set oConexion = Nothing
    Set oConexionFox = Nothing
    Set oRsFox = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_AdminAdmision = Nothing
    Set mo_ReglasSeguridad = Nothing
    Set mo_ReglasAdmision = Nothing
    Set mo_ReglasServiciosHosp = Nothing
    Set mo_ReglasComunes = Nothing
    Set mo_ReglasHoteleria = Nothing
    Set rsReporte = Nothing
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set rsServicio = Nothing
    Set oRsErrores = Nothing
    Set rsOcupacion = Nothing
    Exit Sub
ErrorProceso:
    MsgBox Err.Description & Chr(13) & lcParaMsgbox    '
    Exit Sub
    Resume
End Sub


Sub MensajeCelularEnviarAconvenios(mo_Pacientes As DOPaciente, lcTipoFinanciamiento As String, _
                                   lnIdCuentaAtencion As Long, oConexion1 As Connection, _
                                   lcServicio As String)
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oMensajeCelular As New SIGHProxies.Procesos
Dim lcMensajeCelular As String
         With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion1
             .CommandTimeout = 150
             .CommandText = "sms_diagnosticosgravesPORdx"
             Set oParameter = .CreateParameter("@dx", adVarChar, adParamInput, 7, lcTipoFinanciamiento):  .Parameters.Append oParameter
             Set oRecordset = .Execute
             If oConexion1 Is Nothing Then
                Set oRecordset.ActiveConnection = Nothing
             End If
        End With
        Set oCommand = Nothing
        If oRecordset.RecordCount > 0 Then
           oRecordset.MoveFirst
           Do While Not oRecordset.EOF
            lcMensajeCelular = Trim(oRecordset!mensaje) & _
                               " <>Cta: " & Trim(str(lnIdCuentaAtencion)) & " <> " & _
                               mo_Pacientes.ApellidoPaterno & " " & mo_Pacientes.ApellidoMaterno & " " & _
                               mo_Pacientes.PrimerNombre & "  <>Serv: " & lcServicio & _
                               " (hc: " & mo_Pacientes.NroHistoriaClinica & " <>" & " <>cel: " & _
                               mo_Pacientes.telefono & " <>Dir: " & mo_Pacientes.DireccionDomicilio & ")"
            oMensajeCelular.MensajeCelularEnviar mo_Pacientes, lnIdCuentaAtencion, lcMensajeCelular, _
                            "CONVENIOS", oConexion1, oRecordset!celularEspecialista
            oRecordset.MoveNext
           Loop
        End If
        oRecordset.Close
Set oRecordset = Nothing
Set oMensajeCelular = Nothing
Exit Sub
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Sub
End Sub


Sub EnviaMensajeCelularPorCuenta(lnIdCuenta As Long, lcMensajeCelular As String, lcOpcion As String)
    On Error GoTo ErrEnvMenCel
    If lnIdCuenta > 0 Then
        Dim oMensajeCelular As New SIGHProxies.Procesos
        Dim mo_atenciones As New DOAtencion
        Dim oAtenciones As New Atenciones
        Dim oConexion As New Connection
        Dim mo_Pacientes As New DOPaciente
        Dim oPacientes As New Pacientes
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        Set oAtenciones.Conexion = oConexion
        mo_atenciones.idCuentaAtencion = lnIdCuenta
        If oAtenciones.SeleccionarPorIdCuentaDeAtencion(mo_atenciones) Then
            Set oPacientes.Conexion = oConexion
            mo_Pacientes.idPaciente = mo_atenciones.idPaciente
            If oPacientes.SeleccionarPorId(mo_Pacientes) Then
               lcMensajeCelular = lcMensajeCelular & " en " & lcBuscaParametro.SeleccionaFilaParametro(205)
               oMensajeCelular.MensajeCelularEnviar mo_Pacientes, mo_atenciones.idCuentaAtencion, lcMensajeCelular, _
                                                    lcOpcion, oConexion
            End If
        End If
        oConexion.Close
        Set oMensajeCelular = Nothing
        Set mo_atenciones = Nothing
        Set oAtenciones = Nothing
        Set oConexion = Nothing
        Set mo_Pacientes = Nothing
        Set oPacientes = Nothing
    End If
ErrEnvMenCel:
End Sub


Sub LimpiaHistorico()
    On Error GoTo ErrHistorico
    Dim lcSql As String
    Dim oRsParametros As New Recordset
    Dim oRsMDB As New Recordset
    Dim oConexionMDB As New Connection
    lcSql = "select ValorTexto from Parametros where idparametro=581"
    oRsParametros.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
    oConexionMDB.CommandTimeout = 900
    oConexionMDB.CursorLocation = adUseClient
    oConexionMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =" & Trim(oRsParametros!valorTexto) & "\parametros.mdb"
    lcSql = "delete from   where month(fechaIngreso)<>" & Month(Date)
    If oRsMDB.State = 1 Then oRsMDB.Close
    oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    oConexionMDB.Close
ErrHistorico:
    Set oRsParametros = Nothing
    Set oRsMDB = Nothing
    Set oConexionMDB = Nothing
End Sub

'debb-12/07/2019
Sub ReportePartidaREsumenConNotaCredito(ByRef mrs_Tmp As Recordset, mda_FechaInicio As Date, mda_FechaFin As Date, _
                          ml_IdPartidaFiltro As Long, lnTipoProceso As sghTipoProceso, _
                          moConexion As Connection, lbSeCreaTemporal As Boolean, lnIdCajero As Long, _
                          ByRef lnFacturasTotal As Long, ByRef lnFacturasAnuladas As Long, _
                          ByRef lnBoletasTotal As Long, ByRef lnBoletasAnuladas As Long, _
                          lbSoloProcesaNotaCredito As Boolean)
        Dim lnExoneracionCamaSP As Double, lnExoneracionSOPSP As Double, lnExoneracionRestoSP As Double, lnExoneracionItemSP As Double, lnExoneracionConsultaSP As Double
        Dim lcTexto3 As String, lcTexto2 As String, lcTexto1 As String
        Dim lRecordCount As Long, lnLineas As Long, lnIdTipoServicio As Long, lnIdCentroCosto As Long
        Dim lnPagoCta As Double, lnExoneracion As Double, lnAnulado As Double, lnImpTotal As Double
        Dim lnCama As Double, lnSoloSOP As Double, lnResto As Double, lnQueda As Double
        Dim lnAnuladoCama As Double, lnImptotalCama As Double, lnExoneracionConsulta As Double
        Dim lnAnuladoSOP As Double, lnImptotalSOP As Double
        Dim lnAnuladoResto As Double, lnImptotalResto As Double
        Dim lnConsulta As Double, lnAnuladoConsulta As Double
        Dim lnImptotalConsulta As Double
        Dim lnImpLaboratorio As Double, lnExoneracionLaboratorioSP As Double
        Dim lnImpImagenes As Double, lnExoneracionImagenesSP As Double
        Dim lnAnuladoLaboratorio As Double, lnImpTotalLaboratorio As Double
        Dim lnAnuladoImagenes As Double, lnImpTotalImagenes As Double
        Dim lbPrimeraVez As Boolean, lbEncontroDato As Boolean
        Dim rsTmp As New Recordset
        Dim rsReporte As New Recordset
        Dim oRsExoneradoBoleta As New Recordset
        Dim oRsItemsXCuenta As New Recordset
        Dim rsParamet As New Recordset
        Dim rsTmp1 As New Recordset
        Dim rsTmp2 As New Recordset
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
        Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
        Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
        Dim rsTmp999 As New Recordset
        
        Dim lnExoneracionImpDetall As Double, lnExoneracionCanDetall As Long
        Dim lnAnuladoImpDetall As Double, lnAnuladoCanDetall As Long
        Dim lnImptotalDetall As Double, lnCanTotalDetall As Long
        Dim lnIdProductoDetall As Long, lnCodigoDetall As String, lnDescripcionDetall As String, lnPrecioDetall As Double
        Dim lbHallo As Boolean, lcPie As String
        Dim lnBoletaExo As Double, lnBoletaTot As Double, lnBoletaAnu As Double
        Dim lnBoletaExo1 As Double, lnBoletaTot1 As Double, lnBoletaAnu1 As Double
        Dim lnIdPartida100 As Long, lnIdProducto100 As Long, lcCodigo100 As String, lcProducto100 As String
        Dim lnPrecioVenta100 As Double, lnCantidadPagar100 As Long, lnTotalPagar100 As Double
        Dim lnIdServiciosHospitalarios As Long, lnIdReembolsosServicios As Long
        Dim lnIdEstadoComprobante As Long
        '
        On Error GoTo ErrRXCC
        '
        lnFacturasTotal = 0: lnFacturasAnuladas = 0: lnBoletasTotal = 0: lnBoletasAnuladas = 0

        '
        lnIdServiciosHospitalarios = Val(lcBuscaParametro.SeleccionaFilaParametro(245))
        lnIdReembolsosServicios = Val(lcBuscaParametro.SeleccionaFilaParametro(251))
        '
        lcTexto3 = "..comienza a generar Temporal..."
        If lbSeCreaTemporal = True Then
            If ml_IdPartidaFiltro = 0 Then
                With mrs_Tmp
                   .Fields.Append "idPartida", adInteger
                   .Fields.Append "Codigo", adVarChar, 10, adFldIsNullable
                   .Fields.Append "Descripcion", adVarChar, 150, adFldIsNullable
                   .Fields.Append "ImpAnulado", adDouble
                   .Fields.Append "ImpExonerado", adDouble
                   .Fields.Append "ImpNormal", adDouble
                   .Fields.Append "ImpCancelado", adDouble
                   .Fields.Append "ImpSubTotal", adDouble
                   .Fields.Append "CtaCodigo", adVarChar, 20, adFldIsNullable
                   .Fields.Append "CtaDescripcion", adVarChar, 150, adFldIsNullable
                   .Fields.Append "ImpNotaCredito", adDouble
                   .LockType = adLockOptimistic
                   .Open
                End With
            Else
                With mrs_Tmp
                   .Fields.Append "Identificador", adInteger
                   .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
                   .Fields.Append "Descripcion", adVarChar, 300, adFldIsNullable
                   .Fields.Append "ImpAnulado", adDouble
                   .Fields.Append "ImpExonerado", adDouble
                   .Fields.Append "ImpNormal", adDouble
                   .Fields.Append "ImpCancelado", adDouble
                   .Fields.Append "idPartida", adInteger
                   .Fields.Append "ImpSubTotal", adDouble
                   .LockType = adLockOptimistic
                   .Open
                End With
            End If
        End If
        '
        lcTexto3 = "..llena Temporal..."
        If ml_IdPartidaFiltro = 0 And mrs_Tmp.RecordCount = 0 Then
            'Llena Temporal con Partidas
            Set rsTmp = FactPartidasPresupuestalesSeleccionarTodas(moConexion)
            If rsTmp.RecordCount = 0 Then
               MsgBox "Llene tabla PARTIDAS"
               Exit Sub
            End If
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
               mrs_Tmp.AddNew
               mrs_Tmp.Fields!IdPartida = rsTmp.Fields!IdPartidaPresupuestal
               mrs_Tmp.Fields!codigo = rsTmp.Fields!codigo
               mrs_Tmp.Fields!Descripcion = rsTmp.Fields!Descripcion
               mrs_Tmp.Update
               rsTmp.MoveNext
            Loop
            rsTmp.Close
        End If
        lcTexto3 = "..comienza a jalar boletas servicios..."
        'Boletas Servicios
        If lbSoloProcesaNotaCredito = True Then
           Set rsReporte = NotaDeCreditoDebitoServicios(0, mda_FechaInicio, mda_FechaFin, 0, 0)
        Else
           Set rsReporte = mo_ReglasFacturacion.ServicioConsolidado(0, mda_FechaInicio, mda_FechaFin, 0, 0)
        End If
        '
        If lnIdCajero > 0 Then
           rsReporte.Filter = "idCajero=" & lnIdCajero
        End If
        '
        lRecordCount = rsReporte.RecordCount
        If lRecordCount > 0 Then
           If lbSoloProcesaNotaCredito = True Then
              Set rsTmp999 = NotaDeCreditoServiciosXfechas(mda_FechaInicio, mda_FechaFin)
           Else
              Set rsTmp999 = mo_ReglasCaja.CajaComprobantesPagoXnroSerieYDocumentoFechas(mda_FechaInicio, mda_FechaFin, moConexion)
           End If
           lnLineas = 0

           rsReporte.MoveFirst
           Do While Not rsReporte.EOF
              lcTexto2 = rsReporte.Fields!NroSerie + rsReporte.Fields!nroDocumento
              lnPagoCta = rsReporte.Fields!Adelantos
              lnExoneracion = rsReporte.Fields!exoneraciones
              lnIdEstadoComprobante = rsReporte.Fields!IdEstadoComprobante
              If rsReporte.Fields!IdEstadoComprobante = 9 Then
                 lnAnulado = rsReporte.Fields!totalPagado
                 lnImpTotal = 0
                 If rsReporte!idTipoComprobante = 2 Then
                    lnFacturasAnuladas = lnFacturasAnuladas + 1
                 ElseIf rsReporte!idTipoComprobante = 3 Then
                    lnBoletasAnuladas = lnBoletasAnuladas + 1
                 End If
                 
              Else
                 lnAnulado = 0
                 lnImpTotal = rsReporte.Fields!totalPagado
                 
              End If
              '
              If rsReporte!idTipoComprobante = 2 Then
                 lnFacturasTotal = lnFacturasTotal + 1
              ElseIf rsReporte!idTipoComprobante = 3 Then
                 lnBoletasTotal = lnBoletasTotal + 1
              End If
              '
              lnIdTipoServicio = 0
              lbPrimeraVez = True
              '
              Set oRsExoneradoBoleta = mo_ReglasCaja.FacturacionServicioFinanciamientosExoneracionesEnBoleta(rsReporte.Fields!IdComprobantePago, moConexion)
              '
              lnBoletaExo = rsReporte.Fields!exoneraciones
              lnBoletaTot = rsReporte.Fields!totalPagado: lnBoletaAnu = 0
              Select Case rsReporte.Fields!IdEstadoComprobante
              Case 6   '***Devolucion
                   lnBoletaTot = -rsReporte.Fields!totalPagado
              Case 9   '***anulado
                   lnBoletaTot = 0: lnBoletaAnu = rsReporte.Fields!totalPagado
              End Select
              lnBoletaExo1 = 0: lnBoletaTot1 = 0: lnBoletaAnu1 = 0
              'solo para notas de credito
              If lbSoloProcesaNotaCredito = True Then
                If rsReporte!totalNC <> rsReporte!totalPagado Then
                   lnBoletaTot = rsReporte!totalNC
                   lnBoletaExo = rsReporte!exoneraciones
                   If lnBoletaExo > 0 Then
                      lnBoletaExo = Round((rsReporte!totalNC * rsReporte!exoneraciones) / rsReporte!totalPagado, 2)
                   End If
                   lnBoletaAnu = 0
                End If
              End If
              
              '
              Do While Not rsReporte.EOF And lcTexto2 = (rsReporte.Fields!NroSerie + rsReporte.Fields!nroDocumento)
If Val(rsReporte.Fields!NroSerie) = 22 And Val(rsReporte.Fields!nroDocumento) = 1372884 Then
lcTexto1 = ""
End If
                    lnLineas = lnLineas + 1
                    '
                    If lbPrimeraVez = True Then
                       lbPrimeraVez = False
                       
                       '*****Emergencia
                       lcTexto3 = "..comienza: Emergencia..."
                       lnConsulta = 0: lnResto = 0: lnImpLaboratorio = 0: lnImpImagenes = 0
                       'Set rsTmp = CajaComprobantesPagoXnroSerieYDocumento(rsReporte.Fields!NroSerie, rsReporte.Fields!NroDocumento, moConexion)
                       rsTmp999.Filter = "idComprobantePago=" & rsReporte.Fields!IdComprobantePago
                       
                       
                          If rsTmp999.RecordCount > 0 Then
                             '
                             lnExoneracionConsultaSP = 0: lnExoneracionRestoSP = 0
                             lnExoneracionLaboratorioSP = 0: lnExoneracionImagenesSP = 0
                             '
                             lbEncontroDato = True
                             rsTmp999.MoveFirst
                             Do While Not rsTmp999.EOF
                             
                                '
                                lnExoneracionItemSP = 0
                                If oRsExoneradoBoleta.RecordCount > 0 Then
                                   oRsExoneradoBoleta.MoveFirst
                                   Do While Not oRsExoneradoBoleta.EOF
                                      If oRsExoneradoBoleta.Fields!idOrden = rsTmp999.Fields!idOrden And oRsExoneradoBoleta.Fields!idProducto = rsTmp999.Fields!idProducto Then
                                         lnExoneracionItemSP = lnExoneracionItemSP + oRsExoneradoBoleta.Fields!totalFinanciado
                                      End If
                                      oRsExoneradoBoleta.MoveNext
                                   Loop
                                End If
                                '
                                lnConsulta = lnConsulta + rsTmp999.Fields!Total - lnExoneracionItemSP
                                lnExoneracionConsultaSP = lnExoneracionConsultaSP + lnExoneracionItemSP
                                '
                                If lnExoneracionItemSP > 0 Then
                                   lnExoneracionCanDetall = 1: lnExoneracionImpDetall = lnExoneracionItemSP
                                Else
                                   lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
                                End If
                                Select Case rsReporte.Fields!IdEstadoComprobante
                                Case 6        '***Devolucion
                                   lnImptotalDetall = -(rsTmp999.Fields!Importe - lnExoneracionItemSP): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                   lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                Case 9        '***Anulado
                                   lnImptotalDetall = 0: lnCanTotalDetall = 0
                                   lnAnuladoImpDetall = (rsTmp999.Fields!Importe - lnExoneracionItemSP): lnAnuladoCanDetall = rsTmp999.Fields!Cantidad
                                   lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
                                Case Else     '***Pagado
                                   lnImptotalDetall = (rsTmp999.Fields!Importe - lnExoneracionItemSP): lnCanTotalDetall = rsTmp999.Fields!Cantidad
                                   lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                                   If lnPagoCta > 0 Then
                                      lnImptotalDetall = lnImptotalDetall - lnPagoCta
                                      lnPagoCta = 0
                                   End If
                                End Select
                                lnIdProductoDetall = rsTmp999.Fields!idProducto: lnCodigoDetall = rsTmp999.Fields!codigo
                                lnDescripcionDetall = rsTmp999.Fields!nombreProducto: lnPrecioDetall = rsTmp999.Fields!precio
                                '
                                lnBoletaExo1 = lnBoletaExo1 + lnExoneracionImpDetall
                                lnBoletaTot1 = lnBoletaTot1 + lnImptotalDetall
                                lnBoletaAnu1 = lnBoletaAnu1 + lnAnuladoImpDetall
                                '
                                lnIdPartida100 = rsTmp999.Fields!IdPartida
                                lnIdProducto100 = lnIdProductoDetall
                                lcCodigo100 = lnCodigoDetall
                                lcProducto100 = lnDescripcionDetall
                                lnPrecioVenta100 = lnPrecioDetall
                                lnCantidadPagar100 = rsTmp999.Fields!Cantidad
                                lnTotalPagar100 = rsTmp999.Fields!Total
                                '
                                Select Case rsTmp999.Fields!idProducto
                                Case lnIdServiciosHospitalarios And Not IsNull(rsReporte.Fields!idCuentaAtencion)
                                     '******Servicios Hospitalarios, habrá que desagregarlos en 2 Partes: Farmacia y el resto en ServiciosHospitalarios
                                     mo_ReglasFacturacion.DesagregaCuentasItemsXServHosp oRsItemsXCuenta, _
                                                          rsReporte.Fields!idCuentaAtencion, rsTmp999.Fields!Total, moConexion, _
                                                          lnIdServiciosHospitalarios
                                     If oRsItemsXCuenta.RecordCount > 0 Then
                                        oRsItemsXCuenta.MoveFirst
                                        Do While Not oRsItemsXCuenta.EOF
                                            lnExoneracionImpDetall = 0
                                            lnExoneracionCanDetall = 0
                                            Select Case rsReporte.Fields!IdEstadoComprobante
                                            Case 6        '***Devolucion
                                            Case 9        '***Anulado
                                               lnImptotalDetall = 0
                                               lnCanTotalDetall = 0
                                               lnAnuladoImpDetall = oRsItemsXCuenta.Fields!ImporteProrrateado
                                               lnAnuladoCanDetall = oRsItemsXCuenta.Fields!Cantidad
                                               lnExoneracionImpDetall = 0: lnExoneracionCanDetall = 0
                                            Case Else     '***Pagado
                                               lnImptotalDetall = oRsItemsXCuenta.Fields!ImporteProrrateado
                                               lnCanTotalDetall = oRsItemsXCuenta.Fields!Cantidad
                                               lnAnuladoImpDetall = 0
                                               lnAnuladoCanDetall = 0
                                            End Select
                                            If ml_IdPartidaFiltro > 0 Then
                                                lnIdProductoDetall = oRsItemsXCuenta.Fields!idProducto
                                                lnCodigoDetall = oRsItemsXCuenta.Fields!codigo
                                                lnDescripcionDetall = oRsItemsXCuenta.Fields!nombreProducto
                                                lnPrecioDetall = oRsItemsXCuenta.Fields!precio
                                                GrabaDetalleEnTmpPartida ml_IdPartidaFiltro, oRsItemsXCuenta.Fields!IdPartida, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, oRsItemsXCuenta.Fields!Cantidad, oRsItemsXCuenta.Fields!ImporteProrrateado, lnTipoProceso
                                            Else
                                                mrs_Tmp.MoveFirst
                                                mrs_Tmp.Find "idPartida=" & oRsItemsXCuenta.Fields!IdPartida
                                                If Not mrs_Tmp.EOF Then
                                                   mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
                                                   mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
                                                   mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
                                                   mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
                                                   mrs_Tmp.Update
                                                End If
                                            End If
                                            oRsItemsXCuenta.MoveNext
                                        Loop
                                     End If
                                Case lnIdReembolsosServicios
                                     '******Reembolsos de Servicios, habrá que desagregarlos por Cuenta->Procedimientos
                                     mo_ReglasFacturacion.DesagregaCuentasItemsXreembolso oRsItemsXCuenta, rsReporte.Fields!IdComprobantePago, _
                                                          rsTmp999.Fields!Total, moConexion
                                     If oRsItemsXCuenta.RecordCount > 0 Then
                                        oRsItemsXCuenta.MoveFirst
                                        Do While Not oRsItemsXCuenta.EOF
                                            lnExoneracionImpDetall = 0
                                            lnExoneracionCanDetall = 0
                                            Select Case rsReporte.Fields!IdEstadoComprobante
                                            Case 6        '***Devolucion
                                            Case 9        '***Anulado
                                               lnImptotalDetall = 0
                                               lnCanTotalDetall = 0
                                               lnAnuladoImpDetall = oRsItemsXCuenta.Fields!ImporteProrrateado
                                               lnAnuladoCanDetall = oRsItemsXCuenta.Fields!Cantidad
                                            Case Else     '***Pagado
                                               lnImptotalDetall = oRsItemsXCuenta.Fields!ImporteProrrateado
                                               lnCanTotalDetall = oRsItemsXCuenta.Fields!Cantidad
                                               lnAnuladoImpDetall = 0
                                               lnAnuladoCanDetall = 0
                                            End Select
                                            If ml_IdPartidaFiltro > 0 Then
                                                lnIdProductoDetall = oRsItemsXCuenta.Fields!idProducto
                                                lnCodigoDetall = oRsItemsXCuenta.Fields!codigo
                                                lnDescripcionDetall = oRsItemsXCuenta.Fields!nombreProducto
                                                lnPrecioDetall = oRsItemsXCuenta.Fields!precio
                                                GrabaDetalleEnTmpPartida ml_IdPartidaFiltro, oRsItemsXCuenta.Fields!IdPartida, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, oRsItemsXCuenta.Fields!Cantidad, oRsItemsXCuenta.Fields!ImporteProrrateado, lnTipoProceso
                                            Else
                                                mrs_Tmp.MoveFirst
                                                mrs_Tmp.Find "idPartida=" & oRsItemsXCuenta.Fields!IdPartida
                                                If Not mrs_Tmp.EOF Then
                                                   mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
                                                   mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
                                                   mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
                                                   mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
                                                   mrs_Tmp.Update
                                                End If
                                            End If
                                            oRsItemsXCuenta.MoveNext
                                        Loop
                                     End If
                                Case Else
                                     If ml_IdPartidaFiltro > 0 Then
                                        GrabaDetalleEnTmpPartida ml_IdPartidaFiltro, rsTmp999.Fields!IdPartida, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProductoDetall, lnCodigoDetall, lnDescripcionDetall, lnPrecioDetall, rsTmp999.Fields!Cantidad, rsTmp999.Fields!Total, lnTipoProceso
                                     Else
                                        mrs_Tmp.MoveFirst
                                        mrs_Tmp.Find "idPartida=" & rsTmp999.Fields!IdPartida
                                        If Not mrs_Tmp.EOF Then
                                          mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
                                          mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
                                          mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
                                          mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
                                          mrs_Tmp.Update
                                        End If
                                     End If
                                End Select
                                rsTmp999.MoveNext
                             Loop
                          End If
                         ' rsTmp.Close
                    End If
                    rsReporte.MoveNext
                    If rsReporte.EOF Then
                       Exit Do
                    End If
              Loop
              '
              oRsExoneradoBoleta.Close
              If lnIdEstadoComprobante = 9 Then
                 lnBoletaExo1 = lnBoletaExo
              End If
              'Chequea Redondeo, para que cuadre con Totales de la  Boleta
              If lnBoletaExo <> lnBoletaExo1 Or lnBoletaTot <> lnBoletaTot1 Or lnBoletaAnu <> lnBoletaAnu1 Then
                    lnAnuladoImpDetall = 0
                    lnExoneracionImpDetall = 0
                    lnImptotalDetall = 0
                    If lnBoletaExo <> lnBoletaExo1 Then
                          lnExoneracionImpDetall = lnBoletaExo - lnBoletaExo1
                    End If
                    If lnBoletaTot <> lnBoletaTot1 Then
                          lnImptotalDetall = lnBoletaTot - lnBoletaTot1
                    End If
                    If lnBoletaAnu <> lnBoletaAnu1 Then
                          lnAnuladoImpDetall = lnBoletaAnu - lnBoletaAnu1
                    End If
                    If ml_IdPartidaFiltro > 0 Then
                       GrabaDetalleEnTmpPartida ml_IdPartidaFiltro, lnIdPartida100, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProducto100, lcCodigo100, lcProducto100, lnPrecioVenta100, lnCantidadPagar100, lnTotalPagar100, lnTipoProceso
                    Else
                       mrs_Tmp.MoveFirst
                       mrs_Tmp.Find "idPartida=" & lnIdPartida100
                       If Not mrs_Tmp.EOF Then
                          mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
                          mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
                          mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
                          mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
                          mrs_Tmp.Update
                       End If
                    End If
              End If
              '
           Loop
        End If
        'Farmacia
        lcTexto3 = "..comienza: MEDICAMENTOS emitidos en CAJA SERVICIO..."
        Set rsReporte = Nothing
        
        If lbSoloProcesaNotaCredito = True Then
           Set rsReporte = NotaDeCreditoDebitoFarmacia(0, mda_FechaInicio, mda_FechaFin, 0, 0)
        Else
           Set rsReporte = mo_ReglasFacturacion.FarmaciaConsolidado(0, mda_FechaInicio, mda_FechaFin, 0, 0)
        End If
        '
        If lnIdCajero > 0 Then
           rsReporte.Filter = "idCajero=" & lnIdCajero
        End If
        '
        lRecordCount = rsReporte.RecordCount
        If lRecordCount > 0 Then
           Set oRsExoneradoBoleta = mo_ReglasCaja.FacturacionBienesFinanciamientosExoneracionesEnBoletaTodos(moConexion)
           lnLineas = 0
           '
           rsReporte.MoveFirst
           lnExoneracion = 0: lnAnulado = 0: lnImpTotal = 0
           Do While Not rsReporte.EOF
              lcTexto2 = rsReporte.Fields!NroSerie + rsReporte.Fields!nroDocumento
              lnPagoCta = rsReporte.Fields!Adelantos
              lnExoneracion = lnExoneracion + rsReporte.Fields!exoneraciones
              If rsReporte.Fields!IdEstadoComprobante = 9 Then
                 lnAnulado = lnAnulado + rsReporte.Fields!totalPagado
                 If rsReporte!idTipoComprobante = 2 Then
                    lnFacturasAnuladas = lnFacturasAnuladas + 1
                 ElseIf rsReporte!idTipoComprobante = 3 Then
                    lnBoletasAnuladas = lnBoletasAnuladas + 1
                 End If
                 
              Else
                 lnImpTotal = lnImpTotal + rsReporte.Fields!totalPagado
              End If
              '
              If rsReporte!idTipoComprobante = 2 Then
                 lnFacturasTotal = lnFacturasTotal + 1
              ElseIf rsReporte!idTipoComprobante = 3 Then
                 lnBoletasTotal = lnBoletasTotal + 1
              End If
              

              'Set oRsExoneradoBoleta = FacturacionBienesFinanciamientosExoneracionesEnBoleta(rsReporte.Fields!IdComprobantePago, moConexion)
              oRsExoneradoBoleta.Filter = "idComprobantePago=" & rsReporte.Fields!IdComprobantePago
             '
              lnBoletaExo = rsReporte.Fields!exoneraciones
              lnBoletaTot = rsReporte.Fields!totalPagado: lnBoletaAnu = 0
              lnIdEstadoComprobante = rsReporte.Fields!IdEstadoComprobante
              If rsReporte.Fields!IdEstadoComprobante = 9 Then
                 lnBoletaTot = 0: lnBoletaAnu = rsReporte.Fields!totalPagado
              End If
              'solo para notas de credito
              If lbSoloProcesaNotaCredito = True Then
                If rsReporte!totalNC <> rsReporte!totalPagado Then
                   lnBoletaTot = rsReporte!totalNC
                   lnBoletaExo = rsReporte!exoneraciones
                   If lnBoletaExo > 0 Then
                      lnBoletaExo = Round((rsReporte!totalNC * rsReporte!exoneraciones) / rsReporte!totalPagado, 2)
                   End If
                   lnBoletaAnu = 0
                End If
              End If
              '
              lnBoletaExo1 = 0: lnBoletaTot1 = 0: lnBoletaAnu1 = 0
              '
              Do While Not rsReporte.EOF And lcTexto2 = (rsReporte.Fields!NroSerie + rsReporte.Fields!nroDocumento)
                    lnIdPartida100 = rsReporte.Fields!IdPartida
                    lnIdProducto100 = rsReporte.Fields!idProducto
                    lcCodigo100 = rsReporte.Fields!codigo
                    lcProducto100 = rsReporte.Fields!Producto
                    lnPrecioVenta100 = rsReporte.Fields!PrecioVenta
                    lnCantidadPagar100 = rsReporte.Fields!CantidadPagar
                    lnTotalPagar100 = rsReporte.Fields!TotalPagar
                    lnLineas = lnLineas + 1
                    lnExoneracionImpDetall = 0
                    lnExoneracionCanDetall = 0
                    If oRsExoneradoBoleta.RecordCount > 0 Then
                       oRsExoneradoBoleta.MoveFirst
                       Do While Not oRsExoneradoBoleta.EOF
                          If oRsExoneradoBoleta.Fields!MovTipo = rsReporte.Fields!MovTipo And oRsExoneradoBoleta.Fields!movNumero = rsReporte.Fields!movNumero And oRsExoneradoBoleta.Fields!idProducto = rsReporte.Fields!idProducto Then
                             lnExoneracionImpDetall = lnExoneracionImpDetall + oRsExoneradoBoleta.Fields!totalFinanciado
                             lnExoneracionCanDetall = lnExoneracionCanDetall + oRsExoneradoBoleta.Fields!CantidadFinanciada
                          End If
                          oRsExoneradoBoleta.MoveNext
                       Loop
                    End If
                    '
                    lnAnuladoImpDetall = 0: lnAnuladoCanDetall = 0
                    lnImptotalDetall = 0: lnCanTotalDetall = 0
                    
                    Select Case rsReporte.Fields!IdEstadoComprobante
                    Case 9      '***anulado
                        lnAnuladoImpDetall = rsReporte.Fields!TotalPagar - lnExoneracionImpDetall
                        lnAnuladoCanDetall = rsReporte.Fields!CantidadPagar
                        lnExoneracionCanDetall = 0: lnExoneracionImpDetall = 0
                        
                    Case Else   '***pagado
                        lnImptotalDetall = rsReporte.Fields!TotalPagar - lnExoneracionImpDetall
                        lnCanTotalDetall = rsReporte.Fields!CantidadPagar
                        If lnPagoCta > 0 Then
                           lnImptotalDetall = lnImptotalDetall - lnPagoCta
                           lnPagoCta = 0
                        End If
                    End Select
                    '
                    lnBoletaExo1 = lnBoletaExo1 + lnExoneracionImpDetall
                    lnBoletaTot1 = lnBoletaTot1 + lnImptotalDetall
                    lnBoletaAnu1 = lnBoletaAnu1 + lnAnuladoImpDetall
                    '
                    If ml_IdPartidaFiltro > 0 Then
                       GrabaDetalleEnTmpPartida ml_IdPartidaFiltro, rsReporte.Fields!IdPartida, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, rsReporte.Fields!idProducto, rsReporte.Fields!codigo, rsReporte.Fields!Producto, rsReporte.Fields!PrecioVenta, rsReporte.Fields!CantidadPagar, rsReporte.Fields!TotalPagar, lnTipoProceso
                    Else
                       mrs_Tmp.MoveFirst
                       mrs_Tmp.Find "idPartida=" & rsReporte.Fields!IdPartida
                       If Not mrs_Tmp.EOF Then
                          mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
                          mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
                          mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
                          mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
                          mrs_Tmp.Update
                       End If
                    End If
                    
                    '
                    rsReporte.MoveNext
                    If rsReporte.EOF Then
                       Exit Do
                    End If
              Loop
              'oRsExoneradoBoleta.Close
              If lnIdEstadoComprobante = 9 Then
                 lnBoletaExo1 = lnBoletaExo
              End If
              'Chequea Redondeo, para que cuadre con Totales de la  Boleta
              If lnBoletaExo <> lnBoletaExo1 Or lnBoletaTot <> lnBoletaTot1 Or lnBoletaAnu <> lnBoletaAnu1 Then
                    lnAnuladoImpDetall = 0
                    lnExoneracionImpDetall = 0
                    lnImptotalDetall = 0
                    If lnBoletaExo <> lnBoletaExo1 Then
                          lnExoneracionImpDetall = lnBoletaExo - lnBoletaExo1
                    End If
                    If lnBoletaTot <> lnBoletaTot1 Then
                          lnImptotalDetall = lnBoletaTot - lnBoletaTot1
                    End If
                    If lnBoletaAnu <> lnBoletaAnu1 Then
                          lnAnuladoImpDetall = lnBoletaAnu - lnBoletaAnu1
                    End If
                    If ml_IdPartidaFiltro > 0 Then
                       GrabaDetalleEnTmpPartida ml_IdPartidaFiltro, lnIdPartida100, mrs_Tmp, lnExoneracionImpDetall, lnExoneracionCanDetall, lnAnuladoImpDetall, lnAnuladoCanDetall, lnImptotalDetall, lnCanTotalDetall, lnIdProducto100, lcCodigo100, lcProducto100, lnPrecioVenta100, lnCantidadPagar100, lnTotalPagar100, lnTipoProceso
                    Else
                       mrs_Tmp.MoveFirst
                       mrs_Tmp.Find "idPartida=" & lnIdPartida100
                       If Not mrs_Tmp.EOF Then
                          mrs_Tmp.Fields!impAnulado = mrs_Tmp.Fields!impAnulado + lnAnuladoImpDetall
                          mrs_Tmp.Fields!impExonerado = mrs_Tmp.Fields!impExonerado + lnExoneracionImpDetall
                          mrs_Tmp.Fields!ImpNormal = mrs_Tmp.Fields!ImpNormal + lnImptotalDetall
                          mrs_Tmp.Fields!ImpCancelado = mrs_Tmp.Fields!ImpCancelado + lnImptotalDetall
                          mrs_Tmp.Update
                       End If
                    End If
              End If
              '
           Loop
        End If
        If mrs_Tmp.RecordCount > 0 Then
           mrs_Tmp.MoveFirst
           Do While Not mrs_Tmp.EOF
              mrs_Tmp!ImpSubTotal = mrs_Tmp!ImpNormal - (mrs_Tmp!impAnulado + mrs_Tmp!impExonerado)
              mrs_Tmp.Update
              mrs_Tmp.MoveNext
           Loop
        End If
        
        
        Set rsTmp = Nothing
        Set rsReporte = Nothing
        Set oRsExoneradoBoleta = Nothing
        Set oRsItemsXCuenta = Nothing
        Set rsParamet = Nothing
        Set rsTmp1 = Nothing
        Set rsTmp2 = Nothing
        Set lcBuscaParametro = Nothing
        'Set moConexion = Nothing
        Set mo_ReglasComunes = Nothing
        Set rsTmp999 = Nothing
        Set mo_ReglasCaja = Nothing
        Exit Sub
ErrRXCC:
        If Err.Number = 94 Then
           MsgBox Err.Description & Chr(13) & Chr(13) & "Boleta N° " & lcTexto2 & "     facturacionServiciosPagos.idProducto: " & Trim(str(lnIdProductoDetall))
        Else
           MsgBox Err.Description & Chr(13) & "Boleta: " & lcTexto2
        End If
        Exit Sub
        Resume

End Sub

Function FactPartidasPresupuestalesSeleccionarTodas(oConexion As Connection) As Recordset
'    Dim lcSql As String
    Dim oRs As New Recordset
    Dim oCommand As New ADODB.Command
    Set FactPartidasPresupuestalesSeleccionarTodas = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactPartidasPresupuestalesSeleccionarTodas"
        Set oRs = .Execute
    End With
            
    Set FactPartidasPresupuestalesSeleccionarTodas = oRs
    Set oCommand = Nothing

End Function

Function NotaDeCreditoDebitoServicios(lnIdCaja As Long, ldFechaIni As Date, ldFechaFin As Date, lnIdTurno As Long, lnCajero As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim lcSql As String
    lcSql = "SELECT     dbo.CajaComprobantesPago.NroSerie, dbo.CajaComprobantesPago.NroDocumento, dbo.CajaComprobantesPago.Total AS TotalBoleta, " & _
      "                dbo.CajaComprobantesPago.IdTipoOrden AS idTipoOrden, dbo.CajaComprobantesPago.IdTipoPago AS idTipoPago," & _
      "                dbo.CajaComprobantesPago.RazonSocial, dbo.NotaCreditoDebito.FechaPagado AS FechaCobranza, dbo.Pacientes.ApellidoPaterno AS apellidoPaterno," & _
      "                dbo.Pacientes.NroHistoriaClinica, dbo.Pacientes.ApellidoMaterno AS apellidoMaterno, dbo.Pacientes.PrimerNombre, dbo.CajaGestion.IdGestionCaja," & _
      "                dbo.CajaGestion.FechaApertura, dbo.CajaGestion.IdCaja AS idCaja, dbo.CajaGestion.IdTurno AS idTurno, dbo.CajaGestion.FechaCierre," & _
      "                dbo.CajaGestion.IdCajero AS idCajeroGestion, dbo.FactCatalogoServicios.Codigo AS codigo, dbo.FactCatalogoServicios.Nombre AS dproducto," & _
      "                dbo.CajaComprobantesPago.IdEstadoComprobante AS idEstadoComprobante, dbo.CajaComprobantesPago.Dctos," & _
      "                dbo.FacturacionServicioPagos.idOrdenPago, dbo.FacturacionServicioPagos.idProducto, dbo.FacturacionServicioPagos.Cantidad," & _
      "                dbo.FacturacionServicioPagos.Precio, dbo.FacturacionServicioPagos.Total, dbo.FactOrdenServicioPagos.idComprobantePago," & _
      "                dbo.FactOrdenServicioPagos.idOrden, dbo.CajaComprobantesPago.Exoneraciones, dbo.CajaComprobantesPago.SubTotal," & _
      "                dbo.CajaComprobantesPago.Total AS totalPagado, dbo.FactCatalogoServicios.IdCentroCosto, dbo.CajaComprobantesPago.Adelantos," & _
      "                dbo.FactCatalogoServicios.NombreMINSA, dbo.CajaComprobantesPago.IdTipoComprobante, dbo.CajaComprobantesPago.IdCuentaAtencion," & _
      "                dbo.FactCatalogoServicios.IdServicioGrupo , dbo.CajaComprobantesPago.IdTipoFinanciamiento, dbo.NotaCreditoDebito.IdCajero,NotaCreditoDebito.total as TotalNC" & _
      "   FROM         dbo.FactCatalogoServicios RIGHT OUTER JOIN" & _
      "                dbo.FacturacionServicioPagos ON dbo.FactCatalogoServicios.IdProducto = dbo.FacturacionServicioPagos.idProducto LEFT OUTER JOIN" & _
      "                dbo.FactOrdenServicioPagos LEFT OUTER JOIN" & _
      "                dbo.CajaComprobantesPago INNER JOIN" & _
      "                dbo.NotaCreditoDebito ON dbo.CajaComprobantesPago.IdComprobantePago = dbo.NotaCreditoDebito.IdComprobantePago ON" & _
      "                dbo.FactOrdenServicioPagos.idComprobantePago = dbo.CajaComprobantesPago.IdComprobantePago ON" & _
      "                dbo.FacturacionServicioPagos.idOrdenPago = dbo.FactOrdenServicioPagos.idOrdenPago LEFT OUTER JOIN" & _
      "                dbo.CajaGestion ON dbo.CajaComprobantesPago.IdGestionCaja = dbo.CajaGestion.IdGestionCaja LEFT OUTER JOIN" & _
      "                dbo.FacturacionCuentasAtencion ON" & _
      "                dbo.CajaComprobantesPago.IdCuentaAtencion = dbo.FacturacionCuentasAtencion.IdCuentaAtencion LEFT OUTER JOIN" & _
      "                dbo.Pacientes ON dbo.FacturacionCuentasAtencion.IdPaciente = dbo.Pacientes.IdPaciente"
    If lnIdCaja = 0 And lnIdTurno = 0 And lnCajero = 0 Then
       lcSql = lcSql & " Where (dbo.CajaComprobantesPago.IdTipoOrden = 1)" & _
                       "       and dbo.NotaCreditoDebito.FechaPagado Between (CONVERT(DATETIME,'" & Format(ldFechaIni, "dd/mm/yyyy hh:mm:ss") & "',103))  and  (CONVERT(DATETIME,'" & Format(ldFechaFin, "dd/mm/yyyy hh:mm:ss") & "',103))" & _
                       " ORDER BY dbo.NotaCreditoDebito.FechaPagado, dbo.CajaComprobantesPago.NroSerie, dbo.CajaComprobantesPago.NroDocumento"
    ElseIf lnIdCaja > 0 And lnIdTurno > 0 And lnCajero = 0 Then
       lcSql = lcSql & " Where    (dbo.CajaComprobantesPago.IdTipoOrden = 1)" & _
                       "       and dbo.NotaCreditoDebito.FechaPagado Between (CONVERT(DATETIME,'" & Format(ldFechaIni, "dd/mm/yyyy hh:mm:ss") & "',103))  and  (CONVERT(DATETIME,'" & Format(ldFechaFin, "dd/mm/yyyy hh:mm:ss") & "',103))" & _
                       "       and dbo.CajaComprobantesPago.idCaja=" & lnIdCaja & _
                       "       and dbo.CajaComprobantesPago.idTurno=" & lnIdTurno & _
                       " ORDER BY dbo.NotaCreditoDebito.FechaPagado, dbo.CajaComprobantesPago.NroSerie, dbo.CajaComprobantesPago.NroDocumento"
    Else
       lcSql = lcSql & " Where   (dbo.CajaComprobantesPago.IdTipoOrden = 1)" & _
                       "       and dbo.NotaCreditoDebito.FechaPagado Between (CONVERT(DATETIME,'" & Format(ldFechaIni, "dd/mm/yyyy hh:mm:ss") & "',103))  and  (CONVERT(DATETIME,'" & Format(ldFechaFin, "dd/mm/yyyy hh:mm:ss") & "',103))" & _
                       "       and dbo.CajaComprobantesPago.idCaja=" & lnIdCaja & _
                       "       and dbo.CajaComprobantesPago.idTurno=" & lnIdTurno & _
                       "       and dbo.NotaCreditoDebito.idCajero=" & lnCajero & _
                       " ORDER BY dbo.NotaCreditoDebito.FechaPagado, dbo.CajaComprobantesPago.NroSerie, dbo.CajaComprobantesPago.NroDocumento"
    End If
    oRecordset.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
    
    Set NotaDeCreditoDebitoServicios = oRecordset
    'Set oRecordset.ActiveConnection = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function
'debb-12/07/2019
Function NotaDeCreditoServiciosXfechas(ldFechaIni As Date, ldFechaFin As Date) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim lcSql As String
    lcSql = "SELECT     dbo.FacturacionServicioPagos.idOrdenPago, dbo.FacturacionServicioPagos.idProducto, dbo.FacturacionServicioPagos.Cantidad, " & _
"                      dbo.FacturacionServicioPagos.Precio AS PrecioUnitario, dbo.FacturacionServicioPagos.Total AS totalPorPagar, dbo.FactCatalogoServicios.Codigo," & _
"                      dbo.FactCatalogoServicios.Nombre AS NombreProducto, dbo.FactCatalogoServicios.Nombre, dbo.FacturacionServicioPagos.Precio," & _
"                      dbo.FacturacionServicioPagos.Total AS total, dbo.FacturacionServicioPagos.Total AS Importe," & _
"                      dbo.FacturacionServicioPagos.idProducto AS idProductoCPT, dbo.FactOrdenServicioPagos.idComprobantePago, dbo.FactOrdenServicioPagos.idOrden," & _
"                      dbo.FactOrdenServicioPagos.FechaCreacion, dbo.FactOrdenServicioPagos.IdEstadoFacturacion, dbo.Pacientes.ApellidoPaterno," & _
"                      dbo.Pacientes.ApellidoMaterno, dbo.Pacientes.PrimerNombre, dbo.Pacientes.NroHistoriaClinica, dbo.CajaComprobantesPago.NroSerie," & _
"                      dbo.CajaComprobantesPago.NroDocumento, dbo.CajaComprobantesPago.RazonSocial, dbo.CajaComprobantesPago.RUC," & _
"                      dbo.CajaComprobantesPago.Total AS totalboleta, dbo.NotaCreditoDebito.FechaPagado AS FechaCobranza," & _
"                      dbo.CajaComprobantesPago.IdTipoComprobante, dbo.CajaComprobantesPago.IdCuentaAtencion, dbo.CajaComprobantesPago.IdEstadoComprobante," & _
"                      dbo.CajaComprobantesPago.IdGestionCaja, dbo.CajaComprobantesPago.IdTipoPago, dbo.CajaComprobantesPago.IdTipoOrden," & _
"                      dbo.CajaComprobantesPago.Dctos, dbo.CajaComprobantesPago.IdPaciente, dbo.CajaComprobantesPago.IdCajero," & _
"                      dbo.CajaComprobantesPago.idTurno, dbo.CajaComprobantesPago.idCaja, dbo.CajaComprobantesPago.idFormaPago," & _
"                      dbo.CajaComprobantesPago.Exoneraciones, dbo.CajaComprobantesPago.Adelantos, dbo.CajaComprobantesPago.IGV," & _
"                      dbo.CajaComprobantesPago.Subtotal , dbo.FactCatalogoServicios.IdPartida" & _
" FROM         dbo.NotaCreditoDebito INNER JOIN" & _
"                      dbo.CajaComprobantesPago ON dbo.NotaCreditoDebito.IdComprobantePago = dbo.CajaComprobantesPago.IdComprobantePago RIGHT OUTER JOIN" & _
"                      dbo.FacturacionServicioPagos LEFT OUTER JOIN" & _
"                      dbo.FactCatalogoServicios ON dbo.FacturacionServicioPagos.idProducto = dbo.FactCatalogoServicios.IdProducto LEFT OUTER JOIN" & _
"                      dbo.FactOrdenServicioPagos ON dbo.FacturacionServicioPagos.idOrdenPago = dbo.FactOrdenServicioPagos.idOrdenPago ON" & _
"                      dbo.CajaComprobantesPago.IdComprobantePago = dbo.FactOrdenServicioPagos.idComprobantePago LEFT OUTER JOIN" & _
"                      dbo.Pacientes ON dbo.CajaComprobantesPago.IdPaciente = dbo.Pacientes.IdPaciente" & _
" where    dbo.NotaCreditoDebito.FechaPagado Between (CONVERT(DATETIME,'" & Format(ldFechaIni, "dd/mm/yyyy hh:mm:ss") & "',103))  and  (CONVERT(DATETIME,'" & Format(ldFechaFin, "dd/mm/yyyy hh:mm:ss") & "',103))"
   
    oRecordset.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
    
    Set NotaDeCreditoServiciosXfechas = oRecordset
    'Set oRecordset.ActiveConnection = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'debb-12/07/2019
Function NotaDeCreditoDebitoFarmacia(lnIdCaja As Long, ldFechaIni As Date, ldFechaFin As Date, lnIdTurno As Long, lnCajero As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim lcSql As String
    lcSql = "SELECT     dbo.FactCatalogoBienesInsumos.Codigo, dbo.FactCatalogoBienesInsumos.Nombre AS Producto, dbo.CajaComprobantesPago.NroSerie AS nroSerie, " & _
"                      dbo.CajaComprobantesPago.NroDocumento AS nroDocumento, dbo.CajaComprobantesPago.Total AS TotalBoleta," & _
"                      dbo.FactCatalogoBienesInsumos.IdPartida, dbo.CajaComprobantesPago.RazonSocial, dbo.NotaCreditoDebito.FechaPagado AS FechaCobranza," & _
"                      dbo.CajaComprobantesPago.IdEstadoComprobante, dbo.CajaComprobantesPago.IdTipoOrden, dbo.CajaGestion.FechaApertura," & _
"                      dbo.CajaGestion.IdCaja AS idCaja, dbo.CajaGestion.IdTurno AS idTurno, dbo.CajaGestion.FechaCierre, dbo.CajaComprobantesPago.IdCajero," & _
"                      dbo.FacturacionCuentasAtencion.IdPaciente, dbo.Pacientes.ApellidoPaterno, dbo.Pacientes.ApellidoMaterno, dbo.Pacientes.PrimerNombre," & _
"                      dbo.Pacientes.NroHistoriaClinica, dbo.CajaComprobantesPago.IdTipoPago AS idTipoPago, dbo.CajaComprobantesPago.Dctos," & _
"                      dbo.FacturacionBienesPagos.IdOrden, dbo.FacturacionBienesPagos.IdProducto, dbo.FacturacionBienesPagos.CantidadPagar," & _
"                      dbo.FacturacionBienesPagos.PrecioVenta, dbo.FacturacionBienesPagos.TotalPagar, dbo.FactOrdenesBienes.idComprobantePago," & _
"                      dbo.FactOrdenesBienes.MovNumero, dbo.FactOrdenesBienes.MovTipo, dbo.CajaComprobantesPago.Exoneraciones," & _
"                      dbo.CajaComprobantesPago.SubTotal, dbo.CajaComprobantesPago.Total AS totalPagado, dbo.FactCatalogoBienesInsumos.IdCentroCosto," & _
"                      dbo.CajaComprobantesPago.Adelantos , dbo.CajaComprobantesPago.idTipoComprobante, dbo.CajaComprobantesPago.IdTipoFinanciamiento,NotaCreditoDebito.total as TotalNC" & _
" FROM         dbo.CajaComprobantesPago INNER JOIN" & _
"                      dbo.CajaGestion ON dbo.CajaComprobantesPago.IdGestionCaja = dbo.CajaGestion.IdGestionCaja INNER JOIN" & _
"                      dbo.NotaCreditoDebito ON dbo.CajaComprobantesPago.IdComprobantePago = dbo.NotaCreditoDebito.IdComprobantePago RIGHT OUTER JOIN" & _
"                      dbo.FactOrdenesBienes ON dbo.CajaComprobantesPago.IdComprobantePago = dbo.FactOrdenesBienes.idComprobantePago RIGHT OUTER JOIN" & _
"                      dbo.FacturacionBienesPagos LEFT OUTER JOIN" & _
"                      dbo.FactCatalogoBienesInsumos ON dbo.FacturacionBienesPagos.IdProducto = dbo.FactCatalogoBienesInsumos.IdProducto ON" & _
"                      dbo.FactOrdenesBienes.idOrden = dbo.FacturacionBienesPagos.IdOrden LEFT OUTER JOIN" & _
"                      dbo.FacturacionCuentasAtencion ON" & _
"                      dbo.CajaComprobantesPago.IdCuentaAtencion = dbo.FacturacionCuentasAtencion.IdCuentaAtencion LEFT OUTER JOIN" & _
"                      dbo.Pacientes ON dbo.FacturacionCuentasAtencion.IdPaciente = dbo.Pacientes.IdPaciente"
    If lnIdCaja = 0 And lnIdTurno = 0 And lnCajero = 0 Then
       lcSql = lcSql & " Where  (dbo.CajaComprobantesPago.IdTipoOrden in (2,3) )" & _
                       "       and dbo.NotaCreditoDebito.FechaPagado Between (CONVERT(DATETIME,'" & Format(ldFechaIni, "dd/mm/yyyy hh:mm:ss") & "',103))  and  (CONVERT(DATETIME,'" & Format(ldFechaFin, "dd/mm/yyyy hh:mm:ss") & "',103))" & _
                       " ORDER BY dbo.NotaCreditoDebito.FechaPagado, dbo.CajaComprobantesPago.NroSerie, dbo.CajaComprobantesPago.NroDocumento"
    ElseIf lnIdCaja > 0 And lnIdTurno > 0 And lnCajero = 0 Then
       lcSql = lcSql & " Where (dbo.CajaComprobantesPago.IdTipoOrden in (2,3) )" & _
                       "       and dbo.NotaCreditoDebito.FechaPagado Between (CONVERT(DATETIME,'" & Format(ldFechaIni, "dd/mm/yyyy hh:mm:ss") & "',103))  and  (CONVERT(DATETIME,'" & Format(ldFechaFin, "dd/mm/yyyy hh:mm:ss") & "',103))" & _
                       "       and dbo.CajaComprobantesPago.idCaja=" & lnIdCaja & _
                       "       and dbo.CajaComprobantesPago.idTurno=" & lnIdTurno & _
                       " ORDER BY dbo.NotaCreditoDebito.FechaPagado, dbo.CajaComprobantesPago.NroSerie, dbo.CajaComprobantesPago.NroDocumento"
    Else
       lcSql = lcSql & " Where  (dbo.CajaComprobantesPago.IdTipoOrden in (2,3) )" & _
                       "       and dbo.NotaCreditoDebito.FechaPagado Between (CONVERT(DATETIME,'" & Format(ldFechaIni, "dd/mm/yyyy hh:mm:ss") & "',103))  and  (CONVERT(DATETIME,'" & Format(ldFechaFin, "dd/mm/yyyy hh:mm:ss") & "',103)" & _
                       "       and dbo.CajaComprobantesPago.idCaja=" & lnIdCaja & _
                       "       and dbo.CajaComprobantesPago.idTurno=" & lnIdTurno & _
                       "       and dbo.CajaComprobantesPago.idCajero=" & lnCajero & _
                       " ORDER BY dbo.NotaCreditoDebito.FechaPagado, dbo.CajaComprobantesPago.NroSerie, dbo.CajaComprobantesPago.NroDocumento"
    End If
    oRecordset.Open lcSql, sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
   ' Set oRecordset.ActiveConnection = Nothing
    Set NotaDeCreditoDebitoFarmacia = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function




