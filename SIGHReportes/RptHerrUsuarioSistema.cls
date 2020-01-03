VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptHerrUsuarioSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Auditoría en General
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mrs_Tmp As New ADODB.Recordset
Dim ml_IdResponsable As Long
Dim mo_ReporteUtil As New ReporteUtil
Dim ml_OrdenFiltro As String
Dim mda_FechaInicio As Date
Dim mda_FechaFin As Date
Dim ml_TextoDelFiltro  As String
Dim ml_TipoServicio  As Integer
Dim lcBuscaParametro As New SIGHDatos.Parametros
Public Event ProgressActualizaValor(ByRef lnValorActual As Long, ByRef lnValorTotal As Long)

Property Let TipoServicio(lValue As Integer)
    ml_TipoServicio = lValue
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
Property Let IdResponsable(lValue As Long)
    ml_IdResponsable = lValue
End Property
Property Let OrdenFiltro(lValue As String)
    ml_OrdenFiltro = lValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New HerrUsuarioSistema
    oFormulario.Show 1
End Sub

Sub CreaDatosParaReporte(lbEnExcel As Boolean, lcTitulo As String, lcSubTitulo As String, _
                ml_idUsuario As Long, mda_FechaInicio As Date, mda_FechaFin As Date, lnIdOpcion As Long, _
                lnHwnd As Long)
    On Error GoTo ErrRp
    Dim mrs_Tmp As New Recordset
    Dim mrs_Tmp1 As New Recordset
    Dim mrs_Tmp2 As New Recordset
    Dim mrs_Tmp3 As New Recordset
    Dim ldFecha As Date, ldFechaSolo As String
    Dim lcHoraSolo As String, lcUsuario As String, lcNusuario As String, lcAccion As String
    Dim lcPc As String, lcTexto1 As String, lcTexto2 As String, lcPie As String
    Dim lnIdListItem As Long, lcTexto3 As String
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
    Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
    Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    Dim mo_ReglasHoteleria As New SIGHNegocios.ReglasHoteleria
    Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
    Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim mo_ReglasDeSeguridad As New SIGHNegocios.ReglasDeSeguridad
    Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
    Dim lcTituloColumna As String, lbCambiaTituloColumna As Boolean
    Dim lnSalidasImg As Long
    Dim oDoImagMovimiento As New DoImagMovimiento
    Dim oDoImagMovimientoIngresos As New DoImagMovimientoIngresos
    Dim dFinanciamiento As String
    Dim mo_DoFarmMovimientoVentas As New DoFarmMovimientoVentas
    Dim lRecordCount As Long, f As Long
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    '
    Set mrs_Tmp1 = mo_ReglasFarmacia.AuditoriaPorFechasUsuario(ml_idUsuario, mda_FechaInicio, mda_FechaFin)
    If lnIdOpcion > 0 Then
       lbCambiaTituloColumna = True
       mrs_Tmp1.Filter = "idListItem=" & lnIdOpcion
       lcTituloColumna = "Paciente"
    Else
       mrs_Tmp1.Filter = "idListItem>0"
       lbCambiaTituloColumna = False
    End If
    lRecordCount = mrs_Tmp1.RecordCount
    If lRecordCount = 0 Then
         MsgBox "No existe Datos", vbInformation, "Reporte"
    Else
         With mrs_Tmp
            .Fields.Append "FechaCreacion", adVarChar, 10, adFldIsNullable         ' adDate, 10, adFldIsNullable
            .Fields.Append "HoraCreacion", adVarChar, 8, adFldIsNullable
            .Fields.Append "Accion", adVarChar, 30, adFldIsNullable
            .Fields.Append "Nusuario", adVarChar, 160, adFldIsNullable
            .Fields.Append "Usuario", adVarChar, 30, adFldIsNullable
            .Fields.Append "NombrePC", adVarChar, 30, adFldIsNullable
            .Fields.Append "Observacion1", adVarChar, 250, adFldIsNullable
            .Fields.Append "Observacion2", adVarChar, 250, adFldIsNullable
            .LockType = adLockOptimistic
            .Open
         End With
         '
         mrs_Tmp1.MoveFirst
         f = 0
         Do While Not mrs_Tmp1.EOF
           f = f + 1
           RaiseEvent ProgressActualizaValor(f, lRecordCount)
           '
           ldFecha = mrs_Tmp1.Fields!fechaHora
           ldFechaSolo = Format(mrs_Tmp1.Fields!fechaHora, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
           lcHoraSolo = Format(mrs_Tmp1.Fields!fechaHora, "hh:mm:ss")
           lnIdListItem = mrs_Tmp1.Fields!idListItem
           lcNusuario = Trim(mrs_Tmp1.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp1.Fields!ApellidoMaterno) & " " & mrs_Tmp1.Fields!Nombres
           lcUsuario = IIf(IsNull(mrs_Tmp1.Fields!Usuario), "", mrs_Tmp1.Fields!Usuario)
           lcAccion = IIf(mrs_Tmp1.Fields!Accion = "A", "Agregó", IIf(mrs_Tmp1.Fields!Accion = "M", "Modificó", IIf(mrs_Tmp1.Fields!Accion = "E", "Anuló", "")))      'franklin 2017
           lcPc = IIf(IsNull(mrs_Tmp1.Fields!NombrePC), "", mrs_Tmp1.Fields!NombrePC)
           Select Case mrs_Tmp1.Fields!idListItem
           Case 301      'Pacientes Hosp
                Set mrs_Tmp2 = mo_ReglasAdmision.PacientesSeleccionarPorIdentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre)
                   lcTexto2 = "(Mantenimiento de Pacientes desde Hospitalización)"
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = lcTexto1
                mrs_Tmp.Fields!Observacion2 = lcTexto2
                mrs_Tmp.Update
           Case 302, 1330     'Admision Hosp/Alojados
                If UCase(Trim(mrs_Tmp1.Fields!Tabla)) = "FUA" Then
                    lcTexto1 = mrs_Tmp1.Fields!Tabla
                    lcTexto2 = mrs_Tmp1.Fields!observaciones
                Else
                    Set mrs_Tmp2 = mo_ReglasAdmision.AtencionesSeleccionarPorIdAtencionTipoServicio(mrs_Tmp1.Fields!idRegistro, 3, oConexion)
                    lcTexto1 = mrs_Tmp1.Fields!Tabla
                    lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                    If mrs_Tmp2.RecordCount > 0 Then
                       lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre) & " (" & Trim(mrs_Tmp1.Fields!observaciones) & ")"
                       lcTexto2 = "Adm: (Fecha: " & mrs_Tmp2.Fields!FechaIngreso & ") (N° Cuenta: " & Trim(Str(mrs_Tmp2.Fields!idCuentaAtencion)) & ") (Serv: " & Trim(mrs_Tmp2.Fields!DServicio) & ")"
                    End If
                    mrs_Tmp2.Close
                End If
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 303      'camas
                Set mrs_Tmp2 = mo_ReglasHoteleria.CamasSeleccionarPorIdentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "(Mantenimiento de Camas desde Hospitalización)"
                   lcTexto2 = "(Nro Cama: " & Trim(mrs_Tmp2.Fields!Codigo) & ") (Servicio:" & Trim(mrs_Tmp2.Fields!DServicio) & ")"
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 601     'General - Reg.Procedimientos
                Set mrs_Tmp2 = mo_ReglasAdmision.FactOrdenServicioSeleccionarPorIdOrden(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre)
                   lcTexto2 = "Cpt: (Fecha: " & Format(mrs_Tmp2.Fields!fechaCreacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY) & ") (N° Cuenta: " & Trim(Str(mrs_Tmp2.Fields!idCuentaAtencion)) & ") (Serv: " & Trim(mrs_Tmp2.Fields!DServicio) & ")"
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 502      'Movimientos HC
                Set mrs_Tmp2 = mo_ReglasArchivoClinico.MovimientosHistoriaClinicaSeleccionarPorIdentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 And Not IsNull(mrs_Tmp2.Fields!NroHistoriaClinica) Then
                   lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre)
                   If mrs_Tmp2.Fields!idMotivo = 9 Then
                      lcTexto2 = "(Llegada al A.C.)  desde: " & mrs_Tmp2.Fields!sOrigen
                   Else
                      lcTexto2 = "(Salida del A.C.)  hacia: " & mrs_Tmp2.Fields!sDestino
                   End If
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 1332      'Movimientos Formatos
                Set mrs_Tmp2 = mo_ReglasArchivoClinico.MovimientosFormatosSeleccionarPorIdentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre)
                   If mrs_Tmp2.Fields!idMotivo = 9 Then
                      lcTexto2 = "(Llegada Formato al A.C.)  desde: " & mrs_Tmp2.Fields!sOrigen
                   Else
                      lcTexto2 = "(Salida Formato del A.C.)  hacia: " & mrs_Tmp2.Fields!sDestino
                   End If
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 501      'Archivo Clinico
                Set mrs_Tmp2 = mo_ReglasArchivoClinico.HistoriaClinicasSeleccionarXnrohistoria(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre)
                   lcTexto2 = "(Mantenimiento de H.C.)    "
                   If Not IsNull(mrs_Tmp2.Fields!nroHistoriaClinicaAnterior) Then
                      lcTexto2 = lcTexto2 & "    " & "Nro anterior: " & Trim(Str(mrs_Tmp2.Fields!nroHistoriaClinicaAnterior))
                   Else
                      lcTexto2 = lcTexto2 & mrs_Tmp1.Fields!observaciones
                   End If
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 101      'Pacientes CE
                Set mrs_Tmp2 = mo_ReglasAdmision.PacientesSeleccionarPorIdentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre)
                   lcTexto2 = "(Mantenimiento de Pacientes desde CE)"
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 102      'Citas
                Set mrs_Tmp2 = mo_ReglasAdmision.AtencionesSeleccionarPorIdAtencionTipoServicio(mrs_Tmp1.Fields!idRegistro, 1, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 And Not IsNull(mrs_Tmp2.Fields!NroHistoriaClinica) Then
                   lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre)
                   lcTexto2 = "Reg.Cita: (Fecha: " & mrs_Tmp2.Fields!FechaIngreso & ") (N° Cuenta: " & Trim(Str(mrs_Tmp2.Fields!idCuentaAtencion)) & ") (Serv: " & Trim(mrs_Tmp2.Fields!DServicio) & ")"
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 103     'Atenciones Citas
                If UCase(Trim(mrs_Tmp1.Fields!Tabla)) = "FUA" Then
                    lcTexto1 = mrs_Tmp1.Fields!Tabla
                    lcTexto2 = mrs_Tmp1.Fields!observaciones
                Else
                    Set mrs_Tmp2 = mo_ReglasAdmision.AtencionesSeleccionarPorIdAtencionTipoServicio(mrs_Tmp1.Fields!idRegistro, 1, oConexion)
                    lcTexto1 = mrs_Tmp1.Fields!Tabla
                    lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                    If mrs_Tmp2.RecordCount > 0 And Not IsNull(mrs_Tmp2.Fields!NroHistoriaClinica) Then
                       lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre) & " (Cta: " & Trim(mrs_Tmp2.Fields!idCuentaAtencion) & ")"
                       lcTexto2 = "Reg.Atenc.Cita: " & IIf(IsNull(mrs_Tmp1!observaciones), "(Fecha: " & mrs_Tmp2.Fields!FechaIngreso & ") (N° Cuenta: " & Trim(Str(mrs_Tmp2.Fields!idCuentaAtencion)) & ") (Serv: " & Trim(mrs_Tmp2.Fields!DServicio) & ")", mrs_Tmp1!observaciones)
                    End If
                    mrs_Tmp2.Close
                End If
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 613      'Estado de Cuenta
                If UCase(Trim(mrs_Tmp1.Fields!Tabla)) = "FUA" Then
                    lcTexto1 = mrs_Tmp1.Fields!Tabla
                    lcTexto2 = mrs_Tmp1.Fields!observaciones
                Else
                    lcAccion = Trim(mrs_Tmp1.Fields!Tabla)
                    Set mrs_Tmp2 = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(mrs_Tmp1.Fields!idRegistro, oConexion)
                    lcTexto1 = Trim(mrs_Tmp1.Fields!Tabla) & "  N°: " & Trim(Str(mrs_Tmp1.Fields!idRegistro))
                    If IsNull(mrs_Tmp1.Fields!observaciones) Then
                        lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                        If mrs_Tmp2.RecordCount > 0 Then
                           lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre)
                           lcTexto2 = "E.Cta: " & Trim(Str(mrs_Tmp2.Fields!idCuentaAtencion)) & "  (" & Trim(mrs_Tmp2.Fields!EstadoCta) & ") (" & Trim(mrs_Tmp2.Fields!dFuenteFinanciamiento) & ")"
                        End If
                    Else
                       lcTexto2 = "E.Cta: " & Trim(mrs_Tmp1.Fields!observaciones)
                    End If
                    mrs_Tmp2.Close
                End If
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 702      'Gestion de caja
                Set mrs_Tmp2 = mo_ReglasCaja.CajaComprobantesSeleccionarPorId(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = Trim(mrs_Tmp2.Fields!razonSocial)
                   lcPie = ""
                   If Not IsNull(mrs_Tmp2.Fields!idCuentaAtencion) Then
                      lcPie = Trim(Str(mrs_Tmp2.Fields!idCuentaAtencion))
                   End If
                   lcTexto2 = "Caja: (N°Dcto: " & Trim(mrs_Tmp2.Fields!nroSerie) & "-" & Trim(mrs_Tmp2.Fields!NroDocumento) & ") (Fecha: " & mrs_Tmp2.Fields!fechaCobranza & ") (N°Cuenta: " & lcPie & ")"
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 201      'Pacientes Emerg
                Set mrs_Tmp2 = mo_ReglasAdmision.PacientesSeleccionarPorIdentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre) & " (" & Trim(mrs_Tmp1.Fields!observaciones) & ")"
                   lcTexto2 = "(Mantenimiento de Pacientes desde Emergencia)"
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 202      'Admision Emerg
                If UCase(Trim(mrs_Tmp1.Fields!Tabla)) = "FUA" Then
                    lcTexto1 = mrs_Tmp1.Fields!Tabla
                    lcTexto2 = mrs_Tmp1.Fields!observaciones
                Else
                    Set mrs_Tmp2 = mo_ReglasAdmision.AtencionesSeleccionarPorIdAtencionTipoServicio(mrs_Tmp1.Fields!idRegistro, 2, oConexion)
                    lcTexto1 = mrs_Tmp1.Fields!Tabla
                    lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                    If mrs_Tmp2.RecordCount > 0 Then
                       lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre) & " (" & Trim(mrs_Tmp1.Fields!observaciones) & ")"
                       lcTexto2 = "Adm: (Fecha: " & mrs_Tmp2.Fields!FechaIngreso & ") (N° Cuenta: " & Trim(Str(mrs_Tmp2.Fields!idCuentaAtencion)) & ") (Serv: " & Trim(mrs_Tmp2.Fields!DServicio) & ")"
                    End If
                    mrs_Tmp2.Close
                End If
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 401      'Programacion
                lcTituloColumna = "Médico"
                Set mrs_Tmp2 = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarXidentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!Nombres)
                   lcTexto2 = "Prg.Citas: (Fecha: " & mrs_Tmp2.Fields!Fecha & ") (Serv: " & Trim(mrs_Tmp2.Fields!DServicio) & ")"
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 402      'Turno
                lcTituloColumna = "Turno"
                Set mrs_Tmp2 = mo_ReglasDeProgMedica.TurnosSeleccionarPorIdentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = Trim(mrs_Tmp2.Fields!Codigo) & " " & Trim(mrs_Tmp2.Fields!descripcion)
                   lcTexto2 = "Mantenimiento de Turnos"
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 403      'Profesionales de salud
                lcTituloColumna = "Médico"
                Set mrs_Tmp2 = mo_ReglasDeProgMedica.MedicosSeleccionarXidentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!Nombres)
                   lcTexto2 = "Mantenimiento de Médicos"
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 609      'Centro Costo
                lcTituloColumna = "Centro de Costo"
                Set mrs_Tmp2 = mo_reglasComunes.centrosCostosSeleccionarXidentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = Trim(mrs_Tmp2.Fields!Codigo) & " " & Trim(mrs_Tmp2.Fields!descripcion)
                   lcTexto2 = "Mantenimiento de Centros Costos"
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 610      'Cat.Servicios
                lcTituloColumna = "Servicios"
                Set mrs_Tmp2 = mo_reglasComunes.CatalogoServiciosSeleccionarXidentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = Trim(mrs_Tmp2.Fields!Codigo) & " " & Trim(mrs_Tmp2.Fields!Nombre)
                   lcTexto2 = "Mantenimiento de Servicios (CPT)"
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 611      'Cat.Tipo Financ
                lcTituloColumna = "Tipo Financiamiento"
                Set mrs_Tmp2 = mo_ReglasFacturacion.TiposFinanciamientoXidentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = Trim(Str(mrs_Tmp2.Fields!idTipoFinanciamiento)) & " " & Trim(mrs_Tmp2.Fields!descripcion)
                   lcTexto2 = "Mantenimiento de Tipos Financiamientos"
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 612     'Cat.Partidas
                lcTituloColumna = "Partida Presupuestal"
                Set mrs_Tmp2 = mo_reglasComunes.FactPartidasPresupuestalesXidentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = Trim(mrs_Tmp2.Fields!Codigo) & " " & Trim(mrs_Tmp2.Fields!descripcion)
                   lcTexto2 = "Mantenimiento de Partidas Presupuestales"
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 803      'Cat.Bienes
                lcTituloColumna = "Medicamento e Insumo"
                Set mrs_Tmp2 = mo_reglasComunes.CatalogoBienesInsumosSeleccionarPorFiltro("idProducto=" & mrs_Tmp1.Fields!idRegistro, sghPorDescripcion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = Trim(mrs_Tmp2.Fields!Codigo) & " " & Trim(mrs_Tmp2.Fields!Nombre)
                   lcTexto2 = "Mantenimiento de Medicamentos e Insumos"
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 1201      'Cat.Servicios (areas Hosp)
                lcTituloColumna = "Servicio"
                Set mrs_Tmp2 = mo_reglasComunes.ServiciosSeleccionarXidentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = Trim(Str(mrs_Tmp2.Fields!idServicio)) & " " & Trim(mrs_Tmp2.Fields!Nombre)
                   lcTexto2 = "Mantenimiento de Servicios"
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 1204      'Cat.Establec.NO MINSA
                lcTituloColumna = "Establecimiento NO MINSA"
                Set mrs_Tmp2 = mo_reglasComunes.EstablecimientosNoMinsaXidentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = Trim(Str(mrs_Tmp2.Fields!idEstablecimientoNoMinsa)) & " " & Trim(mrs_Tmp2.Fields!Nombre)
                   lcTexto2 = "Mantenimiento de Establec. NO MINSA"
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 1206      'Cat.Especialidades
                lcTituloColumna = "Especialidad"
                Set mrs_Tmp2 = mo_reglasComunes.EspecialidadesSeleccionarPorFiltro("idEspecialidad=" & mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = Trim(Str(mrs_Tmp2.Fields!IdEspecialidad)) & " " & Trim(mrs_Tmp2.Fields!Nombre)
                   lcTexto2 = "Mantenimiento de Especialidades"
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 1301      'Cat.Empleados
                lcTituloColumna = "Empleado"
                Set mrs_Tmp2 = mo_reglasComunes.EmpleadosSeleccionarPorIdEmpleado(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = Trim(Str(mrs_Tmp2.Fields!IdEmpleado)) & " " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!Nombres)
                   lcTexto2 = IIf(IsNull(mrs_Tmp1.Fields!observaciones), "Mantenimiento de Empleados", mrs_Tmp1.Fields!observaciones)
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 1302      'Roles
                lcTituloColumna = "Rol"
                Set mrs_Tmp2 = mo_ReglasDeSeguridad.RolesSeleccionarXidentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = Trim(Str(mrs_Tmp2.Fields!idRol)) & " " & Trim(mrs_Tmp2.Fields!Nombre)
                   lcTexto2 = "Mantenimiento de Roles"
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 1321    'Anat.patologica
                lcTituloColumna = "Paciente"
                Set mrs_Tmp2 = mo_ReglasLaboratorio.LabMovimientoSeleccionarXidentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre)
                   lcTexto2 = "Anat.patologica:  (N° cuenta: " & Trim(Str(mrs_Tmp2.Fields!idCuentaAtencion)) & ")"
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 1322     ' Banco sangre
                lcTituloColumna = "Paciente"
                Set mrs_Tmp2 = mo_ReglasLaboratorio.LabMovimientoSeleccionarXidentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre)
                   lcTexto2 = "Banco Sangre:  (N° cuenta: " & Trim(Str(mrs_Tmp2.Fields!idCuentaAtencion)) & ")"
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 1312     'pat.clinica
                lcTituloColumna = "Paciente"
                Set mrs_Tmp2 = mo_ReglasLaboratorio.LabMovimientoSeleccionarXidentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   If Not IsNull(mrs_Tmp2.Fields!NroHistoriaClinica) Then
                        lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre)
                        lcTexto2 = "Pat.Clínica:  (N° cuenta: " & Trim(Str(mrs_Tmp2.Fields!idCuentaAtencion)) & ")"
                   End If
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 1314     'salidas Laboratorio
                lcTituloColumna = "Documento"
                Set mrs_Tmp2 = mo_ReglasLaboratorio.LabMovimientoIngresosXidMovimiento(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!NroDocumento)) & ") "
                   lcTexto2 = "Mantenimiento de Salidas Laboratorio"
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 1313     'Ingresos Laboratorio
                lcTituloColumna = "Documento"
                Set mrs_Tmp2 = mo_ReglasLaboratorio.LabMovimientoSalidasXidMovimiento(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!motivo)) & ") "
                   lcTexto2 = "Mantenimiento de Ingresos Laboratorio"
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 1315, 1316, 1317, 1318, 1319, 1320 'Imagenes
                lcTituloColumna = "Paciente"
                lcTexto3 = ""
                If UCase(Left(mrs_Tmp1.Fields!Tabla, 14)) = "IMAGMOVIMIENTO" Then
                        lcTexto1 = Mid(mrs_Tmp1.Fields!Tabla, 16, 1)  'MOVTIPO
                        lcTexto2 = ""
                        lnSalidasImg = Val(Mid(mrs_Tmp1.Fields!Tabla, 18, 15))   'NumeroMovimiento
                        Set oDoImagMovimiento = mo_ReglasImagenes.ImagMovimientoSeleccionarPorId(lnSalidasImg, oConexion)
                        If lcTexto1 = "E" Then
                            'Ingresos
                            oDoImagMovimientoIngresos.NroDocumento = ""
                            Set oDoImagMovimientoIngresos = mo_ReglasImagenes.ImagMovimientoIngresosSeleccionarPorId(lnSalidasImg, oConexion)
                            If Not IsNull(oDoImagMovimientoIngresos.NroDocumento) Then
                                lcTexto2 = "N° Documento: " & oDoImagMovimientoIngresos.NroDocumento
                            End If
                        Else
                            If oDoImagMovimiento.IdTipoConcepto = 2 Then
                               'salidas por Deterioro/vencimiento/etc
                                Set mrs_Tmp2 = mo_ReglasImagenes.ImagMovimientoSalidasSeleccionarPorIdMovimiento(lnSalidasImg, oConexion)
                                If mrs_Tmp2.RecordCount > 0 Then
                                   lcTexto2 = "Motivo: " & mrs_Tmp2.Fields!motivo
                                End If
                                mrs_Tmp2.Close
                            Else
                               'salida  por toma de imagen
                               Set mrs_Tmp2 = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorIdMovimiento(lnSalidasImg, oConexion)
                               If mrs_Tmp2.RecordCount > 0 Then
                                  Set mrs_Tmp3 = mo_reglasComunes.FactPuntosCargaSeleccionarPorId(mrs_Tmp2.Fields!IdPuntoCarga, oConexion)
                                  If mrs_Tmp3.RecordCount > 0 Then
                                     lcTexto3 = Trim(mrs_Tmp3.Fields!descripcion)
                                  End If
                                  mrs_Tmp3.Close
                                  If Not IsNull(mrs_Tmp2.Fields!NroHistoriaClinica) Then
                                     lcTexto2 = Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)) & " " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & mrs_Tmp2.Fields!PrimerNombre
                                  End If
                               End If
                               mrs_Tmp2.Close
                            End If
                        End If
                        Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                           mrs_Tmp1.MoveNext
                           If mrs_Tmp1.EOF Then
                              Exit Do
                           End If
                        Loop
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                        mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                        mrs_Tmp.Fields!Accion = lcAccion
                        mrs_Tmp.Fields!Nusuario = lcNusuario
                        mrs_Tmp.Fields!Usuario = lcUsuario
                        mrs_Tmp.Fields!NombrePC = lcPc
                        mrs_Tmp.Fields!Observacion1 = Left(lcTexto2, 250)
                        mrs_Tmp.Fields!Observacion2 = "Imag: " & IIf(lcTexto1 = "E", "(Entrada)", "(Salida)") & IIf(lcTexto3 = "", "", "    (Pto.Carga: " & lcTexto3 & ")")
                        mrs_Tmp.Update
                Else
                        mrs_Tmp1.MoveNext
                End If
           Case 801, 1304, 1305, 1307, 1308, 1310, 1358, 1357 'Farmacia    debb2014
                
                    lcTituloColumna = "Descripción"
                    lcTexto1 = "Id: " & Trim(Str(mrs_Tmp1.Fields!idRegistro))
                    lcTexto2 = "Farmacia"
                    If UCase(Left(mrs_Tmp1.Fields!Tabla, 11)) = "FARMALMACEN" Then
                        Set mrs_Tmp2 = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idAlmacen=" & Trim(Str(mrs_Tmp1.Fields!idRegistro)))
                        mrs_Tmp2.Filter = "idAlmacen=" & mrs_Tmp1.Fields!idRegistro
                        If mrs_Tmp2.RecordCount > 0 Then
                           lcTexto1 = "Farmacia"
                           If mrs_Tmp1.Fields!idListItem = 1310 Then
                              lcTexto2 = "Depend.Externa: " & mrs_Tmp2!descripcion
                           Else
                              lcTexto2 = "Almacén/Farmacia: " & mrs_Tmp2!descripcion
                           End If
                        End If
                        mrs_Tmp2.Close
                    ElseIf UCase(Left(mrs_Tmp1.Fields!Tabla, 14)) = "FARMINVENTARIO" Then
                        Set mrs_Tmp2 = mo_ReglasFarmacia.FarmInventarioDevuelveLista
                        mrs_Tmp2.Filter = "idInventario=" & mrs_Tmp1.Fields!idRegistro
                        If mrs_Tmp2.RecordCount > 0 Then
                           lcTexto1 = "Farmacia"
                           lcTexto2 = "N° Inventario: " & mrs_Tmp2!numeroInventario
                        End If
                        mrs_Tmp2.Close
                    ElseIf UCase(Left(mrs_Tmp1.Fields!Tabla, 14)) = "FARMMOVIMIENTO" Then
                        lcTexto1 = Mid(mrs_Tmp1.Fields!Tabla, 16, 1)
                        lcTexto2 = Right(Trim(mrs_Tmp1.Fields!Tabla), 9)
                        dFinanciamiento = ""
                        Set mrs_Tmp2 = mo_ReglasFarmacia.FarmMovimientoFiltraPorMovNumeroTipo(lcTexto2, lcTexto1, oConexion)
                        If mrs_Tmp2.RecordCount > 0 Then
                            If mrs_Tmp2.Fields!IdTipoConcepto = 10 Then  'Contado
                                mo_DoFarmMovimientoVentas.movNumero = lcTexto2
                                mo_DoFarmMovimientoVentas.MovTipo = lcTexto1
                                If mo_ReglasFarmacia.FarmMovimientoVentasSeleccionarPorId(mo_DoFarmMovimientoVentas, oConexion) Then
                                   dFinanciamiento = "(Preventa: " & mo_DoFarmMovimientoVentas.idPreVenta & ")"
                                End If
                            End If
                            lcTexto1 = "Mov: " & lcTexto1 & "-" & lcTexto2 & "   Concepto: " & Trim(mrs_Tmp2.Fields!Concepto)
                            lcTexto1 = lcTexto1 & " (" & Trim(mrs_Tmp2.Fields!abreviatura) & " " & Trim(mrs_Tmp2.Fields!DocumentoNumero) & ")"
                            lcTexto2 = IIf(IsNull(mrs_Tmp1!observaciones), "Farm:  (Documento: " & Trim(mrs_Tmp2.Fields!abreviatura) & " " & Trim(mrs_Tmp2.Fields!DocumentoNumero) & ")", mrs_Tmp1!observaciones)
                        End If
                        mrs_Tmp2.Close
                    ElseIf UCase(Trim(mrs_Tmp1.Fields!Tabla)) = "FARMPREVENTA" Then
                        Set mrs_Tmp2 = mo_ReglasFarmacia.FarmPreVentaFiltraPorIdPreVenta(mrs_Tmp1.Fields!idRegistro, oConexion)
                        If mrs_Tmp2.RecordCount > 0 Then
                            lcTexto1 = "N° Preventa: " & Trim(Str(mrs_Tmp1.Fields!idRegistro))
                            lcTexto2 = IIf(IsNull(mrs_Tmp1!observaciones), "Farm:  (Mantenimiento de PreVenta)", mrs_Tmp1!observaciones)
                        End If
                        mrs_Tmp2.Close
                    End If
                    Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                       mrs_Tmp1.MoveNext
                       If mrs_Tmp1.EOF Then
                          Exit Do
                       End If
                    Loop
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                    mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                    mrs_Tmp.Fields!Accion = lcAccion
                    mrs_Tmp.Fields!Nusuario = lcNusuario
                    mrs_Tmp.Fields!Usuario = lcUsuario
                    mrs_Tmp.Fields!NombrePC = lcPc
                    mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                    mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                    mrs_Tmp.Update

           Case 1325   'Constancias
                lcTituloColumna = "Paciente"
                Set mrs_Tmp2 = mo_ReglasLaboratorio.ConstanciasSeleccionarXidentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre)
                   lcTexto2 = "Constancia N° " & mrs_Tmp1.Fields!idRegistro
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 1331   'Reembolsos
                lcTituloColumna = "Documento"
                Set mrs_Tmp2 = mo_ReglasFacturacion.FactReembolsosSeleccionarXidentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "(Año: " & Trim(Str(mrs_Tmp2.Fields!Anio)) & "    mes:" & Trim(Str(mrs_Tmp2.Fields!Mes)) & ") " & Trim(mrs_Tmp2.Fields!descripcion)
                   lcTexto2 = "Reembolso N° " & mrs_Tmp1.Fields!idRegistro & "     Area T.S: " & Trim(mrs_Tmp2.Fields!AreaTS) & "     IAFA: " & mrs_Tmp2.Fields!Plan
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 1303  'Triaje
                lcTituloColumna = "Triaje"
                Set mrs_Tmp2 = mo_reglasComunes.AtencionesSeleccionarXidentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = mrs_Tmp1.Fields!Tabla
                lcTexto2 = "Id: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "N° Cuenta: " & Trim(Str(mrs_Tmp2.Fields!idCuentaAtencion))
                   lcTexto2 = "Triaje: (" & Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre)
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 1343, 1344 'Recetas Automaticas Emerg/Hosp
                lcTituloColumna = "Recetas"
                lcTexto1 = "N° Cuenta: " & Trim(Str(mrs_Tmp1.Fields!idRegistro)) & IIf(mrs_Tmp1.Fields!idListItem = 1343, " (Emergencia)", " (Hospitalización)")
                lcTexto2 = "(Recetas) " & IIf(IsNull(mrs_Tmp1.Fields!observaciones), "", Trim(mrs_Tmp1.Fields!observaciones))
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case Else
                'Menu Herramientas
                    lcTexto1 = Trim(mrs_Tmp1.Fields!Tabla)
                    If UCase(lcTexto1) = "REPORTES" Then
                       lcTexto2 = IIf(IsNull(mrs_Tmp1.Fields!observaciones), "", Trim(mrs_Tmp1.Fields!observaciones))
                    Else
                       lcTexto2 = "(Herramientas) " & IIf(IsNull(mrs_Tmp1.Fields!observaciones), "", Trim(mrs_Tmp1.Fields!observaciones))
                    End If
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                    mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                    mrs_Tmp.Fields!Accion = lcAccion
                    mrs_Tmp.Fields!Nusuario = lcNusuario
                    mrs_Tmp.Fields!Usuario = lcUsuario
                    mrs_Tmp.Fields!NombrePC = lcPc
                    mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                    mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                    mrs_Tmp.Update
                
                    mrs_Tmp1.MoveNext
           End Select
          
        Loop
        lcPie = "Nro Movimientos: " & Trim(Str(mrs_Tmp.RecordCount))
        If lbEnExcel = True Then
           Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
           'mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, lcSubTitulo, lcPie, lnHwnd
           mo_ReglasReportes.ExportarRecordSetAexcelFast mrs_Tmp, lcTitulo, lcSubTitulo, lcPie, lnHwnd
           Set mo_ReglasReportes = Nothing
        Else
            Set RpAHCauditoria.DataSource = mrs_Tmp
            RpAHCauditoria.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
            RpAHCauditoria.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
            RpAHCauditoria.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
            RpAHCauditoria.Sections("cabecera").Controls("lblFecha").Caption = "Fecha: " & lcBuscaParametro.RetornaFechaServidorSQL
            RpAHCauditoria.Sections("cabecera").Controls("lblHora").Caption = "Hora: " & lcBuscaParametro.RetornaHoraServidorSQL
            
            RpAHCauditoria.Sections("cabecera").Controls("lblTitulo").Caption = lcTitulo
            RpAHCauditoria.Sections("cabecera").Controls("lblSubTitulo").Caption = lcSubTitulo
            
            If lbCambiaTituloColumna = True Then
               RpAHCauditoria.Sections("cabecera").Controls("lblPaciente").Caption = lcTituloColumna
            Else
               RpAHCauditoria.Sections("cabecera").Controls("lblPaciente").Caption = "Descripción"
            End If
            Set RpAHCauditoria.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
            RpAHCauditoria.Sections("pie").Controls("lblPie").Caption = lcPie
            RpAHCauditoria.Orientation = rptOrientLandscape
            RpAHCauditoria.Show 1
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
    Exit Sub
ErrRp:
    MsgBox Err.Description
    Exit Sub
    Resume
End Sub



