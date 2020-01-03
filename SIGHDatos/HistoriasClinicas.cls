VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistoriasClinicas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla HistoriasClinicas
'        Programado por: Castro W
'        Fecha: Noviembre 2006
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Conexion As ADODB.Connection
Dim ms_MensajeError As String
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property

'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla HistoriasClinicas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "HistoriasClinicasAgregar"
           Set oParameter = .CreateParameter("@IdTipoNumeracionAnterior", adInteger, adParamInput, 0, IIf(oTabla.IdTipoNumeracionAnterior = 0, Null, oTabla.IdTipoNumeracionAnterior)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroHistoriaClinicaAnterior", adInteger, adParamInput, 0, IIf(oTabla.NroHistoriaClinicaAnterior = 0, Null, oTabla.NroHistoriaClinicaAnterior)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, IIf(oTabla.IdTipoNumeracion = 0, Null, oTabla.IdTipoNumeracion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, IIf(oTabla.NroHistoriaClinica = 0, Null, oTabla.NroHistoriaClinica)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaPasoAPasivo", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaPasoAPasivo = 0, Null, oTabla.FechaPasoAPasivo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoHistoria", adInteger, adParamInput, 0, IIf(oTabla.IdTipoHistoria = 0, Null, oTabla.IdTipoHistoria)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoHistoria", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoHistoria = 0, Null, oTabla.IdEstadoHistoria)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
          ' oTabla.NroHistoriaClinica = .Parameters("@NroHistoriaClinica")
   End With
 
   Insertar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
'   If lnNroErrores = 0 And Err.Number = -2147217873 Then
'      'El idPaciente ya existe
'      Dim oPacientes As New Pacientes
'      Set oPacientes.Conexion = mo_Conexion
'      oPacientes.ActualizaCorrelativosDeNroHistoria
'      Set oPacientes = Nothing
'      lnNroErrores = lnNroErrores + 1
'      Resume
'   End If
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
   GrabaHistorico oTabla
Exit Function
End Function


Sub GrabaHistorico(oTabla As DOHistoriaClinica)
    On Error GoTo ErrHistorico
    Dim lcSql As String
    Dim oRsParametros As New Recordset
    Dim oRsMDB As New Recordset
    Dim oConexionMDB As New Connection
    lcSql = "select ValorTexto from Parametros where idparametro=581"
    oRsParametros.Open lcSql, mo_Conexion, adOpenKeyset, adLockOptimistic
    oConexionMDB.CommandTimeout = 900
    oConexionMDB.CursorLocation = adUseClient
    oConexionMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =" & Trim(oRsParametros!ValorTexto) & "\parametros.mdb"
    lcSql = "INSERT INTO HistoriasClinicas" & _
"           (NroHistoriaClinica,FechaCreacion" & _
"           ,FechaPasoAPasivo,IdTipoHistoria" & _
"           ,IdEstadoHistoria,IdPaciente" & _
"           ,IdTipoNumeracion,NroHistoriaClinicaAnterior" & _
"           ,IdTipoNumeracionAnterior)" & _
"     Values" & _
"           (" & _
"           " & oTabla.NroHistoriaClinica & ",'" & oTabla.FechaCreacion & _
"           ','" & oTabla.FechaPasoAPasivo & "'," & oTabla.IdTipoHistoria & _
"           ," & oTabla.IdEstadoHistoria & "," & oTabla.IdPaciente & _
"           ," & oTabla.IdTipoNumeracion & "," & oTabla.NroHistoriaClinicaAnterior & _
"           ," & oTabla.IdTipoNumeracionAnterior & _
"           )"
   oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    oConexionMDB.Close
ErrHistorico:
    Set oRsParametros = Nothing
    Set oRsMDB = Nothing
    Set oConexionMDB = Nothing
    
End Sub

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla HistoriasClinicas
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "HistoriasClinicasModificar"
           Set oParameter = .CreateParameter("@IdTipoNumeracionAnterior", adInteger, adParamInput, 0, IIf(oTabla.IdTipoNumeracionAnterior = 0, Null, oTabla.IdTipoNumeracionAnterior)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroHistoriaClinicaAnterior", adInteger, adParamInput, 0, IIf(oTabla.NroHistoriaClinicaAnterior = 0, Null, oTabla.NroHistoriaClinicaAnterior)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, IIf(oTabla.IdTipoNumeracion = 0, Null, oTabla.IdTipoNumeracion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, IIf(oTabla.NroHistoriaClinica = 0, Null, oTabla.NroHistoriaClinica)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaPasoAPasivo", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaPasoAPasivo = 0, Null, oTabla.FechaPasoAPasivo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoHistoria", adInteger, adParamInput, 0, IIf(oTabla.IdTipoHistoria = 0, Null, oTabla.IdTipoHistoria)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoHistoria", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoHistoria = 0, Null, oTabla.IdEstadoHistoria)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
 
   Modificar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla HistoriasClinicas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "HistoriasClinicasEliminar"
           Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, IIf(oTabla.NroHistoriaClinica = 0, Null, oTabla.NroHistoriaClinica)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
 
   Eliminar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla HistoriasClinicas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOHistoriaClinica) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String: Dim lnIdPaciente As Long: Dim lnIdTipoNumeracion As Long
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "HistoriasClinicasSeleccionarPorId"
       Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, oTabla.NroHistoriaClinica): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdTipoNumeracionAnterior = IIf(IsNull(oRecordset!IdTipoNumeracionAnterior), 0, oRecordset!IdTipoNumeracionAnterior)
       oTabla.NroHistoriaClinicaAnterior = IIf(IsNull(oRecordset!NroHistoriaClinicaAnterior), 0, oRecordset!NroHistoriaClinicaAnterior)
       oTabla.IdTipoNumeracion = IIf(IsNull(oRecordset!IdTipoNumeracion), 0, oRecordset!IdTipoNumeracion)
       oTabla.NroHistoriaClinica = IIf(IsNull(oRecordset!NroHistoriaClinica), 0, oRecordset!NroHistoriaClinica)
       oTabla.FechaCreacion = IIf(IsNull(oRecordset!FechaCreacion), 0, oRecordset!FechaCreacion)
       oTabla.FechaPasoAPasivo = IIf(IsNull(oRecordset!FechaPasoAPasivo), 0, oRecordset!FechaPasoAPasivo)
       oTabla.IdTipoHistoria = IIf(IsNull(oRecordset!IdTipoHistoria), 0, oRecordset!IdTipoHistoria)
       oTabla.IdEstadoHistoria = IIf(IsNull(oRecordset!IdEstadoHistoria), 0, oRecordset!IdEstadoHistoria)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.FechaUltimoMovimiento = IIf(IsNull(oRecordset!FultimoMovimiento), 0, oRecordset!FultimoMovimiento)
       oRecordset.Close
   Else
       SeleccionarPorId = False

       
       
   End If

   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        GenerarNroHistoria
'   Descripción:    Genera el Nro De Historia Clinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function GenerarNroHistoria(ByVal lTipoNumeracion As sghTipoNumeracionDeNroHistoria) As Long

On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lnNroHistoriaClinica As Long
   GenerarNroHistoria = 0
   With oCommand
       .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "HistoriasClinicasGenerarNroHistoria"
        Set oParameter = .CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, lTipoNumeracion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdHistoriaClinica", adInteger, adParamOutput, 0): .Parameters.Append oParameter
        .Execute
        lnNroHistoriaClinica = IIf(IsNull(.Parameters("@IdHistoriaClinica")), 0, .Parameters("@IdHistoriaClinica"))
   End With
   GenerarNroHistoria = lnNroHistoriaClinica
   If lTipoNumeracion = sghHistoriaDefinitivaAutomatica Then
           Dim oRsTmp1 As New Recordset
           Set oCommand = Nothing
           Set oParameter = Nothing
           With oCommand
                .CommandType = adCmdStoredProc
                 Set .ActiveConnection = mo_Conexion
                 .CommandText = "HistoriasClinicasSeleccionarPorId"
                 Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): .Parameters.Append oParameter
                 Set oRsTmp1 = .Execute
           End With
           If oRsTmp1.RecordCount > 0 Then
                oRsTmp1.Close
                Set oCommand = Nothing
                Set oParameter = Nothing
                With oCommand
                     .CommandType = adCmdStoredProc
                      Set .ActiveConnection = mo_Conexion
                      .CommandText = "HistoriasClinicasUltimoGenerado"
                      Set oRsTmp1 = .Execute
                End With
                lnNroHistoriaClinica = oRsTmp1.Fields!NroHistoriaClinica + 1
                '
                Set oCommand = Nothing
                Set oParameter = Nothing
                With oCommand
                     .CommandType = adCmdStoredProc
                      Set .ActiveConnection = mo_Conexion
                      .CommandText = "generadorNroHistoriaClinicaActualizaNroHistoria"
                      Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, lnNroHistoriaClinica): .Parameters.Append oParameter
                      .Execute
                End With
                GenerarNroHistoria = lnNroHistoriaClinica
           End If
           oRsTmp1.Close
   End If
   Set oCommand = Nothing
   ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function Filtrar(ByVal oTabla As DoPaciente, lcSinApellido As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    If oTabla.apellidopaterno <> "" Then
         If oTabla.apellidopaterno = lcSinApellido Then
            sWhere = sWhere + " Pacientes.ApellidoPaterno= '" + oTabla.apellidopaterno + "' and "
         Else
            sWhere = sWhere + " Pacientes.ApellidoPaterno like '" + oTabla.apellidopaterno + "%' and "
         End If
    End If
    If oTabla.apellidomaterno <> "" Then
         If oTabla.apellidomaterno = lcSinApellido Then
            sWhere = sWhere + " Pacientes.ApellidoMaterno='" + oTabla.apellidomaterno + "' and "
         Else
            sWhere = sWhere + " Pacientes.ApellidoMaterno like '" + oTabla.apellidomaterno + "%' and "
         End If
    End If
    If oTabla.PrimerNombre <> "" Then
         sWhere = sWhere + " Pacientes.PrimerNombre like '%" + oTabla.PrimerNombre + "%' and "
    End If
    If oTabla.NroHistoriaClinica <> 0 Then
         sWhere = sWhere + " Pacientes.NroHistoriaClinica ='" & oTabla.NroHistoriaClinica & "' and "
    End If
    If oTabla.IdTipoNumeracion <> 0 Then
         sWhere = sWhere + " Pacientes.IdTipoNumeracion ='" & oTabla.IdTipoNumeracion & "' and "
    End If
    If sWhere <> "" Then
         sSql = sSql + " where " + Left(sWhere, Len(sWhere) - 4)
    End If
    sSql = sSql + " order by Pacientes.ApellidoPAterno, Pacientes.ApellidoMaterno, Pacientes.PrimerNombre"
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HistoriasClinicasSegunFiltro"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set Filtrar = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
'mgaray201503
Function UltimoNroHistoria() As Long
On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lnNroHistoriaClinica As Long
    
    UltimoNroHistoria = 0
   
    Dim oRsTmp1 As New Recordset
    
    Set oCommand = Nothing
    Set oParameter = Nothing
    With oCommand
         .CommandType = adCmdStoredProc
          Set .ActiveConnection = mo_Conexion
          .CommandText = "HistoriasClinicasUltimoGenerado"
          Set oRsTmp1 = .Execute
    End With
    lnNroHistoriaClinica = oRsTmp1.Fields!NroHistoriaClinica
    oRsTmp1.Close
    
    UltimoNroHistoria = lnNroHistoriaClinica
    Set oCommand = Nothing
    ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
