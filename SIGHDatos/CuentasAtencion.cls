VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CuentasAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FacturacionCuentasAtencion
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
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
'   Descripción:    Inserta un registro a la tabla FacturacionCuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCuentaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionCuentasAtencionAgregar"
           Set oParameter = .CreateParameter("@TotalPorPagar", adCurrency, adParamInput, 0, IIf(oTabla.TotalPorPagar = 0, Null, oTabla.TotalPorPagar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, IIf(oTabla.IdEstado = 0, Null, oTabla.IdEstado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalPagado", adCurrency, adParamInput, 0, IIf(oTabla.TotalPagado = 0, Null, oTabla.TotalPagado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalAsegurado", adCurrency, adParamInput, 0, IIf(oTabla.TotalAsegurado = 0, Null, oTabla.TotalAsegurado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalExonerado", adCurrency, adParamInput, 0, IIf(oTabla.TotalExonerado = 0, Null, oTabla.TotalExonerado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraCierre", adChar, adParamInput, 5, IIf(oTabla.HoraCierre = "", Null, oTabla.HoraCierre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCierre", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCierre = 0, Null, oTabla.FechaCierre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraApertura", adChar, adParamInput, 5, IIf(oTabla.HoraApertura = "", Null, oTabla.HoraApertura)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaApertura", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaApertura = 0, Null, oTabla.FechaApertura)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdCuentaAtencion = .Parameters("@IdCuentaAtencion")
   End With
 
   Insertar = True
   ms_MensajeError = ""
   GrabaHistorico oTabla
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Sub GrabaHistorico(oTabla As DOCuentaAtencion)
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
    lcSql = "INSERT INTO FacturacionCuentasAtencion" & _
"           (idCuentaAtencion, IdPaciente,FechaApertura" & _
"           ,HoraApertura,FechaCierre" & _
"           ,HoraCierre,TotalExonerado" & _
"           ,TotalAsegurado,TotalPagado" & _
"           ,IdEstado,TotalPorPagar" & _
"           ,FechaCreacion)" & _
"     Values" & _
"           (" & _
"           " & oTabla.IdCuentaAtencion & "," & oTabla.IdPaciente & ",'" & oTabla.FechaApertura & _
"           ','" & oTabla.HoraApertura & "','" & oTabla.FechaCierre & _
"           ','" & oTabla.HoraCierre & "'," & oTabla.TotalExonerado & _
"           ," & oTabla.TotalAsegurado & "," & oTabla.TotalPagado & _
"           ," & oTabla.IdEstado & "," & oTabla.TotalPorPagar & _
"           ,'" & oTabla.FechaCreacion & _
"           ')"
    oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    oConexionMDB.Close
ErrHistorico:
    Set oRsParametros = Nothing
    Set oRsMDB = Nothing
    Set oConexionMDB = Nothing

End Sub

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla FacturacionCuentasAtencion
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCuentaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionCuentasAtencionModificar"
           Set oParameter = .CreateParameter("@TotalPorPagar", adCurrency, adParamInput, 0, IIf(oTabla.TotalPorPagar = 0, Null, oTabla.TotalPorPagar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, IIf(oTabla.IdEstado = 0, Null, oTabla.IdEstado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalPagado", adCurrency, adParamInput, 0, IIf(oTabla.TotalPagado = 0, Null, oTabla.TotalPagado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalAsegurado", adCurrency, adParamInput, 0, IIf(oTabla.TotalAsegurado = 0, Null, oTabla.TotalAsegurado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalExonerado", adCurrency, adParamInput, 0, IIf(oTabla.TotalExonerado = 0, Null, oTabla.TotalExonerado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraCierre", adChar, adParamInput, 5, IIf(oTabla.HoraCierre = "", Null, oTabla.HoraCierre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCierre", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCierre = 0, Null, oTabla.FechaCierre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraApertura", adChar, adParamInput, 5, IIf(oTabla.HoraApertura = "", Null, oTabla.HoraApertura)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaApertura", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaApertura = 0, Null, oTabla.FechaApertura)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
 
   Modificar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla FacturacionCuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCuentaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionCuentasAtencionEliminar"
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
 
   Eliminar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla FacturacionCuentasAtencion
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As sighcomun.DOCuentaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandTimeout = 150
     .CommandText = "FacturacionCuentasAtencionSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oTabla.IdCuentaAtencion): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.TotalPorPagar = IIf(IsNull(oRecordset!TotalPorPagar), 0, oRecordset!TotalPorPagar)
       oTabla.IdEstado = IIf(IsNull(oRecordset!IdEstado), 0, oRecordset!IdEstado)
       oTabla.TotalPagado = IIf(IsNull(oRecordset!TotalPagado), 0, oRecordset!TotalPagado)
       oTabla.TotalAsegurado = IIf(IsNull(oRecordset!TotalAsegurado), 0, oRecordset!TotalAsegurado)
       oTabla.TotalExonerado = IIf(IsNull(oRecordset!TotalExonerado), 0, oRecordset!TotalExonerado)
       oTabla.HoraCierre = IIf(IsNull(oRecordset!HoraCierre), "", oRecordset!HoraCierre)
       oTabla.FechaCierre = IIf(IsNull(oRecordset!FechaCierre), 0, oRecordset!FechaCierre)
       oTabla.HoraApertura = IIf(IsNull(oRecordset!HoraApertura), "", oRecordset!HoraApertura)
       oTabla.FechaApertura = IIf(IsNull(oRecordset!FechaApertura), 0, oRecordset!FechaApertura)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.FechaCreacion = IIf(IsNull(oRecordset!FechaCreacion), 0, oRecordset!FechaCreacion)
       'cada vez que se agregue un CAMPO añadirlo tambien en ATENCIONES.SeleccionarPorId
   Else
       SeleccionarPorId = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function SeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CuentasAtencionSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarIdPorIdAtencion(lIdAtencion As Long) As Long
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarIdPorIdAtencion = 0
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CuentasAtencionSeleccionarIdPorIdAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamOutput, 0): .Parameters.Append oParameter
       .Execute
       SeleccionarIdPorIdAtencion = .Parameters("@IdCuentaAtencion")
   End With
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarConsultaExterna() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarConsultaExterna = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CuentasAtencionSeleccionarConsultaExterna"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarConsultaExterna = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarTopicoEmergencia() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarTopicoEmergencia = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CuentasAtencionSeleccionarTopicoEmergencia"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarTopicoEmergencia = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
Function SeleccionarObservacionEmergencia() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarObservacionEmergencia = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CuentasAtencionSeleccionarObservacionEmergencia"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarObservacionEmergencia = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarHospitalizacion() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarHospitalizacion = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CuentasAtencionSeleccionarHospitalizacion"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarHospitalizacion = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'WCG Ini
Function ValidarCuentaAtencionFacturable(IdCuentaAtencion As Long) As Boolean

    ValidarCuentaAtencionFacturable = True

End Function

Function DatosPacientePorIdCuentaAtencion(lIdCuentaAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FacturacionCuentasAtencionDatosPacientePorIdCuentaAtencion"
        Set oParameter = .CreateParameter("@lIdCuentaAtencion", adInteger, adParamInput, 0, lIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set DatosPacientePorIdCuentaAtencion = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'WCG-05 Ini
Function ObtenerCuentasAtencionPorHistoriaClinica(NroHistoriaClinica As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FacturacionCuentasAtencionObtenerCuentasAtencionPorHistoriaClinica"
        Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, NroHistoriaClinica): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerCuentasAtencionPorHistoriaClinica = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ObtenerUltimaCuentaAtencionPorIdPaciente(lIdPaciente As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FacturacionCuentasAtencionObtenerUltimaCuentaAtencionPorIdPaciente"
        Set oParameter = .CreateParameter("@lIdPaciente", adInteger, adParamInput, 0, lIdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerUltimaCuentaAtencionPorIdPaciente = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function ObtenerCuentasAtencionPorHistoriaClinicaV2(NroHistoriaClinica As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sOrder As String
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FacturacionCuentasAtencionObtenerCuentasAtencionPorHistoriaClinicaV2"
        Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, NroHistoriaClinica): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerCuentasAtencionPorHistoriaClinicaV2 = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function





Function Abrir(ByVal oTabla As DOCuentaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRsTmp1 As New Recordset
   Abrir = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionCuentasAtencionAbrir"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
   Abrir = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function Cerrar(ByVal oTabla As DOCuentaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lcHora As String
   lcHora = Format(Time, SIGHENTIDADES.DevuelveHoraSoloFormato_HM)
   Cerrar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionCuentasAtencionCerrar"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcHoraCierre", adChar, adParamInput, 5, lcHora): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnDeudaPendiente", adCurrency, adParamInput, 0, oTabla.TotalPorPagar): .Parameters.Append oParameter
       .Execute
   End With
 
   Cerrar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function PendientePagoSeguro(ByVal oTabla As DOCuentaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lcHora As String
   lcHora = Format(Time, SIGHENTIDADES.DevuelveHoraSoloFormato_HM)
   PendientePagoSeguro = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionCuentasAtencionPendientePagoSeguro"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcHoraCierre", adChar, adParamInput, 5, lcHora): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnDeudaPendiente", adCurrency, adParamInput, 0, oTabla.TotalPorPagar): .Parameters.Append oParameter
       .Execute
   End With
 
   PendientePagoSeguro = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function Pagada(ByVal oTabla As DOCuentaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lcHora As String
   lcHora = Format(Time, SIGHENTIDADES.DevuelveHoraSoloFormato_HM)
   Pagada = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionCuentasAtencionPagada"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcHoraCierre", adChar, adParamInput, 5, lcHora): .Parameters.Append oParameter
       .Execute
   End With
 
   Pagada = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function Anulada(ByVal oTabla As DOCuentaAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lcHora As String
   lcHora = Format(Time, SIGHENTIDADES.DevuelveHoraSoloFormato_HM)
   Anulada = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionCuentasAtencionAnulada"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcHoraCierre", adChar, adParamInput, 5, lcHora): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnDeudaPendiente", adCurrency, adParamInput, 0, oTabla.TotalPorPagar): .Parameters.Append oParameter
       .Execute
   End With
 
   Anulada = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function AltaConDeudaYGarante(ByVal oTabla As DOCuentaAtencion) As Boolean
  On Error GoTo ManejadorDeError
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim lcHora As String
  lcHora = Format(Time, SIGHENTIDADES.DevuelveHoraSoloFormato_HM)
  AltaConDeudaYGarante = False
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = mo_Conexion
    .CommandText = "FacturacionCuentasAtencionAltaConDeudaYGarante"    '"FacturacionAltaConDeudaYGarante"
    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcHoraCierre", adChar, adParamInput, 5, lcHora): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnDeudaPendiente", adCurrency, adParamInput, 0, oTabla.TotalPorPagar): .Parameters.Append oParameter
    .Execute
  End With
 
  AltaConDeudaYGarante = True
  ms_MensajeError = ""
  Exit Function
  
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description
  Exit Function
End Function
