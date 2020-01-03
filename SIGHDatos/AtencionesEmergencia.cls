VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionesEmergencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla AtencionesEmergencia
'        Programado por: Castro W
'        Fecha: Agosto 2006
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
'   Descripción:    Inserta un registro a la tabla AtencionesEmergencia
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOAtencionEmergencia) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesEmergenciaAgregar"
           Set oParameter = .CreateParameter("@IdTipoAgenteAGAN", adInteger, adParamInput, 0, IIf(oTabla.IdTipoAgenteAGAN = 0, Null, oTabla.IdTipoAgenteAGAN)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGrupoOcupacionalALAB", adInteger, adParamInput, 0, IIf(oTabla.IdGrupoOcupacionalALAB = 0, Null, oTabla.IdGrupoOcupacionalALAB)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPosicionLesionadoALAB", adInteger, adParamInput, 0, IIf(oTabla.IdPosicionLesionadoALAB = 0, Null, oTabla.IdPosicionLesionadoALAB)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUbicacionLesionado", adInteger, adParamInput, 0, IIf(oTabla.IdUbicacionLesionado = 0, Null, oTabla.IdUbicacionLesionado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoTransporte", adInteger, adParamInput, 0, IIf(oTabla.IdTipoTransporte = 0, Null, oTabla.IdTipoTransporte)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoVehiculo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoVehiculo = 0, Null, oTabla.IdTipoVehiculo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdClaseAccidente", adInteger, adParamInput, 0, IIf(oTabla.IdClaseAccidente = 0, Null, oTabla.IdClaseAccidente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdRelacionAgresorVictima", adInteger, adParamInput, 0, IIf(oTabla.IdRelacionAgresorVictima = 0, Null, oTabla.IdRelacionAgresorVictima)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdSeguridad", adInteger, adParamInput, 0, IIf(oTabla.IdSeguridad = 0, Null, oTabla.IdSeguridad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoEvento", adInteger, adParamInput, 0, IIf(oTabla.IdTipoEvento = 0, Null, oTabla.IdTipoEvento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdLugarEvento", adInteger, adParamInput, 0, IIf(oTabla.IdLugarEvento = 0, Null, oTabla.IdLugarEvento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCausaExternaMorbilidad", adInteger, adParamInput, 0, IIf(oTabla.IdCausaExternaMorbilidad = 0, Null, oTabla.IdCausaExternaMorbilidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionEmergencia", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@comoLlego", adInteger, adParamInput, 0, oTabla.comoLlego): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@tipoAtencion", adInteger, adParamInput, 0, oTabla.tipoAtencion): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idEstadoLlegada", adInteger, adParamInput, 0, oTabla.idEstadoLlegada): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdAtencionEmergencia = .Parameters("@IdAtencionEmergencia")
   End With
 
   Insertar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla AtencionesEmergencia
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtencionEmergencia) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesEmergenciaModificar"
           Set oParameter = .CreateParameter("@IdTipoAgenteAGAN", adInteger, adParamInput, 0, IIf(oTabla.IdTipoAgenteAGAN = 0, Null, oTabla.IdTipoAgenteAGAN)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGrupoOcupacionalALAB", adInteger, adParamInput, 0, IIf(oTabla.IdGrupoOcupacionalALAB = 0, Null, oTabla.IdGrupoOcupacionalALAB)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPosicionLesionadoALAB", adInteger, adParamInput, 0, IIf(oTabla.IdPosicionLesionadoALAB = 0, Null, oTabla.IdPosicionLesionadoALAB)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUbicacionLesionado", adInteger, adParamInput, 0, IIf(oTabla.IdUbicacionLesionado = 0, Null, oTabla.IdUbicacionLesionado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoTransporte", adInteger, adParamInput, 0, IIf(oTabla.IdTipoTransporte = 0, Null, oTabla.IdTipoTransporte)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoVehiculo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoVehiculo = 0, Null, oTabla.IdTipoVehiculo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdClaseAccidente", adInteger, adParamInput, 0, IIf(oTabla.IdClaseAccidente = 0, Null, oTabla.IdClaseAccidente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdRelacionAgresorVictima", adInteger, adParamInput, 0, IIf(oTabla.IdRelacionAgresorVictima = 0, Null, oTabla.IdRelacionAgresorVictima)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdSeguridad", adInteger, adParamInput, 0, IIf(oTabla.IdSeguridad = 0, Null, oTabla.IdSeguridad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoEvento", adInteger, adParamInput, 0, IIf(oTabla.IdTipoEvento = 0, Null, oTabla.IdTipoEvento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdLugarEvento", adInteger, adParamInput, 0, IIf(oTabla.IdLugarEvento = 0, Null, oTabla.IdLugarEvento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCausaExternaMorbilidad", adInteger, adParamInput, 0, IIf(oTabla.IdCausaExternaMorbilidad = 0, Null, oTabla.IdCausaExternaMorbilidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionEmergencia", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionEmergencia = 0, Null, oTabla.IdAtencionEmergencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@comoLlego", adInteger, adParamInput, 0, oTabla.comoLlego): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@tipoAtencion", adInteger, adParamInput, 0, oTabla.tipoAtencion): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idEstadoLlegada", adInteger, adParamInput, 0, oTabla.idEstadoLlegada): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
 
   Modificar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   MsgBox Err.Description

   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla AtencionesEmergencia
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtencionEmergencia) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesEmergenciaEliminar"
        Set oParameter = .CreateParameter("@IdAtencionEmergencia", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionEmergencia = 0, Null, oTabla.IdAtencionEmergencia)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla AtencionesEmergencia
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtencionEmergencia) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtencionesEmergenciaSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdAtencionEmergencia", adInteger, adParamInput, 0, oTabla.IdAtencionEmergencia): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdTipoAgenteAGAN = IIf(IsNull(oRecordset!IdTipoAgenteAGAN), 0, oRecordset!IdTipoAgenteAGAN)
       oTabla.IdGrupoOcupacionalALAB = IIf(IsNull(oRecordset!IdGrupoOcupacionalALAB), 0, oRecordset!IdGrupoOcupacionalALAB)
       oTabla.IdPosicionLesionadoALAB = IIf(IsNull(oRecordset!IdPosicionLesionadoALAB), 0, oRecordset!IdPosicionLesionadoALAB)
       oTabla.IdUbicacionLesionado = IIf(IsNull(oRecordset!IdUbicacionLesionado), 0, oRecordset!IdUbicacionLesionado)
       oTabla.IdTipoTransporte = IIf(IsNull(oRecordset!IdTipoTransporte), 0, oRecordset!IdTipoTransporte)
       oTabla.IdTipoVehiculo = IIf(IsNull(oRecordset!IdTipoVehiculo), 0, oRecordset!IdTipoVehiculo)
       oTabla.IdClaseAccidente = IIf(IsNull(oRecordset!IdClaseAccidente), 0, oRecordset!IdClaseAccidente)
       oTabla.IdRelacionAgresorVictima = IIf(IsNull(oRecordset!IdRelacionAgresorVictima), 0, oRecordset!IdRelacionAgresorVictima)
       oTabla.IdSeguridad = IIf(IsNull(oRecordset!IdSeguridad), 0, oRecordset!IdSeguridad)
       oTabla.IdTipoEvento = IIf(IsNull(oRecordset!IdTipoEvento), 0, oRecordset!IdTipoEvento)
       oTabla.IdLugarEvento = IIf(IsNull(oRecordset!IdLugarEvento), 0, oRecordset!IdLugarEvento)
       oTabla.IdCausaExternaMorbilidad = IIf(IsNull(oRecordset!IdCausaExternaMorbilidad), 0, oRecordset!IdCausaExternaMorbilidad)
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       oTabla.IdAtencionEmergencia = IIf(IsNull(oRecordset!IdAtencionEmergencia), 0, oRecordset!IdAtencionEmergencia)
       oTabla.comoLlego = IIf(IsNull(oRecordset!comoLlego), 0, oRecordset!comoLlego)
       oTabla.tipoAtencion = IIf(IsNull(oRecordset!tipoAtencion), 0, oRecordset!tipoAtencion)
       oTabla.idEstadoLlegada = IIf(IsNull(oRecordset!idEstadoLlegada), 0, oRecordset!idEstadoLlegada)
   Else
       SeleccionarPorId = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function CausaExternaMorbilidadSeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set CausaExternaMorbilidadSeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EmergenciaCausaExternaMorbilidadSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set CausaExternaMorbilidadSeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
    Set oCommand = Nothing
   Set oRecordset = Nothing

Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function
Function ClaseAccidenteSeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set ClaseAccidenteSeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EmergenciaClaseAccidenteSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set ClaseAccidenteSeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function GrupoOcupacionalALABSeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set GrupoOcupacionalALABSeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EmergenciaGrupoOcupacionalALABSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set GrupoOcupacionalALABSeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function LugarEventoSeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set LugarEventoSeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EmergenciaLugarEventoSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set LugarEventoSeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function PosicionLesionadoALABSeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set PosicionLesionadoALABSeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EmergenciaPosicionLesionadoALABSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set PosicionLesionadoALABSeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function RelacionAgresorVictimaSeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set RelacionAgresorVictimaSeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EmergenciaRelacionAgresorVictimaSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set RelacionAgresorVictimaSeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function SeguridadSeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeguridadSeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EmergenciaSeguridadSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeguridadSeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function TipoAgenteAGANSeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set TipoAgenteAGANSeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EmergenciaTipoAgenteAGANSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set TipoAgenteAGANSeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function TipoEventoSeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set TipoEventoSeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EmergenciaTipoEventoSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set TipoEventoSeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function TipoTransporteSeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set TipoTransporteSeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EmergenciaTipoTransporteSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set TipoTransporteSeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function TipoVehiculoSeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set TipoVehiculoSeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EmergenciaTipoVehiculoSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set TipoVehiculoSeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function UbicacionLesionadoSeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set UbicacionLesionadoSeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EmergenciaUbicacionLesionadoSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set UbicacionLesionadoSeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function SeleccionarIdPorIdAtencion(lIdAtencion As Long) As Long
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarIdPorIdAtencion = 0
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesEmergenciaSeleccionarIdPorIdAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtencionEmergencia", adInteger, adParamOutput, 0): .Parameters.Append oParameter
       .Execute
       'Actualizado 16092014
       SeleccionarIdPorIdAtencion = IIf(IsNull(.Parameters("@IdAtencionEmergencia")), 0, .Parameters("@IdAtencionEmergencia"))
       
   End With
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function EliminarAtencionEmergenciaPorIdAtencion(lIdAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesEmergenciaEliminarXidAtencion"
        Set oParameter = .CreateParameter("@lIdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        .Execute
   End With
   Set oRecordset = Nothing
   Set oCommand = Nothing
   EliminarAtencionEmergenciaPorIdAtencion = True
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

