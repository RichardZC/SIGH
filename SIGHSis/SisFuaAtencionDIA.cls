VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SisFuaAtencionDIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para el Mantenimiento de la tabla SisAtencionDIA
'        Programado por: Barrantes D
'        Fecha: Enero 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ms_MensajeError As String
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla SisFuaAtencionDIA
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoSisFuaAtencionDIA) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaAtencionDIAAgregar"
   Set oParameter = .CreateParameter("@id", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.idCuentaAtencion = 0, Null, oTabla.idCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DxNumero", adInteger, adParamInput, 0, IIf(oTabla.DxNumero = 0, Null, oTabla.DxNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DxTipoIE", adVarChar, adParamInput, 1, IIf(oTabla.DxTipoIE = "", Null, oTabla.DxTipoIE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DxCodigo", adVarChar, adParamInput, 5, IIf(oTabla.DxCodigo = "", Null, oTabla.DxCodigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DxTipoDPR", adVarChar, adParamInput, 1, IIf(oTabla.DxTipoDPR = "", Null, oTabla.DxTipoDPR))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabDniUsuarioRegistra", adVarChar, adParamInput, 10, IIf(oTabla.CabDniUsuarioRegistra = "", Null, Left(oTabla.CabDniUsuarioRegistra, 10)))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabFechaFuaPrimeraVez", adVarChar, adParamInput, 10, IIf(oTabla.CabFechaFuaPrimeraVez = "", Null, oTabla.CabFechaFuaPrimeraVez))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabEstado", adChar, adParamInput, 1, IIf(oTabla.CabEstado = "", Null, oTabla.CabEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabNroEnvioAlSIS", adVarChar, adParamInput, 2, IIf(oTabla.CabNroEnvioAlSIS = "", Null, oTabla.CabNroEnvioAlSIS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabCodigoPuntoDigitacion", adInteger, adParamInput, 0, IIf(oTabla.CabCodigoPuntoDigitacion = 0, Null, oTabla.CabCodigoPuntoDigitacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabCodigoUDR", adVarChar, adParamInput, 3, IIf(oTabla.CabCodigoUDR = "", Null, oTabla.CabCodigoUDR))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaDisa", adVarChar, adParamInput, 3, IIf(oTabla.FuaDisa = "", Null, oTabla.FuaDisa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaLote", adVarChar, adParamInput, 2, IIf(oTabla.FuaLote = "", Null, oTabla.FuaLote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaNumero", adVarChar, adParamInput, 16, IIf(oTabla.FuaNumero = "", Null, oTabla.FuaNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabOrigenDelRegistro", adVarChar, adParamInput, 9, IIf(oTabla.CabOrigenDelRegistro = "", Null, oTabla.CabOrigenDelRegistro))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabVersionAplicativo", adVarChar, adParamInput, 9, IIf(oTabla.CabVersionAplicativo = "", Null, oTabla.CabVersionAplicativo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabIdentificacionPaquete", adInteger, adParamInput, 0, oTabla.CabIdentificacionPaquete)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.id = .Parameters("@id")
End With
 
Insertar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
'   Resume
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla SisFuaAtencionDIA
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoSisFuaAtencionDIA) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaAtencionDIAModificar"
   Set oParameter = .CreateParameter("@id", adInteger, adParamInput, 0, IIf(oTabla.id = 0, Null, oTabla.id))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.idCuentaAtencion = 0, Null, oTabla.idCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DxNumero", adInteger, adParamInput, 0, IIf(oTabla.DxNumero = 0, Null, oTabla.DxNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DxTipoIE", adVarChar, adParamInput, 1, IIf(oTabla.DxTipoIE = "", Null, oTabla.DxTipoIE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DxCodigo", adVarChar, adParamInput, 5, IIf(oTabla.DxCodigo = "", Null, oTabla.DxCodigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DxTipoDPR", adVarChar, adParamInput, 1, IIf(oTabla.DxTipoDPR = "", Null, oTabla.DxTipoDPR))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabDniUsuarioRegistra", adVarChar, adParamInput, 10, IIf(oTabla.CabDniUsuarioRegistra = "", Null, Left(oTabla.CabDniUsuarioRegistra, 10)))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabFechaFuaPrimeraVez", adVarChar, adParamInput, 10, IIf(oTabla.CabFechaFuaPrimeraVez = "", Null, oTabla.CabFechaFuaPrimeraVez))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabEstado", adChar, adParamInput, 1, IIf(oTabla.CabEstado = "", Null, oTabla.CabEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabNroEnvioAlSIS", adVarChar, adParamInput, 2, IIf(oTabla.CabNroEnvioAlSIS = "", Null, oTabla.CabNroEnvioAlSIS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabCodigoPuntoDigitacion", adInteger, adParamInput, 0, IIf(oTabla.CabCodigoPuntoDigitacion = 0, Null, oTabla.CabCodigoPuntoDigitacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabCodigoUDR", adVarChar, adParamInput, 3, IIf(oTabla.CabCodigoUDR = "", Null, oTabla.CabCodigoUDR))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaDisa", adVarChar, adParamInput, 3, IIf(oTabla.FuaDisa = "", Null, oTabla.FuaDisa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaLote", adVarChar, adParamInput, 2, IIf(oTabla.FuaLote = "", Null, oTabla.FuaLote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaNumero", adVarChar, adParamInput, 16, IIf(oTabla.FuaNumero = "", Null, oTabla.FuaNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabOrigenDelRegistro", adVarChar, adParamInput, 9, IIf(oTabla.CabOrigenDelRegistro = "", Null, oTabla.CabOrigenDelRegistro))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabVersionAplicativo", adVarChar, adParamInput, 9, IIf(oTabla.CabVersionAplicativo = "", Null, oTabla.CabVersionAplicativo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CabIdentificacionPaquete", adInteger, adParamInput, 0, oTabla.CabIdentificacionPaquete)
   .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla SisFuaAtencionDIA
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoSisFuaAtencionDIA) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaAtencionDIAEliminar"
   Set oParameter = .CreateParameter("@id", adInteger, adParamInput, 0, oTabla.id)
   .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla SisFuaAtencionDIA
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoSisFuaAtencionDIA) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "SisFuaAtencionDIASeleccionarPorId"
     Set oParameter = .CreateParameter("@id", adInteger, adParamInput, 0, oTabla.id)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.id = IIf(IsNull(oRecordset!id), 0, oRecordset!id)
       oTabla.idCuentaAtencion = IIf(IsNull(oRecordset!idCuentaAtencion), 0, oRecordset!idCuentaAtencion)
       oTabla.DxNumero = IIf(IsNull(oRecordset!DxNumero), 0, oRecordset!DxNumero)
       oTabla.DxTipoIE = IIf(IsNull(oRecordset!DxTipoIE), "", oRecordset!DxTipoIE)
       oTabla.DxCodigo = IIf(IsNull(oRecordset!DxCodigo), "", oRecordset!DxCodigo)
       oTabla.DxTipoDPR = IIf(IsNull(oRecordset!DxTipoDPR), "", oRecordset!DxTipoDPR)
       oTabla.CabDniUsuarioRegistra = IIf(IsNull(oRecordset!CabDniUsuarioRegistra), "", oRecordset!CabDniUsuarioRegistra)
       oTabla.CabFechaFuaPrimeraVez = IIf(IsNull(oRecordset!CabFechaFuaPrimeraVez), "", oRecordset!CabFechaFuaPrimeraVez)
       oTabla.CabEstado = IIf(IsNull(oRecordset!CabEstado), "", oRecordset!CabEstado)
       oTabla.CabNroEnvioAlSIS = IIf(IsNull(oRecordset!CabNroEnvioAlSIS), "", oRecordset!CabNroEnvioAlSIS)
       oTabla.CabCodigoPuntoDigitacion = IIf(IsNull(oRecordset!CabCodigoPuntoDigitacion), 0, oRecordset!CabCodigoPuntoDigitacion)
       oTabla.CabCodigoUDR = IIf(IsNull(oRecordset!CabCodigoUDR), "", oRecordset!CabCodigoUDR)
       oTabla.FuaDisa = IIf(IsNull(oRecordset!FuaDisa), "", oRecordset!FuaDisa)
       oTabla.FuaLote = IIf(IsNull(oRecordset!FuaLote), "", oRecordset!FuaLote)
       oTabla.FuaNumero = IIf(IsNull(oRecordset!FuaNumero), "", oRecordset!FuaNumero)
       oTabla.CabOrigenDelRegistro = IIf(IsNull(oRecordset!CabOrigenDelRegistro), "", oRecordset!CabOrigenDelRegistro)
       oTabla.CabVersionAplicativo = IIf(IsNull(oRecordset!CabVersionAplicativo), "", oRecordset!CabVersionAplicativo)
       oTabla.CabIdentificacionPaquete = IIf(IsNull(oRecordset!CabIdentificacionPaquete), 0, oRecordset!CabIdentificacionPaquete)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoSisFuaAtencionDIA)
   oTabla.id = 0
   oTabla.idCuentaAtencion = 0
   oTabla.DxNumero = 0
   oTabla.DxTipoIE = ""
   oTabla.DxCodigo = ""
   oTabla.DxTipoDPR = ""
   oTabla.CabDniUsuarioRegistra = ""
   oTabla.CabFechaFuaPrimeraVez = ""
   oTabla.CabEstado = ""
   oTabla.CabNroEnvioAlSIS = ""
   oTabla.CabCodigoPuntoDigitacion = 0
   oTabla.CabCodigoUDR = ""
   oTabla.FuaDisa = ""
   oTabla.FuaLote = ""
   oTabla.FuaNumero = ""
   oTabla.CabOrigenDelRegistro = ""
   oTabla.CabVersionAplicativo = ""
   oTabla.CabIdentificacionPaquete = 0
End Sub

Function SisFuaAtencionDIAeliminarPorCuenta(lnIdCuentaAtencion As Long, lnIdUsuarioAuditoria As Long) As Boolean

On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
SisFuaAtencionDIAeliminarPorCuenta = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaAtencionDIAEliminarIdCuentaAtencion"
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, lnIdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
SisFuaAtencionDIAeliminarPorCuenta = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function





