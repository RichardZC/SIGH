VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SisFuaAtencionINS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de c�digo autogenerado para la clase: SisFuaAtencionINS
'        Autor: Daniel Enrique Barrantes
'        Fecha: 17/07/2012 12:23:28 p.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

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
'   Funci�n:        Insertar
'   Descripci�n:    Inserta un registro a la tabla SisFuaAtencionINS
'   Par�metros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOSisFuaAtencionINS) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaAtencionINSAgregar"
   Set oParameter = .CreateParameter("@id", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTablaDx", adInteger, adParamInput, 0, IIf(oTabla.idTablaDx = 0, Null, oTabla.idTablaDx))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DxNumero", adInteger, adParamInput, 0, IIf(oTabla.DxNumero = 0, Null, oTabla.DxNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 5, IIf(oTabla.codigo = "", Null, oTabla.codigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadPrescrita", adInteger, adParamInput, 0, IIf(oTabla.CantidadPrescrita = 0, Null, oTabla.CantidadPrescrita))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadEntregada", adInteger, adParamInput, 0, oTabla.CantidadEntregada)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, oTabla.PrecioUnitario)
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
   Set oParameter = .CreateParameter("@FuaNumero", adVarChar, adParamInput, 8, IIf(oTabla.FuaNumero = "", Null, oTabla.FuaNumero))
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
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Funci�n:        Modificar
'   Descripci�n:    modifica un registro a la tabla SisFuaAtencionINS
'   Par�metros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOSisFuaAtencionINS) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaAtencionINSModificar"
   Set oParameter = .CreateParameter("@id", adInteger, adParamInput, 0, IIf(oTabla.id = 0, Null, oTabla.id))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTablaDx", adInteger, adParamInput, 0, IIf(oTabla.idTablaDx = 0, Null, oTabla.idTablaDx))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DxNumero", adInteger, adParamInput, 0, IIf(oTabla.DxNumero = 0, Null, oTabla.DxNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 5, IIf(oTabla.codigo = "", Null, oTabla.codigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadPrescrita", adInteger, adParamInput, 0, IIf(oTabla.CantidadPrescrita = 0, Null, oTabla.CantidadPrescrita))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadEntregada", adInteger, adParamInput, 0, oTabla.CantidadEntregada)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, oTabla.PrecioUnitario)
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
   Set oParameter = .CreateParameter("@FuaNumero", adVarChar, adParamInput, 8, IIf(oTabla.FuaNumero = "", Null, oTabla.FuaNumero))
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
'   Funci�n:        Eliminar
'   Descripci�n:    elimia o borra un registro a la tabla SisFuaAtencionINS
'   Par�metros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOSisFuaAtencionINS) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaAtencionINSEliminar"
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
'   Funci�n:        SeleccionarPorId
'   Descripci�n:    Selecciona un unico registro de la tabla SisFuaAtencionINS
'   Par�metros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOSisFuaAtencionINS) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "SisFuaAtencionINSSeleccionarPorId"
     Set oParameter = .CreateParameter("@id", adInteger, adParamInput, 0, oTabla.id)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.id = IIf(IsNull(oRecordset!id), 0, oRecordset!id)
       oTabla.idTablaDx = IIf(IsNull(oRecordset!idTablaDx), 0, oRecordset!idTablaDx)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.DxNumero = IIf(IsNull(oRecordset!DxNumero), 0, oRecordset!DxNumero)
       oTabla.codigo = IIf(IsNull(oRecordset!codigo), "", oRecordset!codigo)
       oTabla.CantidadPrescrita = IIf(IsNull(oRecordset!CantidadPrescrita), 0, oRecordset!CantidadPrescrita)
       oTabla.CantidadEntregada = IIf(IsNull(oRecordset!CantidadEntregada), 0, oRecordset!CantidadEntregada)
       oTabla.PrecioUnitario = oRecordset!PrecioUnitario
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
Sub SetDefaults(ByVal oTabla As DOSisFuaAtencionINS)
   oTabla.id = 0
   oTabla.idTablaDx = 0
   oTabla.IdCuentaAtencion = 0
   oTabla.DxNumero = 0
   oTabla.codigo = ""
   oTabla.CantidadPrescrita = 0
   oTabla.CantidadEntregada = 0
   oTabla.PrecioUnitario = 0
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

Function SisFuaAtencionINSeliminarPorCuenta(lnIdCuentaAtencion As Long) As Boolean
    On Error GoTo ErrSisFuaElimCta
    SisFuaAtencionINSeliminarPorCuenta = False
    Dim oRsTmp1 As New Recordset
    oRsTmp1.Open "delete from SisFuaAtencionINS where idCuentaAtencion=" & lnIdCuentaAtencion, mo_Conexion, adOpenKeyset, adLockOptimistic
    SisFuaAtencionINSeliminarPorCuenta = True
    Exit Function
ErrSisFuaElimCta:
    ms_MensajeError = Err.Number & " " + Err.Description
End Function

Function SisFuaAtencionINSanularPorCuenta(lnIdCuentaAtencion As Long) As Boolean
    On Error GoTo ErrSisFuaElimCta
    SisFuaAtencionINSanularPorCuenta = False
    Dim oRsTmp1 As New Recordset
    oRsTmp1.Open "update SisFuaAtencionINS set CabEstado='1' where idCuentaAtencion=" & lnIdCuentaAtencion, mo_Conexion, adOpenKeyset, adLockOptimistic
    SisFuaAtencionINSanularPorCuenta = True
    Exit Function
ErrSisFuaElimCta:
    ms_MensajeError = Err.Number & " " + Err.Description
End Function


'------------------------------------------------------------------------------------
'        Fin de c�digo autogenerado para la clase: SisFuaAtencionINS
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de c�digo manual para la clase: SisFuaAtencionINS
'        Autor:   DAniel Barrantes B.
'        Fecha: 17/07/2012 12:23:28 p.m.
'------------------------------------------------------------------------------------
 

