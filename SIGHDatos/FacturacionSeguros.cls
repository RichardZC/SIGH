VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturacionSeguros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FacturacionSeguros
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
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
'   Descripción:    Inserta un registro a la tabla FacturacionSeguros
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOFacturacionSeguros) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionSegurosAgregar"
           Set oParameter = .CreateParameter("@NroPlaca", adChar, adParamInput, 8, IIf(oTabla.NroPlaca = "", Null, oTabla.NroPlaca)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoAutorizacion", adVarChar, adParamInput, 10, IIf(oTabla.CodigoAutorizacion = "", Null, oTabla.CodigoAutorizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFacturacionSeguro", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalAsegurado", adCurrency, adParamInput, 0, IIf(oTabla.TotalAsegurado = 0, Null, oTabla.TotalAsegurado)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdFacturacionSeguro = .Parameters("@IdFacturacionSeguro")
   End With
 
   Insertar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla FacturacionSeguros
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOFacturacionSeguros) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionSegurosModificar"
           Set oParameter = .CreateParameter("@NroPlaca", adChar, adParamInput, 8, IIf(oTabla.NroPlaca = "", Null, oTabla.NroPlaca)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoAutorizacion", adVarChar, adParamInput, 10, IIf(oTabla.CodigoAutorizacion = "", Null, oTabla.CodigoAutorizacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFacturacionSeguro", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionSeguro = 0, Null, oTabla.IdFacturacionSeguro)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TotalAsegurado", adCurrency, adParamInput, 0, IIf(oTabla.TotalAsegurado = 0, Null, oTabla.TotalAsegurado)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla FacturacionSeguros
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOFacturacionSeguros) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FacturacionSegurosEliminar"
           Set oParameter = .CreateParameter("@IdFacturacionSeguro", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionSeguro = 0, Null, oTabla.IdFacturacionSeguro)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla FacturacionSeguros
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOFacturacionSeguros) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FacturacionSegurosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdFacturacionSeguro", adInteger, adParamInput, 0, oTabla.IdFacturacionSeguro): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.NroPlaca = IIf(IsNull(oRecordset!NroPlaca), "", oRecordset!NroPlaca)
       oTabla.CodigoAutorizacion = IIf(IsNull(oRecordset!CodigoAutorizacion), "", oRecordset!CodigoAutorizacion)
       oTabla.idTipoFinanciamiento = IIf(IsNull(oRecordset!idTipoFinanciamiento), 0, oRecordset!idTipoFinanciamiento)
       oTabla.idFuenteFinanciamiento = IIf(IsNull(oRecordset!idFuenteFinanciamiento), 0, oRecordset!idFuenteFinanciamiento)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.IdFacturacionSeguro = IIf(IsNull(oRecordset!IdFacturacionSeguro), 0, oRecordset!IdFacturacionSeguro)
       oTabla.TotalAsegurado = IIf(IsNull(oRecordset!TotalAsegurado), 0, oRecordset!TotalAsegurado)
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

Function SeleccionarPorCuentaAtencionParaEstadoCuenta(IdCuentaAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oRecordset As New Recordset
   
    Set SeleccionarPorCuentaAtencionParaEstadoCuenta = Nothing
   
    sSql = " select a.idfacturacionseguro as idfacturacionseguro, a.IdTipoFinanciamiento, b.descripcion as DescTipoFinanciamiento, a.IdFuenteFinanciamiento, a.CodigoAutorizacion as CodigoAutorizacion,"
    sSql = sSql & " a.nroPlaca, a.totalAsegurado, a.IdCuentaAtencion "
    sSql = sSql & " From facturacionseguros a,tiposFinanciamiento b"
    sSql = sSql & " where b.idTipoFinanciamiento = a.idTipoFinanciamiento"
    sSql = sSql & " and idcuentaatencion = " & IdCuentaAtencion
   
'    With oCommand
'        .CommandType = adCmdText
'        Set .ActiveConnection = mo_Conexion
'        .CommandText = sSQL
'        Set oRecordset = .Execute
'    End With
 
    oRecordset.Open sSql, mo_Conexion, adOpenDynamic, adLockOptimistic
    
    Set oRecordset.ActiveConnection = Nothing
 
    Set SeleccionarPorCuentaAtencionParaEstadoCuenta = oRecordset
    
    
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function


