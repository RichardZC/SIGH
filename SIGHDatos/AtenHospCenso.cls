VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtenHospCenso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: AtenHospCenso
'        Autor: Franklin Cachay Velasquez
'        Fecha: 27/02/2015 03:24:07
'        Auspiciado por: Ministerio de Salud
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
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla AtenHospCenso
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoAtencionHospCenso) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenHospCensoAgregar"
   Set oParameter = .CreateParameter("@IdRangoCensoHosp", adInteger, adParamInput, 0, IIf(oTabla.IdRangoCensoHosp = 0, Null, oTabla.IdRangoCensoHosp))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RangoInicial", adCurrency, adParamInput, 0, oTabla.RangoInicial)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RangoFinal", adCurrency, adParamInput, 0, oTabla.RangoFinal)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RGBRojo", adInteger, adParamInput, 0, IIf(oTabla.RGBRojo = 0, Null, oTabla.RGBRojo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RGBVerde", adInteger, adParamInput, 0, IIf(oTabla.RGBVerde = 0, Null, oTabla.RGBVerde))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RGBAzul", adInteger, adParamInput, 0, IIf(oTabla.RGBAzul = 0, Null, oTabla.RGBAzul))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
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
'   Descripción:    modifica un registro a la tabla AtenHospCenso
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoAtencionHospCenso) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenHospCensoModificar"
   Set oParameter = .CreateParameter("@IdRangoCensoHosp", adInteger, adParamInput, 0, IIf(oTabla.IdRangoCensoHosp = 0, Null, oTabla.IdRangoCensoHosp))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RangoInicial", adCurrency, adParamInput, 0, oTabla.RangoInicial)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RangoFinal", adCurrency, adParamInput, 0, oTabla.RangoFinal)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RGBRojo", adInteger, adParamInput, 0, IIf(oTabla.RGBRojo = 0, Null, oTabla.RGBRojo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RGBVerde", adInteger, adParamInput, 0, IIf(oTabla.RGBVerde = 0, Null, oTabla.RGBVerde))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RGBAzul", adInteger, adParamInput, 0, IIf(oTabla.RGBAzul = 0, Null, oTabla.RGBAzul))
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
'   Descripción:    elimia o borra un registro a la tabla AtenHospCenso
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoAtencionHospCenso) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenHospCensoEliminar"
   Set oParameter = .CreateParameter("@IdRangoCensoHosp", adInteger, adParamInput, 0, oTabla.IdRangoCensoHosp)
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
'   Descripción:    Selecciona un unico registro de la tabla AtenHospCenso
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoAtencionHospCenso) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtenHospCensoSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdRangoCensoHosp", adInteger, adParamInput, 0, oTabla.IdRangoCensoHosp)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdRangoCensoHosp = IIf(IsNull(oRecordset!IdRangoCensoHosp), 0, oRecordset!IdRangoCensoHosp)
       oTabla.RangoInicial = oRecordset!RangoInicial
       oTabla.RangoFinal = oRecordset!RangoFinal
       oTabla.RGBRojo = IIf(IsNull(oRecordset!RGBRojo), 0, oRecordset!RGBRojo)
       oTabla.RGBVerde = IIf(IsNull(oRecordset!RGBVerde), 0, oRecordset!RGBVerde)
       oTabla.RGBAzul = IIf(IsNull(oRecordset!RGBAzul), 0, oRecordset!RGBAzul)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoAtencionHospCenso)
   oTabla.IdRangoCensoHosp = 0
   oTabla.RangoInicial = 0
   oTabla.RangoFinal = 0
   oTabla.RGBRojo = 0
   oTabla.RGBVerde = 0
   oTabla.RGBAzul = 0
End Sub
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: AtenHospCenso
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: AtenHospCenso
'        Autor:   DAniel Barrantes B.
'        Fecha: 27/02/2015 03:24:07
'------------------------------------------------------------------------------------
