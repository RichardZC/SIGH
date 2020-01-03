VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProProcedimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla ProProcedimientos
'        Programado por: Castro W
'        Fecha: Agosto 2006
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
'   Descripción:    Inserta un registro a la tabla ProProcedimientos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOProProcedimientos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ProProcedimientosAgregar"
   Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, IIf(oTabla.IdPrograma = 0, Null, oTabla.IdPrograma))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, IIf(oTabla.IdProCabecera = 0, Null, oTabla.IdProCabecera))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, IIf(oTabla.IdControl = 0, Null, oTabla.IdControl))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, 0, oTabla.IdDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdResultado", adInteger, adParamInput, 0, IIf(oTabla.IdResultado = 0, Null, oTabla.IdResultado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@labConfHIS", adVarChar, adParamInput, 3, IIf(oTabla.labConfHIS = "", Null, oTabla.labConfHIS))
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
'   Descripción:    modifica un registro a la tabla ProProcedimientos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOProProcedimientos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ProProcedimientosModificar"
   Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, IIf(oTabla.IdPrograma = 0, Null, oTabla.IdPrograma))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, IIf(oTabla.IdProCabecera = 0, Null, oTabla.IdProCabecera))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, IIf(oTabla.IdControl = 0, Null, oTabla.IdControl))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdResultado", adInteger, adParamInput, 0, IIf(oTabla.IdResultado = 0, Null, oTabla.IdResultado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@labConfHIS", adVarChar, adParamInput, 3, IIf(oTabla.labConfHIS = "", Null, oTabla.labConfHIS))
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
'   Descripción:    elimia o borra un registro a la tabla ProProcedimientos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOProProcedimientos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ProProcedimientosEliminar"
   Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, IIf(oTabla.IdPrograma = 0, Null, oTabla.IdPrograma))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, IIf(oTabla.IdProCabecera = 0, Null, oTabla.IdProCabecera))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, IIf(oTabla.IdControl = 0, Null, oTabla.IdControl))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
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

Sub SetDefaults(ByVal oTabla As DOProProcedimientos)
   oTabla.IdPrograma = 0
   oTabla.IdProCabecera = 0
   oTabla.IdControl = 0
   oTabla.IdDiagnostico = 0
   oTabla.idProducto = 0
   oTabla.IdResultado = 0
   oTabla.labConfHIS = ""
End Sub
