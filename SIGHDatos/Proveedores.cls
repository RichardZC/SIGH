VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla Proveedores
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String

Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property


'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla Proveedores
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOProveedores) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ProveedoresAgregar"
   Set oParameter = .CreateParameter("@idProveedor", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Ruc", adVarChar, adParamInput, 11, IIf(oTabla.RUC = "", Null, oTabla.RUC))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RazonSocial", adVarChar, adParamInput, 100, IIf(oTabla.RazonSocial = "", Null, oTabla.RazonSocial))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@email", adVarChar, adParamInput, 100, IIf(oTabla.Email = "", Null, oTabla.Email))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@direccion", adVarChar, adParamInput, 100, IIf(oTabla.Direccion = "", Null, oTabla.Direccion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.idProveedor = .Parameters("@idProveedor")
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
'   Descripción:    modifica un registro a la tabla Proveedores
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOProveedores) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ProveedoresModificar"
   Set oParameter = .CreateParameter("@idProveedor", adInteger, adParamInput, 0, IIf(oTabla.idProveedor = 0, Null, oTabla.idProveedor))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Ruc", adVarChar, adParamInput, 11, IIf(oTabla.RUC = "", Null, oTabla.RUC))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RazonSocial", adVarChar, adParamInput, 100, IIf(oTabla.RazonSocial = "", Null, oTabla.RazonSocial))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@email", adVarChar, adParamInput, 100, IIf(oTabla.Email = "", Null, oTabla.Email))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@direccion", adVarChar, adParamInput, 100, IIf(oTabla.Direccion = "", Null, oTabla.Direccion))
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
'   Descripción:    elimia o borra un registro a la tabla Proveedores
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOProveedores) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ProveedoresEliminar"
   Set oParameter = .CreateParameter("@idProveedor", adInteger, adParamInput, 0, oTabla.idProveedor)
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
'   Descripción:    Selecciona un unico registro de la tabla Proveedores
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOProveedores) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ProveedoresSeleccionarPorId"
     Set oParameter = .CreateParameter("@idProveedor", adInteger, adParamInput, 0, oTabla.idProveedor)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idProveedor = IIf(IsNull(oRecordset!idProveedor), 0, oRecordset!idProveedor)
       oTabla.RUC = IIf(IsNull(oRecordset!RUC), "", oRecordset!RUC)
       oTabla.RazonSocial = IIf(IsNull(oRecordset!RazonSocial), "", oRecordset!RazonSocial)
       oTabla.Email = IIf(IsNull(oRecordset!Email), "", oRecordset!Email)
       oTabla.Direccion = IIf(IsNull(oRecordset!Direccion), "", oRecordset!Direccion)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
   Set oRecordset = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOProveedores)
   oTabla.idProveedor = 0
   oTabla.RUC = ""
   oTabla.RazonSocial = ""
End Sub

'MARIANO  debb2014b
Function SeleccionarSegunFiltro(lcFiltro As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FarmProveedorFiltrar"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarSegunFiltro = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

