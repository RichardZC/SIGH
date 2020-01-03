VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProDiagnosticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla ProDiagnosticos
'        Programado por: Garay M
'        Fecha: Agosto 2014
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
'   Descripción:    Inserta un registro a la tabla ProDiagnosticos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOProDiagnosticos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ProDiagnosticosAgregar"
   Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, IIf(oTabla.IdPrograma = 0, Null, oTabla.IdPrograma))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, IIf(oTabla.IdProCabecera = 0, Null, oTabla.IdProCabecera))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, IIf(oTabla.IdControl = 0, Null, oTabla.IdControl))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Principal", adBoolean, adParamInput, 0, oTabla.Principal)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@labConfHIS", adVarChar, adParamInput, 3, IIf(oTabla.labConfHIS = "", Null, oTabla.labConfHIS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdSubclasificacionDx", adInteger, adParamInput, 0, IIf(oTabla.IdSubClasificacionDX = 0, Null, oTabla.IdSubClasificacionDX))
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
'   Descripción:    modifica un registro a la tabla ProDiagnosticos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOProDiagnosticos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ProDiagnosticosModificar"
   Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, IIf(oTabla.IdPrograma = 0, Null, oTabla.IdPrograma))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProCabecera", adInteger, adParamInput, 0, IIf(oTabla.IdProCabecera = 0, Null, oTabla.IdProCabecera))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdControl", adInteger, adParamInput, 0, IIf(oTabla.IdControl = 0, Null, oTabla.IdControl))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Principal", adBoolean, adParamInput, 0, IIf(oTabla.Principal = 0, Null, oTabla.Principal))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@labConfHIS", adVarChar, adParamInput, 3, IIf(oTabla.labConfHIS = "", Null, oTabla.labConfHIS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdSubclasificacionDx", adInteger, adParamInput, 0, IIf(oTabla.IdSubClasificacionDX = 0, Null, oTabla.IdSubClasificacionDX))
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
'   Descripción:    elimia o borra un registro a la tabla ProDiagnosticos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOProDiagnosticos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ProDiagnosticosEliminar"
   Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, oTabla.IdPrograma)
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
'   Descripción:    Selecciona un unico registro de la tabla ProDiagnosticos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOProDiagnosticos) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ProDiagnosticosSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdPrograma", adInteger, adParamInput, 0, oTabla.IdPrograma)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdPrograma = IIf(IsNull(oRecordset!IdPrograma), 0, oRecordset!IdPrograma)
       oTabla.IdProCabecera = IIf(IsNull(oRecordset!IdProCabecera), 0, oRecordset!IdProCabecera)
       oTabla.IdControl = IIf(IsNull(oRecordset!IdControl), 0, oRecordset!IdControl)
       oTabla.IdDiagnostico = IIf(IsNull(oRecordset!IdDiagnostico), 0, oRecordset!IdDiagnostico)
       oTabla.Principal = IIf(IsNull(oRecordset!Principal), 0, oRecordset!Principal)
       oTabla.labConfHIS = IIf(IsNull(oRecordset!labConfHIS), "", oRecordset!labConfHIS)
       oTabla.IdSubClasificacionDX = IIf(IsNull(oRecordset!IdSubClasificacionDX), "", oRecordset!IdSubClasificacionDX)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOProDiagnosticos)
   oTabla.IdPrograma = 0
   oTabla.IdProCabecera = 0
   oTabla.IdControl = 0
   oTabla.IdDiagnostico = 0
   oTabla.Principal = False
   oTabla.labConfHIS = ""
   oTabla.IdSubClasificacionDX = 0
End Sub

