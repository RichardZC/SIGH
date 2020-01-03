VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlanesCatalogosServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: PlanesCatalogosServicios
'        Autor: William Castro Grijalva
'        Fecha: 09/07/2005 08:28:26 p.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------
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
'   Descripción:    Inserta un registro a la tabla PlanesCatalogosServicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOPlanesCatalogosServicios) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "PlanesCatalogosServiciosAgregar"
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.IdProducto = 0, Null, oTabla.IdProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, IIf(oTabla.PrecioUnitario = 0, Null, oTabla.PrecioUnitario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPlan", adInteger, adParamInput, 0, IIf(oTabla.IdPlan = 0, Null, oTabla.IdPlan)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPlanCatalogo", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdPlanCatalogo = .Parameters("@IdPlanCatalogo")
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
'   Descripción:    modifica un registro a la tabla PlanesCatalogosServicios
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOPlanesCatalogosServicios) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "PlanesCatalogosServiciosModificar"
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.IdProducto = 0, Null, oTabla.IdProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, IIf(oTabla.PrecioUnitario = 0, Null, oTabla.PrecioUnitario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPlan", adInteger, adParamInput, 0, IIf(oTabla.IdPlan = 0, Null, oTabla.IdPlan)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPlanCatalogo", adInteger, adParamInput, 0, IIf(oTabla.IdPlanCatalogo = 0, Null, oTabla.IdPlanCatalogo)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla PlanesCatalogosServicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOPlanesCatalogosServicios) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "PlanesCatalogosServiciosEliminar"
           Set oParameter = .CreateParameter("@IdPlanCatalogo", adInteger, adParamInput, 0, IIf(oTabla.IdPlanCatalogo = 0, Null, oTabla.IdPlanCatalogo)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla PlanesCatalogosServicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOPlanesCatalogosServicios) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "PlanesCatalogosServiciosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdPlanCatalogo", adInteger, adParamInput, 0, oTabla.IdPlanCatalogo): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdProducto = IIf(IsNull(oRecordset!IdProducto), 0, oRecordset!IdProducto)
       oTabla.PrecioUnitario = IIf(IsNull(oRecordset!PrecioUnitario), 0, oRecordset!PrecioUnitario)
       oTabla.IdPlan = IIf(IsNull(oRecordset!IdPlan), 0, oRecordset!IdPlan)
       oTabla.IdPlanCatalogo = IIf(IsNull(oRecordset!IdPlanCatalogo), 0, oRecordset!IdPlanCatalogo)
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

Function ObtenerPrecioUnitario(lIdProducto As Long, lIdTipoFinanciamiento As Long) As Currency
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String

    ObtenerPrecioUnitario = 0
    
    sSQL = "select PrecioUnitario from FactCatalogoServiciosHosp where " + _
            " IdProducto = " & lIdProducto & " and IdTipoFinanciamiento = " & lIdTipoFinanciamiento
 
   With oCommand
     .CommandType = adCmdText
     Set .ActiveConnection = mo_Conexion
     .CommandText = sSQL
     Set oRecordset = .Execute
   End With
   
    If Not (oRecordset.EOF And oRecordset.BOF) Then
        ObtenerPrecioUnitario = oRecordset!PrecioUnitario
    Else
        ObtenerPrecioUnitario = -1
    End If
   
   oRecordset.Close
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
