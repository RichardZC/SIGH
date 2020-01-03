VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FarmaciaRecetasDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: FarmaciaRecetasDetalle
'        Autor: William Castro Grijalva
'        Fecha: 05/08/2005 08:46:44
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
'   Descripción:    Inserta un registro a la tabla FarmaciaRecetasDetalle
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOFarmaciaRecetasDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FarmaciaRecetasDetalleAgregar"
           Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, IIf(oTabla.Cantidad = "", Null, oTabla.Cantidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFacturacionBienes", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionBienes = "", Null, oTabla.IdFacturacionBienes)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionReceta", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionReceta = "", Null, oTabla.IdAtencionReceta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdRecetaDetalle", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdRecetaDetalle = .Parameters("@IdRecetaDetalle")
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
'   Descripción:    modifica un registro a la tabla FarmaciaRecetasDetalle
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOFarmaciaRecetasDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FarmaciaRecetasDetalleModificar"
           Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, IIf(oTabla.Cantidad = "", Null, oTabla.Cantidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFacturacionBienes", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionBienes = "", Null, oTabla.IdFacturacionBienes)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionReceta", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionReceta = "", Null, oTabla.IdAtencionReceta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdRecetaDetalle", adInteger, adParamInput, 0, IIf(oTabla.IdRecetaDetalle = "", Null, oTabla.IdRecetaDetalle)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla FarmaciaRecetasDetalle
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOFarmaciaRecetasDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FarmaciaRecetasDetalleEliminar"
           Set oParameter = .CreateParameter("@IdRecetaDetalle", adInteger, adParamInput, 0, IIf(oTabla.IdRecetaDetalle = "", Null, oTabla.IdRecetaDetalle)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla FarmaciaRecetasDetalle
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOFarmaciaRecetasDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FarmaciaRecetasDetalleSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdRecetaDetalle", adInteger, adParamInput, 0, oTabla.IdRecetaDetalle): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.Cantidad = IIf(IsNull(oRecordset!Cantidad), "", oRecordset!Cantidad)
       oTabla.IdFacturacionBienes = IIf(IsNull(oRecordset!IdFacturacionBienes), "", oRecordset!IdFacturacionBienes)
       oTabla.IdAtencionReceta = IIf(IsNull(oRecordset!IdAtencionReceta), "", oRecordset!IdAtencionReceta)
       oTabla.IdRecetaDetalle = IIf(IsNull(oRecordset!IdRecetaDetalle), "", oRecordset!IdRecetaDetalle)
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

Function SeleccionarPorIdReceta(lIdReceta As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String
Dim sWhere As String

        Set SeleccionarPorIdReceta = Nothing
     
        sSQL = "select * from FarmaciaRecetasDetalle " + _
        " where IdAtencionReceta = " & lIdReceta
                
        With oCommand
           .CommandType = adCmdText
           Set .ActiveConnection = mo_Conexion
           .CommandText = sSQL
           Set oRecordset = .Execute
           Set oRecordset.ActiveConnection = Nothing
        End With
    
        Set SeleccionarPorIdReceta = oRecordset
    
        ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function


