VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImagMovimientoCPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla ImagMovimientoCPT
'        Programado por: Barrantes D
'        Fecha: Abril 2009
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
'   Descripción:    Inserta un registro a la tabla ImagMovimientoCPT
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOImagMovimientoCPT) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ImagMovimientoCPTAgregar"
   Set oParameter = .CreateParameter("@idMovimiento", adInteger, adParamInput, 0, IIf(oTabla.idMovimiento = 0, Null, oTabla.idMovimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProductoCPT", adInteger, adParamInput, 0, IIf(oTabla.idProductoCPT = 0, Null, oTabla.idProductoCPT))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cantidad", adInteger, adParamInput, 0, IIf(oTabla.Cantidad = 0, Null, oTabla.Cantidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@precio", adCurrency, adParamInput, 0, oTabla.PRECIO)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@importe", adCurrency, adParamInput, 0, oTabla.Importe)
   .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ResultadoAutomatico", adInteger, adParamInput, 0, oTabla.ResultadoAutomatico)
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
'   Descripción:    modifica un registro a la tabla ImagMovimientoCPT
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOImagMovimientoCPT) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ImagMovimientoCPTModificar"
   Set oParameter = .CreateParameter("@idMovimiento", adInteger, adParamInput, 0, IIf(oTabla.idMovimiento = 0, Null, oTabla.idMovimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProductoCPT", adInteger, adParamInput, 0, IIf(oTabla.idProductoCPT = 0, Null, oTabla.idProductoCPT))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cantidad", adInteger, adParamInput, 0, IIf(oTabla.Cantidad = 0, Null, oTabla.Cantidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@precio", adCurrency, adParamInput, 0, oTabla.PRECIO)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@importe", adCurrency, adParamInput, 0, oTabla.Importe)
   .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ResultadoAutomatico", adInteger, adParamInput, 0, oTabla.ResultadoAutomatico)
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
'   Descripción:    elimia o borra un registro a la tabla ImagMovimientoCPT
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOImagMovimientoCPT) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ImagMovimientoCPTEliminar"
   Set oParameter = .CreateParameter("@idMovimiento", adInteger, adParamInput, 0, oTabla.idMovimiento)
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
'   Descripción:    Selecciona un unico registro de la tabla ImagMovimientoCPT
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOImagMovimientoCPT) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ImagMovimientoCPTSeleccionarPorId"
     Set oParameter = .CreateParameter("@idMovimiento", adInteger, adParamInput, 0, oTabla.idMovimiento)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idMovimiento = IIf(IsNull(oRecordset!idMovimiento), 0, oRecordset!idMovimiento)
       oTabla.idProductoCPT = IIf(IsNull(oRecordset!idProductoCPT), 0, oRecordset!idProductoCPT)
       oTabla.Cantidad = IIf(IsNull(oRecordset!Cantidad), 0, oRecordset!Cantidad)
       oTabla.PRECIO = oRecordset!PRECIO
       oTabla.Importe = oRecordset!Importe
       oTabla.ResultadoAutomatico = IIf(IsNull(oRecordset!ResultadoAutomatico), 0, oRecordset!ResultadoAutomatico)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOImagMovimientoCPT)
   oTabla.idMovimiento = 0
   oTabla.idProductoCPT = 0
   oTabla.Cantidad = 0
   oTabla.PRECIO = 0
   oTabla.Importe = 0
End Sub
