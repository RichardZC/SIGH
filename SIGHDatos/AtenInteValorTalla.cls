VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtenInteValorTalla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtenInteValorTalla
'        Programado por: Garay M
'        Fecha: Noviembre 2014
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

Function Insertar(ByVal oTabla As DOAtenInteValorTalla) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteValorTallaAgregar"
   Set oParameter = .CreateParameter("@IdValorTalla", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSexo = 0, Null, oTabla.IdTipoSexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadMeses", adInteger, adParamInput, 0, IIf(oTabla.EdadMeses = 0, Null, oTabla.EdadMeses))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDesviacion", adInteger, adParamInput, 0, IIf(oTabla.NroDesviacion = 0, Null, oTabla.NroDesviacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorTalla", adCurrency, adParamInput, 0, oTabla.ValorTalla)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdValorTalla = .Parameters("@IdValorTalla")
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
'   Descripción:    modifica un registro a la tabla AtenInteValorTalla
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtenInteValorTalla) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteValorTallaModificar"
   Set oParameter = .CreateParameter("@IdValorTalla", adInteger, adParamInput, 0, IIf(oTabla.IdValorTalla = 0, Null, oTabla.IdValorTalla))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSexo = 0, Null, oTabla.IdTipoSexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadMeses", adInteger, adParamInput, 0, IIf(oTabla.EdadMeses = 0, Null, oTabla.EdadMeses))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDesviacion", adInteger, adParamInput, 0, IIf(oTabla.NroDesviacion = 0, Null, oTabla.NroDesviacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorTalla", adCurrency, adParamInput, 0, oTabla.ValorTalla)
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
'   Descripción:    elimia o borra un registro a la tabla AtenInteValorTalla
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtenInteValorTalla) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteValorTallaEliminar"
   Set oParameter = .CreateParameter("@IdValorTalla", adInteger, adParamInput, 0, oTabla.IdValorTalla)
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
'   Descripción:    Selecciona un unico registro de la tabla AtenInteValorTalla
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtenInteValorTalla) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtenInteValorTallaSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdValorTalla", adInteger, adParamInput, 0, oTabla.IdValorTalla)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdValorTalla = IIf(IsNull(oRecordset!IdValorTalla), 0, oRecordset!IdValorTalla)
       oTabla.IdTipoSexo = IIf(IsNull(oRecordset!IdTipoSexo), 0, oRecordset!IdTipoSexo)
       oTabla.EdadMeses = IIf(IsNull(oRecordset!EdadMeses), 0, oRecordset!EdadMeses)
       oTabla.NroDesviacion = IIf(IsNull(oRecordset!NroDesviacion), 0, oRecordset!NroDesviacion)
       oTabla.ValorTalla = oRecordset!ValorTalla
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOAtenInteValorTalla)
   oTabla.IdValorTalla = 0
   oTabla.IdTipoSexo = 0
   oTabla.EdadMeses = 0
   oTabla.NroDesviacion = 0
   oTabla.ValorTalla = 0
End Sub

Function SeleccionarPorSexoYDesviacion(ByVal oTabla As DOAtenInteValorTalla) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtenInteValorTallaPorSexoYDesviacion"
        Set oParameter = .CreateParameter("@InIdTipoSexo", adInteger, adParamInput, 0, oTabla.IdTipoSexo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroDesviacion", adInteger, adParamInput, 0, oTabla.NroDesviacion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorSexoYDesviacion = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

