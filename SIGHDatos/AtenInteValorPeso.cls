VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtenInteValorPeso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: AtenInteValorPeso
'        Autor: Daniel Enrique Barrantes
'        Fecha: 11/11/2014 07:12:29
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
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla AtenInteValorPeso
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOAtenInteValorPeso) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteValorPesoAgregar"
   Set oParameter = .CreateParameter("@IdValorPeso", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSexo = 0, Null, oTabla.IdTipoSexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadMeses", adInteger, adParamInput, 0, IIf(oTabla.EdadMeses = 0, Null, oTabla.EdadMeses))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDesviacion", adInteger, adParamInput, 0, IIf(oTabla.NroDesviacion = 0, Null, oTabla.NroDesviacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorPeso", adCurrency, adParamInput, 0, oTabla.ValorPeso)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdValorPeso = .Parameters("@IdValorPeso")
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
'   Descripción:    modifica un registro a la tabla AtenInteValorPeso
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtenInteValorPeso) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteValorPesoModificar"
   Set oParameter = .CreateParameter("@IdValorPeso", adInteger, adParamInput, 0, IIf(oTabla.IdValorPeso = 0, Null, oTabla.IdValorPeso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSexo = 0, Null, oTabla.IdTipoSexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadMeses", adInteger, adParamInput, 0, IIf(oTabla.EdadMeses = 0, Null, oTabla.EdadMeses))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDesviacion", adInteger, adParamInput, 0, IIf(oTabla.NroDesviacion = 0, Null, oTabla.NroDesviacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorPeso", adCurrency, adParamInput, 0, oTabla.ValorPeso)
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
'   Descripción:    elimia o borra un registro a la tabla AtenInteValorPeso
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtenInteValorPeso) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtenInteValorPesoEliminar"
   Set oParameter = .CreateParameter("@IdValorPeso", adInteger, adParamInput, 0, oTabla.IdValorPeso)
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
'   Descripción:    Selecciona un unico registro de la tabla AtenInteValorPeso
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtenInteValorPeso) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtenInteValorPesoSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdValorPeso", adInteger, adParamInput, 0, oTabla.IdValorPeso)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdValorPeso = IIf(IsNull(oRecordset!IdValorPeso), 0, oRecordset!IdValorPeso)
       oTabla.IdTipoSexo = IIf(IsNull(oRecordset!IdTipoSexo), 0, oRecordset!IdTipoSexo)
       oTabla.EdadMeses = IIf(IsNull(oRecordset!EdadMeses), 0, oRecordset!EdadMeses)
       oTabla.NroDesviacion = IIf(IsNull(oRecordset!NroDesviacion), 0, oRecordset!NroDesviacion)
       oTabla.ValorPeso = oRecordset!ValorPeso
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOAtenInteValorPeso)
   oTabla.IdValorPeso = 0
   oTabla.IdTipoSexo = 0
   oTabla.EdadMeses = 0
   oTabla.NroDesviacion = 0
   oTabla.ValorPeso = 0
End Sub

Function SeleccionarPorSexoYDesviacion(ByVal oTabla As DOAtenInteValorPeso) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtenInteValorPesoPorSexoYDesviacion"
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
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: AtenInteValorPeso
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: AtenInteValorPeso
'        Autor:   DAniel Barrantes B.
'        Fecha: 11/11/2014 07:12:29
'------------------------------------------------------------------------------------
 

