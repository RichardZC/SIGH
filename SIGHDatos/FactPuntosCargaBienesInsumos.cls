VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FactPuntosCargaBienesInsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FactPuntosCargaBienesInsumos
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
'   Descripción:    Inserta un registro a la tabla FactPuntosCargaBienesInsumos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOFactPuntoCargaBienesInsumos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FactPuntosCargaBienesInsumosAgregar"
           Set oParameter = .CreateParameter("@IdSubGrupoFarmacologico", adInteger, adParamInput, 0, IIf(oTabla.IdSubGrupoFarmacologico = 0, Null, oTabla.IdSubGrupoFarmacologico)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.idPuntoCarga = 0, Null, oTabla.idPuntoCarga)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPuntoCargaBienInsumo", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdPuntoCargaBienInsumo = .Parameters("@IdPuntoCargaBienInsumo")
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
'   Descripción:    modifica un registro a la tabla FactPuntosCargaBienesInsumos
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOFactPuntoCargaBienesInsumos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FactPuntosCargaBienesInsumosModificar"
           Set oParameter = .CreateParameter("@IdSubGrupoFarmacologico", adInteger, adParamInput, 0, IIf(oTabla.IdSubGrupoFarmacologico = 0, Null, oTabla.IdSubGrupoFarmacologico)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.idPuntoCarga = 0, Null, oTabla.idPuntoCarga)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPuntoCargaBienInsumo", adInteger, adParamInput, 0, IIf(oTabla.IdPuntoCargaBienInsumo = 0, Null, oTabla.IdPuntoCargaBienInsumo)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla FactPuntosCargaBienesInsumos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOFactPuntoCargaBienesInsumos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FactPuntosCargaBienesInsumosEliminar"
           Set oParameter = .CreateParameter("@IdPuntoCargaBienInsumo", adInteger, adParamInput, 0, IIf(oTabla.IdPuntoCargaBienInsumo = 0, Null, oTabla.IdPuntoCargaBienInsumo)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla FactPuntosCargaBienesInsumos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOFactPuntoCargaBienesInsumos) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FactPuntosCargaBienesInsumosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdPuntoCargaBienInsumo", adInteger, adParamInput, 0, oTabla.IdPuntoCargaBienInsumo): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdSubGrupoFarmacologico = IIf(IsNull(oRecordset!IdSubGrupoFarmacologico), 0, oRecordset!IdSubGrupoFarmacologico)
       oTabla.idPuntoCarga = IIf(IsNull(oRecordset!idPuntoCarga), 0, oRecordset!idPuntoCarga)
       oTabla.IdPuntoCargaBienInsumo = IIf(IsNull(oRecordset!IdPuntoCargaBienInsumo), 0, oRecordset!IdPuntoCargaBienInsumo)
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


