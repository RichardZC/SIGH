VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasServGeograf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para SERVICIOS GEOGRAFICOS
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_MensajeError As String
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

Function PaisesSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Paises
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set PaisesSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
    Set oConexion = Nothing

End Function

Function DepartamentosSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Departamentos
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set DepartamentosSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oTabla = Nothing
    Set oConexion = Nothing

End Function
Function ProvinciasSeleccionarPorDepartamento(IdDepartamento As Integer) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Provincias
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set ProvinciasSeleccionarPorDepartamento = oTabla.SeleccionarPorDepartamento(IdDepartamento)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function DistritoSeleccionarPorProvincia(IdProvincia As Integer) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Distritos
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set DistritoSeleccionarPorProvincia = oTabla.SeleccionarPorProvincia(IdProvincia)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function DistritoSeleccionarProvincia(IdDistrito As Long) As DOProvincia
Dim oConexion As New ADODB.Connection
Dim oTabla As New Distritos
Dim oDODistrito As New DODistrito
Dim oProvincia As New Provincias
Dim oDOProvincia As New DOProvincia
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    Set oProvincia.Conexion = oConexion
    
    oDODistrito.IdDistrito = IdDistrito
    If oTabla.SeleccionarPorId(oDODistrito) Then
        oDOProvincia.IdProvincia = oDODistrito.IdProvincia
        If oProvincia.SeleccionarPorId(oDOProvincia) Then
            Set DistritoSeleccionarProvincia = oDOProvincia
        End If
    End If
    
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
    Set oProvincia = Nothing
    Set oDODistrito = Nothing
End Function
Function ProvinciaSeleccionarDepartamento(IdDepartamento As Long) As doDepartamento
Dim oConexion As New ADODB.Connection
Dim oTabla As New Provincias
Dim oDOProvincia As New DOProvincia
Dim oDepartamento As New Departamentos
Dim oDODepartamento As New doDepartamento
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    oDOProvincia.IdDepartamento = IdDepartamento
    If oTabla.SeleccionarPorId(oDOProvincia) Then
        oDODepartamento.IdDepartamento = oDOProvincia.IdDepartamento
        If oDepartamento.SeleccionarPorId(oDODepartamento) Then
            Set ProvinciaSeleccionarDepartamento = oDODepartamento
        End If
    End If
    
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
    Set oDOProvincia = Nothing
    Set oDepartamento = Nothing

End Function


Function CentroPobladoSeleccionarPorDistrito(IdDistrito As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CentroPoblados
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set CentroPobladoSeleccionarPorDistrito = oTabla.SeleccionarPorDistrito(IdDistrito)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function PaisesXfiltro(lcFiltro As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PaisesXfiltro"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   oRecordset.Filter = lcFiltro
   Set PaisesXfiltro = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description

End Function

Function DepartamentosSeleccionarXidentificador(lnIdDepartamento As Long) As Recordset


On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DepartamentosSeleccionarXidentificador"
        Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamInput, 0, lnIdDepartamento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set DepartamentosSeleccionarXidentificador = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



