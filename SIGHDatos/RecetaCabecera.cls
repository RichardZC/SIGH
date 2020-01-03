VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RecetaCabecera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla RecetaCabecera
'        Programado por: Barrantes D
'        Fecha: Marzo 2012
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
'   Descripción:    Inserta un registro a la tabla RecetaCabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoRecetaCabecera) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "RecetaCabeceraAgregar"
   Set oParameter = .CreateParameter("@idReceta", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.IdPuntoCarga = 0, Null, oTabla.IdPuntoCarga))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaReceta", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fechaReceta = 0, Null, oTabla.fechaReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.idCuentaAtencion = 0, Null, oTabla.idCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idServicioReceta", adInteger, adParamInput, 0, IIf(oTabla.idServicioReceta = 0, Null, oTabla.idServicioReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, IIf(oTabla.IdEstado = 0, Null, oTabla.IdEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idMedicoReceta", adInteger, adParamInput, 0, IIf(oTabla.idMedicoReceta = 0, Null, oTabla.idMedicoReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaVigencia", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaVigencia = 0, Null, oTabla.FechaVigencia))
   .Parameters.Append oParameter
  
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdReceta = .Parameters("@idReceta")
End With
 
Insertar = True
ms_MensajeError = ""
 GrabaHistorico oTabla
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


Sub GrabaHistorico(oTabla As DoRecetaCabecera)
    On Error GoTo ErrHistorico
    Dim lcSql As String
    Dim oRsParametros As New Recordset
    Dim oRsMDB As New Recordset
    Dim oConexionMDB As New Connection
    lcSql = "select ValorTexto from Parametros where idparametro=581"
    oRsParametros.Open lcSql, mo_Conexion, adOpenKeyset, adLockOptimistic
    oConexionMDB.CommandTimeout = 900
    oConexionMDB.CursorLocation = adUseClient
    oConexionMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =" & Trim(oRsParametros!ValorTexto) & "\parametros.mdb"
    lcSql = "INSERT INTO RecetaCabecera" & _
"           (IdPuntoCarga" & _
"           ,FechaReceta" & _
"           ,idCuentaAtencion" & _
"           ,idServicioReceta" & _
"           ,idEstado" & _
"           ,idComprobantePago" & _
"           ,idMedicoReceta" & _
"           ,fechaVigencia,idreceta)" & _
"     Values" & _
"           (" & _
"           " & oTabla.IdPuntoCarga & _
"           ,'" & oTabla.fechaReceta & _
"           '," & oTabla.idCuentaAtencion & _
"           ," & oTabla.idServicioReceta & _
"           ," & oTabla.IdEstado & _
"           ," & oTabla.IdComprobantePago & _
"           ," & oTabla.idMedicoReceta & _
"           ,'" & oTabla.FechaVigencia & "'," & oTabla.IdReceta & _
"           )"
    oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    oConexionMDB.Close
ErrHistorico:
    Set oRsParametros = Nothing
    Set oRsMDB = Nothing
    Set oConexionMDB = Nothing
End Sub


'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla RecetaCabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoRecetaCabecera) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "RecetaCabeceraModificar"
   Set oParameter = .CreateParameter("@idReceta", adInteger, adParamInput, 0, IIf(oTabla.IdReceta = 0, Null, oTabla.IdReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.IdPuntoCarga = 0, Null, oTabla.IdPuntoCarga))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaReceta", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fechaReceta = 0, Null, oTabla.fechaReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.idCuentaAtencion = 0, Null, oTabla.idCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idServicioReceta", adInteger, adParamInput, 0, IIf(oTabla.idServicioReceta = 0, Null, oTabla.idServicioReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, IIf(oTabla.IdEstado = 0, Null, oTabla.IdEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idMedicoReceta", adInteger, adParamInput, 0, IIf(oTabla.idMedicoReceta = 0, Null, oTabla.idMedicoReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaVigencia", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaVigencia = 0, Null, oTabla.FechaVigencia))
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
'   Descripción:    elimia o borra un registro a la tabla RecetaCabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoRecetaCabecera) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "RecetaCabeceraEliminar"
   Set oParameter = .CreateParameter("@idReceta", adInteger, adParamInput, 0, oTabla.IdReceta)
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
'   Descripción:    Selecciona un unico registro de la tabla RecetaCabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoRecetaCabecera) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "RecetaCabeceraSeleccionarPorId"
     Set oParameter = .CreateParameter("@idReceta", adInteger, adParamInput, 0, oTabla.IdReceta)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       CargaCampos oTabla, oRecordset
   End If
   SeleccionarPorId = True
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoRecetaCabecera)
   oTabla.IdReceta = 0
   oTabla.IdPuntoCarga = 0
   oTabla.fechaReceta = Now
   oTabla.idCuentaAtencion = 0
   oTabla.idServicioReceta = 0
   oTabla.IdEstado = 0
   oTabla.IdComprobantePago = 0
   oTabla.idMedicoReceta = 0
End Sub

Sub CargaCampos(ByRef oTabla As DoRecetaCabecera, oRecordset As Recordset)
       oTabla.IdReceta = IIf(IsNull(oRecordset!IdReceta), 0, oRecordset!IdReceta)
       oTabla.IdPuntoCarga = IIf(IsNull(oRecordset!IdPuntoCarga), 0, oRecordset!IdPuntoCarga)
       oTabla.fechaReceta = IIf(IsNull(oRecordset!fechaReceta), 0, oRecordset!fechaReceta)
       oTabla.idCuentaAtencion = IIf(IsNull(oRecordset!idCuentaAtencion), 0, oRecordset!idCuentaAtencion)
       oTabla.idServicioReceta = IIf(IsNull(oRecordset!idServicioReceta), 0, oRecordset!idServicioReceta)
       oTabla.IdEstado = IIf(IsNull(oRecordset!IdEstado), 0, oRecordset!IdEstado)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.idMedicoReceta = IIf(IsNull(oRecordset!idMedicoReceta), 0, oRecordset!idMedicoReceta)
       oTabla.FechaVigencia = IIf(IsNull(oRecordset!FechaVigencia), 0, oRecordset!FechaVigencia)

End Sub


Function SeleccionarPorIdComprobantePagoRecetaDetalleItem(ByVal oTabla As DoRecetaCabecera) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorIdComprobantePagoRecetaDetalleItem = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "RecetaCabeceraSeleccionarPorIdComprobantePago"
     Set oParameter = .CreateParameter("@idComprobantePago", adInteger, adParamInput, 0, oTabla.IdComprobantePago)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       CargaCampos oTabla, oRecordset
   End If
   SeleccionarPorIdComprobantePagoRecetaDetalleItem = True
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function



Function SeleccionarPorIdCuentaAtencion(lnIdCuentaAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "RecetaCabeceraSeleccionarPorIdCuentaAtencion"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorIdCuentaAtencion = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function ConsultaRecetaCabeceraFarmacia(oTabla As DoRecetaCabecera, lcDocumentoDespacho As String) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ConsultaRecetaCabeceraXIdPuntoCarga"
        Set oParameter = .CreateParameter("@DocumentoDespacho", adVarChar, adParamInput, 20, lcDocumentoDespacho): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       CargaCampos oTabla, oRecordset
       ConsultaRecetaCabeceraFarmacia = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



