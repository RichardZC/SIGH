VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FactOrdenesBienes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FactOrdenesBienes
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
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
'   Descripción:    Inserta un registro a la tabla FactOrdenesBienes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoFactOrdenesBienes) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FactOrdenesBienesAgregar"
   Set oParameter = .CreateParameter("@idOrden", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.IdPuntoCarga = 0, Null, oTabla.IdPuntoCarga))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, IIf(oTabla.movNumero = "", Null, oTabla.movNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPreventa", adInteger, adParamInput, 0, IIf(oTabla.idPreventa = 0, Null, oTabla.idPreventa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoFacturacion", adInteger, adParamInput, 0, oTabla.IdEstadoFacturacion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ImporteExonerado", adCurrency, adParamInput, 0, oTabla.ImporteExonerado)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuarioExonera", adInteger, adParamInput, 0, oTabla.idUsuarioExonera)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 150, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion))
   .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Paquete", adVarChar, adParamInput, 100, IIf(oTabla.Paquete = "", Null, oTabla.Paquete))
   .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Dni", adVarChar, adParamInput, 12, IIf(oTabla.DNI = "", Null, oTabla.DNI))
   .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@NombPaciente", adVarChar, adParamInput, 100, oTabla.nombpaciente)
   .Parameters.Append oParameter
   
   .Execute
   oTabla.IdOrden = .Parameters("@idOrden")
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
'   Descripción:    modifica un registro a la tabla FactOrdenesBienes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoFactOrdenesBienes) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FactOrdenesBienesModificar"
   Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.IdPuntoCarga = 0, Null, oTabla.IdPuntoCarga))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, IIf(oTabla.movNumero = "", Null, oTabla.movNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPreventa", adInteger, adParamInput, 0, IIf(oTabla.idPreventa = 0, Null, oTabla.idPreventa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoFacturacion", adInteger, adParamInput, 0, oTabla.IdEstadoFacturacion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ImporteExonerado", adCurrency, adParamInput, 0, oTabla.ImporteExonerado)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuarioExonera", adInteger, adParamInput, 0, oTabla.idUsuarioExonera)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 150, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion))
   .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Paquete", adVarChar, adParamInput, 100, IIf(oTabla.Paquete = "", Null, oTabla.Paquete))
   .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Dni", adVarChar, adParamInput, 12, IIf(oTabla.DNI = "", Null, oTabla.DNI))
   .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@NombPaciente", adVarChar, adParamInput, 100, oTabla.nombpaciente)
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
'   Descripción:    elimia o borra un registro a la tabla FactOrdenesBienes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoFactOrdenesBienes) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FactOrdenesBienesEliminar"
   Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, oTabla.IdOrden)
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
'   Descripción:    Selecciona un unico registro de la tabla FactOrdenesBienes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoFactOrdenesBienes) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FactOrdenesBienesSeleccionarPorId"
     Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, oTabla.IdOrden)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdOrden = IIf(IsNull(oRecordset!IdOrden), 0, oRecordset!IdOrden)
       oTabla.IdPuntoCarga = IIf(IsNull(oRecordset!IdPuntoCarga), 0, oRecordset!IdPuntoCarga)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.movNumero = IIf(IsNull(oRecordset!movNumero), "", oRecordset!movNumero)
       oTabla.movTipo = IIf(IsNull(oRecordset!movTipo), "", oRecordset!movTipo)
       oTabla.idPreventa = IIf(IsNull(oRecordset!idPreventa), 0, oRecordset!idPreventa)
       oTabla.FechaCreacion = IIf(IsNull(oRecordset!FechaCreacion), 0, oRecordset!FechaCreacion)
       oTabla.IdUsuario = IIf(IsNull(oRecordset!IdUsuario), 0, oRecordset!IdUsuario)
       oTabla.IdEstadoFacturacion = IIf(IsNull(oRecordset!IdEstadoFacturacion), 0, oRecordset!IdEstadoFacturacion)
       oTabla.ImporteExonerado = IIf(IsNull(oRecordset!ImporteExonerado), 0, oRecordset!ImporteExonerado)
       oTabla.DNI = IIf(IsNull(oRecordset!DNI), "", oRecordset!DNI)
       oTabla.nombpaciente = IIf(IsNull(oRecordset!nombpaciente), "", oRecordset!nombpaciente)
       oTabla.Paquete = IIf(IsNull(oRecordset!Paquete), "", oRecordset!Paquete)
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoFactOrdenesBienes)
   oTabla.IdOrden = 0
   oTabla.IdPuntoCarga = 0
   oTabla.IdPaciente = 0
   oTabla.IdCuentaAtencion = 0
   oTabla.IdComprobantePago = 0
   oTabla.movNumero = ""
   oTabla.movTipo = ""
   oTabla.idPreventa = 0
   oTabla.FechaCreacion = Now
   oTabla.IdUsuario = 0
   oTabla.IdEstadoFacturacion = 0
   oTabla.ImporteExonerado = 0
   oTabla.DNI = ""
   oTabla.nombpaciente = ""
   oTabla.Descripcion = ""
   oTabla.Paquete = ""
   
End Sub
Function SeleccionarPorIdPreventa(ByVal oTabla As DoFactOrdenesBienes) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorIdPreventa = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FactOrdenesBienesSeleccionarPorIdPreVenta"
     Set oParameter = .CreateParameter("@IdPreVenta", adInteger, adParamInput, 0, oTabla.idPreventa)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdOrden = IIf(IsNull(oRecordset!IdOrden), 0, oRecordset!IdOrden)
       oTabla.IdPuntoCarga = IIf(IsNull(oRecordset!IdPuntoCarga), 0, oRecordset!IdPuntoCarga)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.movNumero = IIf(IsNull(oRecordset!movNumero), "", oRecordset!movNumero)
       oTabla.movTipo = IIf(IsNull(oRecordset!movTipo), "", oRecordset!movTipo)
       oTabla.idPreventa = IIf(IsNull(oRecordset!idPreventa), 0, oRecordset!idPreventa)
       oTabla.FechaCreacion = IIf(IsNull(oRecordset!FechaCreacion), 0, oRecordset!FechaCreacion)
       oTabla.IdUsuario = IIf(IsNull(oRecordset!IdUsuario), 0, oRecordset!IdUsuario)
       oTabla.IdEstadoFacturacion = IIf(IsNull(oRecordset!IdEstadoFacturacion), 0, oRecordset!IdEstadoFacturacion)
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
       oTabla.nombpaciente = IIf(IsNull(oRecordset!nombpaciente), "", oRecordset!nombpaciente)
       oTabla.DNI = IIf(IsNull(oRecordset!DNI), "", oRecordset!DNI)
       oTabla.Paquete = IIf(IsNull(oRecordset!Paquete), "", oRecordset!Paquete)
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
   SeleccionarPorIdPreventa = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function SeleccionarPorMovNumero(ByVal oTabla As DoFactOrdenesBienes) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorMovNumero = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FactOrdenesBienesSeleccionarPorMovNumero"
     Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, IIf(oTabla.movNumero = "", Null, oTabla.movNumero))
     .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       cargaTabla oTabla, oRecordset                    'debb 2017
   
   ms_MensajeError = ""
   Else
       oTabla.IdOrden = 0
       oTabla.IdPuntoCarga = 0
       oTabla.IdPaciente = 0
       oTabla.IdCuentaAtencion = 0
       oTabla.IdComprobantePago = 0
       oTabla.movNumero = ""
       oTabla.movTipo = ""
       oTabla.idPreventa = 0
       oTabla.FechaCreacion = 0
       oTabla.IdUsuario = 0
       oTabla.IdEstadoFacturacion = 0
       oTabla.DNI = ""
       oTabla.nombpaciente = ""
       oTabla.Descripcion = ""
       oTabla.Paquete = ""
   End If
   SeleccionarPorMovNumero = True
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


'debb 2017
Function SeleccionarPorIdPreventaSoloUNO(ByVal oTabla As DoFactOrdenesBienes) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorIdPreventaSoloUNO = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FactOrdenesBienesSeleccionarPorIdPreVenta"
     Set oParameter = .CreateParameter("@IdPreVenta", adInteger, adParamInput, 0, oTabla.idPreventa)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       
       oRecordset.MoveFirst
       cargaTabla oTabla, oRecordset
       If oRecordset.RecordCount > 1 Then
          oRecordset.MoveNext
          Do While Not oRecordset.EOF
             If oRecordset!IdEstadoFacturacion = 1 Then
                cargaTabla oTabla, oRecordset
                Exit Do
             End If
             oRecordset.MoveNext
          Loop
       End If
       
       
       SeleccionarPorIdPreventaSoloUNO = True
       ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'debb 2017
Sub cargaTabla(oTabla As DoFactOrdenesBienes, oRecordset As Recordset)
       oTabla.IdOrden = IIf(IsNull(oRecordset!IdOrden), 0, oRecordset!IdOrden)
       oTabla.IdPuntoCarga = IIf(IsNull(oRecordset!IdPuntoCarga), 0, oRecordset!IdPuntoCarga)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.movNumero = IIf(IsNull(oRecordset!movNumero), "", oRecordset!movNumero)
       oTabla.movTipo = IIf(IsNull(oRecordset!movTipo), "", oRecordset!movTipo)
       oTabla.idPreventa = IIf(IsNull(oRecordset!idPreventa), 0, oRecordset!idPreventa)
       oTabla.FechaCreacion = IIf(IsNull(oRecordset!FechaCreacion), 0, oRecordset!FechaCreacion)
       oTabla.IdUsuario = IIf(IsNull(oRecordset!IdUsuario), 0, oRecordset!IdUsuario)
       oTabla.IdEstadoFacturacion = IIf(IsNull(oRecordset!IdEstadoFacturacion), 0, oRecordset!IdEstadoFacturacion)
       oTabla.ImporteExonerado = IIf(IsNull(oRecordset!ImporteExonerado), 0, oRecordset!ImporteExonerado)
       oTabla.idUsuarioExonera = IIf(IsNull(oRecordset!idUsuarioExonera), 0, oRecordset!idUsuarioExonera)
       oTabla.DNI = IIf(IsNull(oRecordset!DNI), "", oRecordset!DNI)
       oTabla.nombpaciente = IIf(IsNull(oRecordset!nombpaciente), "", oRecordset!nombpaciente)
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
       oTabla.Paquete = IIf(IsNull(oRecordset!Paquete), "", oRecordset!Paquete)
End Sub


