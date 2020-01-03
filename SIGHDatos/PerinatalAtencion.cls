VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PerinatalAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla PerinatalAtencion
'        Programado por: Garay M
'        Fecha: Agosto 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Const lnPercentilNull As Long = 0
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
'   Descripción:    Inserta un registro a la tabla PerinatalAtencion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoPerinatalAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PerinatalAtencionAgregar"
   Set oParameter = .CreateParameter("@idPerinatalAtencion", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idModulo", adInteger, adParamInput, 0, IIf(oTabla.idModulo = 0, Null, oTabla.idModulo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrafXedadEnMeses", adInteger, adParamInput, 0, IIf(oTabla.GrafXedadEnMeses = 0, Null, oTabla.GrafXedadEnMeses))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrafYpercentilTE", adInteger, adParamInput, 0, IIf(oTabla.GrafYpercentilTE = lnPercentilNull, 0, oTabla.GrafYpercentilTE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrafYpercentilPT", adInteger, adParamInput, 0, IIf(oTabla.GrafYpercentilPT = lnPercentilNull, 0, oTabla.GrafYpercentilPT))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrafYpercentilPE", adInteger, adParamInput, 0, IIf(oTabla.GrafYpercentilPE = lnPercentilNull, 0, oTabla.GrafYpercentilPE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrafYimc", adCurrency, adParamInput, 0, IIf(oTabla.GrafYimc = lnPercentilNull, Null, oTabla.GrafYimc))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaAtencion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaAtencion = 0, Null, oTabla.FechaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CredN", adInteger, adParamInput, 0, oTabla.credn)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.idPerinatalAtencion = .Parameters("@idPerinatalAtencion")
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
'   Descripción:    modifica un registro a la tabla PerinatalAtencion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoPerinatalAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PerinatalAtencionModificar"
   Set oParameter = .CreateParameter("@idPerinatalAtencion", adInteger, adParamInput, 0, IIf(oTabla.idPerinatalAtencion = 0, Null, oTabla.idPerinatalAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idModulo", adInteger, adParamInput, 0, IIf(oTabla.idModulo = 0, Null, oTabla.idModulo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrafXedadEnMeses", adInteger, adParamInput, 0, IIf(oTabla.GrafXedadEnMeses = 0, Null, oTabla.GrafXedadEnMeses))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrafYpercentilTE", adInteger, adParamInput, 0, IIf(oTabla.GrafYpercentilTE = lnPercentilNull, Null, oTabla.GrafYpercentilTE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrafYpercentilPT", adInteger, adParamInput, 0, IIf(oTabla.GrafYpercentilPT = lnPercentilNull, Null, oTabla.GrafYpercentilPT))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrafYpercentilPE", adInteger, adParamInput, 0, IIf(oTabla.GrafYpercentilPE = lnPercentilNull, Null, oTabla.GrafYpercentilPE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrafYimc", adCurrency, adParamInput, 0, IIf(oTabla.GrafYimc = lnPercentilNull, Null, oTabla.GrafYimc))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaAtencion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaAtencion = 0, Null, oTabla.FechaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CredN", adInteger, adParamInput, 0, oTabla.credn)
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
'   Descripción:    elimia o borra un registro a la tabla PerinatalAtencion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoPerinatalAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PerinatalAtencionEliminar"
   Set oParameter = .CreateParameter("@idPerinatalAtencion", adInteger, adParamInput, 0, oTabla.idPerinatalAtencion)
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

Sub ActualizaTabla(ByRef oTabla As DoPerinatalAtencion, oRecordset As Recordset)

       oTabla.idPerinatalAtencion = IIf(IsNull(oRecordset!idPerinatalAtencion), 0, oRecordset!idPerinatalAtencion)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.idModulo = IIf(IsNull(oRecordset!idModulo), 0, oRecordset!idModulo)
       oTabla.GrafXedadEnMeses = IIf(IsNull(oRecordset!GrafXedadEnMeses), 0, oRecordset!GrafXedadEnMeses)
       oTabla.GrafYpercentilTE = IIf(IsNull(oRecordset!GrafYpercentilTE), 0, oRecordset!GrafYpercentilTE)
       oTabla.GrafYpercentilPT = IIf(IsNull(oRecordset!GrafYpercentilPT), 0, oRecordset!GrafYpercentilPT)
       oTabla.GrafYpercentilPE = IIf(IsNull(oRecordset!GrafYpercentilPE), 0, oRecordset!GrafYpercentilPE)
       oTabla.FechaAtencion = IIf(IsNull(oRecordset!FechaAtencion), 0, oRecordset!FechaAtencion)
       oTabla.GrafYimc = IIf(IsNull(oRecordset!GrafYimc), 0, oRecordset!GrafYimc)
       oTabla.credn = IIf(IsNull(oRecordset!credn), 0, oRecordset!credn)
End Sub


'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla PerinatalAtencion
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoPerinatalAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "PerinatalAtencionSeleccionarPorId"
     Set oParameter = .CreateParameter("@idPerinatalAtencion", adInteger, adParamInput, 0, oTabla.idPerinatalAtencion)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
      ActualizaTabla oTabla, oRecordset
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoPerinatalAtencion)
   oTabla.idPerinatalAtencion = 0
   oTabla.IdPaciente = 0
   oTabla.idModulo = 0
   oTabla.GrafXedadEnMeses = 0
   oTabla.GrafYpercentilTE = lnPercentilNull
   oTabla.GrafYpercentilPT = lnPercentilNull
   oTabla.GrafYpercentilPE = lnPercentilNull
   oTabla.GrafYimc = lnPercentilNull
   oTabla.FechaAtencion = 0
End Sub

















Function SeleccionarPorIdAtencion(ByVal oTabla As DoPerinatalAtencion, lnIdAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorIdAtencion = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "PerinatalAtencionSeleccionarPorIdAtencion"
     Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, lnIdAtencion)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
    If Not (oRecordset.EOF And oRecordset.BOF) Then
       ActualizaTabla oTabla, oRecordset
   SeleccionarPorIdAtencion = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


Function SeleccionarPorIdPaciente(ByVal oTabla As DoPerinatalAtencion, lnIdPaciente As Long, ldFechaAtencion As Date) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "PerinatalAtencionXidPaciente"
        Set oParameter = .CreateParameter("@lnIdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       ActualizaTabla oTabla, oRecordset
       SeleccionarPorIdPaciente = True
       ms_MensajeError = ""
   End If
   oRecordset.Close
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
