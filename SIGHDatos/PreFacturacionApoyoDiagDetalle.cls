VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionApoyoDiagDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: AtencionApoyoDiagDetalle
'        Autor: William Castro Grijalva
'        Fecha: 14/06/2005 10:49:25 p.m.
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
'   Descripción:    Inserta un registro a la tabla AtencionApoyoDiagDetalle
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOAtencionApoyoDiagDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionApoyoDiagDetalleAgregar"
           Set oParameter = .CreateParameter("@IdFacturacionServicio", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionServicio = 0, Null, oTabla.IdFacturacionServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraResultado", adChar, adParamInput, 5, IIf(oTabla.HoraResultado = "", Null, oTabla.HoraResultado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaResultado", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaResultado = 0, Null, oTabla.FechaResultado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioRealiza", adInteger, adParamInput, 0, IIf(oTabla.IdServicioRealiza = 0, Null, oTabla.IdServicioRealiza)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProcedimiento", adInteger, adParamInput, 0, IIf(oTabla.IdProcedimiento = 0, Null, oTabla.IdProcedimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionApoyoDx", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionApoyoDx = 0, Null, oTabla.IdAtencionApoyoDx)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionApoyoDetalle", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdAtencionApoyoDetalle = .Parameters("@IdAtencionApoyoDetalle")
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
'   Descripción:    modifica un registro a la tabla AtencionApoyoDiagDetalle
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtencionApoyoDiagDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionApoyoDiagDetalleModificar"
           Set oParameter = .CreateParameter("@IdFacturacionServicio", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionServicio = 0, Null, oTabla.IdFacturacionServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraResultado", adChar, adParamInput, 5, IIf(oTabla.HoraResultado = "", Null, oTabla.HoraResultado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaResultado", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaResultado = 0, Null, oTabla.FechaResultado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioRealiza", adInteger, adParamInput, 0, IIf(oTabla.IdServicioRealiza = 0, Null, oTabla.IdServicioRealiza)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProcedimiento", adInteger, adParamInput, 0, IIf(oTabla.IdProcedimiento = 0, Null, oTabla.IdProcedimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionApoyoDx", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionApoyoDx = 0, Null, oTabla.IdAtencionApoyoDx)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionApoyoDetalle", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionApoyoDetalle = 0, Null, oTabla.IdAtencionApoyoDetalle)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla AtencionApoyoDiagDetalle
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtencionApoyoDiagDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionApoyoDiagDetalleEliminar"
           Set oParameter = .CreateParameter("@IdAtencionApoyoDetalle", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionApoyoDetalle = 0, Null, oTabla.IdAtencionApoyoDetalle)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla AtencionApoyoDiagDetalle
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtencionApoyoDiagDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtencionApoyoDiagDetalleSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdAtencionApoyoDetalle", adInteger, adParamInput, 0, oTabla.IdAtencionApoyoDetalle): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdFacturacionServicio = IIf(IsNull(oRecordset!IdFacturacionServicio), 0, oRecordset!IdFacturacionServicio)
       oTabla.HoraResultado = IIf(IsNull(oRecordset!HoraResultado), "", oRecordset!HoraResultado)
       oTabla.FechaResultado = IIf(IsNull(oRecordset!FechaResultado), 0, oRecordset!FechaResultado)
       oTabla.IdServicioRealiza = IIf(IsNull(oRecordset!IdServicioRealiza), 0, oRecordset!IdServicioRealiza)
       oTabla.IdProcedimiento = IIf(IsNull(oRecordset!IdProcedimiento), 0, oRecordset!IdProcedimiento)
       oTabla.IdAtencionApoyoDx = IIf(IsNull(oRecordset!IdAtencionApoyoDx), 0, oRecordset!IdAtencionApoyoDx)
       oTabla.IdAtencionApoyoDetalle = IIf(IsNull(oRecordset!IdAtencionApoyoDetalle), 0, oRecordset!IdAtencionApoyoDetalle)
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



Function SeleccionarPorIdAtencionApoyoDx(IdPreFacturacionApoyo As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String
Dim sWhere As String

        Set SeleccionarPorIdAtencionApoyoDx = Nothing
     
        sSQL = "select AtencionApoyoDiagDetalle.IdAtencionApoyoDetalle, " + _
        "AtencionApoyoDiagDetalle.IdProcedimiento, " + _
        " Procedimientos.CodigoCPT2004 as CodigoCPT , " + _
        " Procedimientos.Descripcion, " + _
        " AtencionApoyoDiagDetalle.IdServicioRealiza, " + _
        " Servicios.Nombre as NombreServicio, " + _
        " AtencionApoyoDiagDetalle.FechaResultado, " + _
        " AtencionApoyoDiagDetalle.HoraResultado, " + _
        " AtencionApoyoDiagDetalle.IdFacturacionServicio " + _
        " From " + _
        " (AtencionApoyoDiagDetalle " + _
        " left join Servicios on AtencionApoyoDiagDetalle.IdServicioRealiza = Servicios.IdServicio) " + _
        " left join Procedimientos on AtencionApoyoDiagDetalle.IdProcedimiento =  Procedimientos.IdProcedimiento " + _
        " where IdAtencionApoyoDx = " & IdPreFacturacionApoyo & _
        " order by FechaResultado desc"
                
        With oCommand
           .CommandType = adCmdText
           Set .ActiveConnection = mo_Conexion
           .CommandText = sSQL
           Set oRecordset = .Execute
           Set oRecordset.ActiveConnection = Nothing
        End With
    
        Set SeleccionarPorIdAtencionApoyoDx = oRecordset
    
        ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function

Function EliminarPorIdPreFacturacionApoyoDiagnostico(lIdPreFacturacionApoyo As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSQL As String
 
   EliminarPorIdPreFacturacionApoyoDiagnostico = False
   
   sSQL = "Delete from AtencionApoyoDiagDetalle where IdAtencionApoyoDx = " & lIdPreFacturacionApoyo
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = sSQL
       .Execute
   End With
   
   EliminarPorIdPreFacturacionApoyoDiagnostico = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function





