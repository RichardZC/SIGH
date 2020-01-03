VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionesExamenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: AtencionesExamenes
'        Autor: William Castro Grijalva
'        Fecha: 28/10/2004 09:43:17 p.m.
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
'   Descripción:    Inserta un registro a la tabla AtencionesExamenes
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOAtencionExamen) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesExamenesAgregar"
           Set oParameter = .CreateParameter("@IdDetalleProducto", adInteger, adParamInput, 0, IIf(oTabla.IdDetalleProducto = 0, Null, oTabla.IdDetalleProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraOrden", adChar, adParamInput, 5, IIf(oTabla.HoraOrden = "", Null, oTabla.HoraOrden)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioOrdena", adInteger, adParamInput, 0, IIf(oTabla.IdServicioOrdena = 0, Null, oTabla.IdServicioOrdena)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdExamen", adInteger, adParamInput, 0, IIf(oTabla.IdExamen = 0, Null, oTabla.IdExamen)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaOrden", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaOrden = 0, Null, oTabla.FechaOrden)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@OrdenNro", adVarChar, adParamInput, 10, IIf(oTabla.OrdenNro = "", Null, oTabla.OrdenNro)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoOrdena", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoOrdena = 0, Null, oTabla.IdMedicoOrdena)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraResultado", adChar, adParamInput, 5, IIf(oTabla.HoraResultado = "", Null, oTabla.HoraResultado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaResultado", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaResultado = 0, Null, oTabla.FechaResultado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionExamen", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdAtencionExamen = .Parameters("@IdAtencionExamen")
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
'   Descripción:    modifica un registro a la tabla AtencionesExamenes
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtencionExamen) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesExamenesModificar"
           Set oParameter = .CreateParameter("@IdDetalleProducto", adInteger, adParamInput, 0, IIf(oTabla.IdDetalleProducto = 0, Null, oTabla.IdDetalleProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraOrden", adChar, adParamInput, 5, IIf(oTabla.HoraOrden = "", Null, oTabla.HoraOrden)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioOrdena", adInteger, adParamInput, 0, IIf(oTabla.IdServicioOrdena = 0, Null, oTabla.IdServicioOrdena)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdExamen", adInteger, adParamInput, 0, IIf(oTabla.IdExamen = 0, Null, oTabla.IdExamen)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaOrden", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaOrden = 0, Null, oTabla.FechaOrden)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@OrdenNro", adVarChar, adParamInput, 10, IIf(oTabla.OrdenNro = "", Null, oTabla.OrdenNro)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoOrdena", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoOrdena = 0, Null, oTabla.IdMedicoOrdena)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraResultado", adChar, adParamInput, 5, IIf(oTabla.HoraResultado = "", Null, oTabla.HoraResultado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaResultado", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaResultado = 0, Null, oTabla.FechaResultado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionExamen", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionExamen = 0, Null, oTabla.IdAtencionExamen)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla AtencionesExamenes
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtencionExamen) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesExamenesEliminar"
           Set oParameter = .CreateParameter("@IdAtencionExamen", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionExamen = 0, Null, oTabla.IdAtencionExamen)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla AtencionesExamenes
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtencionExamen) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtencionesExamenesSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdAtencionExamen", adInteger, adParamInput, 0, oTabla.IdAtencionExamen): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdDetalleProducto = IIf(IsNull(oRecordset!IdDetalleProducto), 0, oRecordset!IdDetalleProducto)
       oTabla.HoraOrden = IIf(IsNull(oRecordset!HoraOrden), "", oRecordset!HoraOrden)
       oTabla.IdServicioOrdena = IIf(IsNull(oRecordset!IdServicioOrdena), 0, oRecordset!IdServicioOrdena)
       oTabla.IdExamen = IIf(IsNull(oRecordset!IdExamen), 0, oRecordset!IdExamen)
       oTabla.FechaOrden = IIf(IsNull(oRecordset!FechaOrden), 0, oRecordset!FechaOrden)
       oTabla.OrdenNro = IIf(IsNull(oRecordset!OrdenNro), 0, oRecordset!OrdenNro)
       oTabla.IdMedicoOrdena = IIf(IsNull(oRecordset!IdMedicoOrdena), 0, oRecordset!IdMedicoOrdena)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.HoraResultado = IIf(IsNull(oRecordset!HoraResultado), "", oRecordset!HoraResultado)
       oTabla.FechaResultado = IIf(IsNull(oRecordset!FechaResultado), 0, oRecordset!FechaResultado)
       oTabla.IdAtencionExamen = IIf(IsNull(oRecordset!IdAtencionExamen), 0, oRecordset!IdAtencionExamen)
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


Function SeleccionarPorCuentaAtencion(lIdCuentaAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorCuentaAtencion = Nothing
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtencionesExamenesSeleccionarPorCuentaAtencion"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 4, lIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorCuentaAtencion = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function ActualizarExamenesAtencion(oExamenes As Collection, lIdCuentaAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSQL As String
Dim oDOAtencionExamen As New DOAtencionExamen
 
   ActualizarExamenesAtencion = False
   
   'elimina los complicaciones
   sSQL = "delete from aPOYO where IdCuentaAtencion = " & lIdCuentaAtencion
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = sSQL
       .Execute
   End With
   
   If Not oExamenes Is Nothing Then
        For Each oDOAtencionExamen In oExamenes
            oDOAtencionExamen.IdCuentaAtencion = lIdCuentaAtencion
            If Not Insertar(oDOAtencionExamen) Then
                Exit Function
            End If
         Next
    End If
    
   ActualizarExamenesAtencion = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function EliminarExamenesPorCuentaAtencion(lIdCuentaAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSQL As String
 
   EliminarExamenesPorCuentaAtencion = False
   
   'elimina los diagnosticos
   sSQL = "delete from AtencionesExamenes where IdCuentaAtencion = " & lIdCuentaAtencion
   With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = sSQL
       .Execute
   End With
   
   EliminarExamenesPorCuentaAtencion = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


