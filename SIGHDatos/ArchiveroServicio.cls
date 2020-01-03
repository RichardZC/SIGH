VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArchiveroServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla ArchiveroServicio
'        Programado por: Castro W
'        Fecha: Diciembre 2005
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
'   Descripción:    Inserta un registro a la tabla ArchiveroServicio
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOArchiveroServicio) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ArchiveroServicioAgregar"
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleado = 0, Null, oTabla.IdEmpleado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdArchivero", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsConsultorioAsignado", adBoolean, adParamInput, 0, oTabla.EsConsultorioAsignado): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdArchivero = .Parameters("@IdArchivero")
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
'   Descripción:    modifica un registro a la tabla ArchiveroServicio
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOArchiveroServicio) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ArchiveroServicioModificar"
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleado = 0, Null, oTabla.IdEmpleado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdArchivero", adInteger, adParamInput, 0, IIf(oTabla.IdArchivero = 0, Null, oTabla.IdArchivero)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsConsultorioAsignado", adBoolean, adParamInput, 0, oTabla.EsConsultorioAsignado): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla ArchiveroServicio
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOArchiveroServicio) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ArchiveroServicioEliminar"
           Set oParameter = .CreateParameter("@IdArchivero", adInteger, adParamInput, 0, IIf(oTabla.IdArchivero = 0, Null, oTabla.IdArchivero)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla ArchiveroServicio
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOArchiveroServicio) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ArchiveroServicioSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdArchivero", adInteger, adParamInput, 0, oTabla.IdArchivero): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdServicio = IIf(IsNull(oRecordset!IdServicio), 0, oRecordset!IdServicio)
       oTabla.IdEmpleado = IIf(IsNull(oRecordset!IdEmpleado), 0, oRecordset!IdEmpleado)
       oTabla.IdArchivero = IIf(IsNull(oRecordset!IdArchivero), 0, oRecordset!IdArchivero)
       oTabla.EsConsultorioAsignado = IIf(IsNull(oRecordset!EsConsultorioAsignado), 0, oRecordset!EsConsultorioAsignado)
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

Function AgregarVarios(oArchiveroServicio As Collection) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oDOArchivero As New DOArchiveroServicio
   AgregarVarios = False
   
   If Not oArchiveroServicio Is Nothing Then
        For Each oDOArchivero In oArchiveroServicio
            If Not Insertar(oDOArchivero) Then
                Exit Function
            End If
         Next
    End If
    
   AgregarVarios = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function ModificarVarios(oArchiveroServicio As Collection) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oDOArchivero As New DOArchiveroServicio
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ArchiveroServicioEliminarXIdEmpleado"
        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, oArchiveroServicio.Item(1).IdEmpleado): .Parameters.Append oParameter
        .Execute
   End With
   If Not oArchiveroServicio Is Nothing Then
        For Each oDOArchivero In oArchiveroServicio
            If Not Insertar(oDOArchivero) Then
                Exit Function
            End If
         Next
   End If
   
   Set oRecordset = Nothing
   Set oCommand = Nothing
   ModificarVarios = True
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function EliminarVarios(oArchiveroServicio As Collection) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ArchiveroServicioEliminarXIdEmpleado"
        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, oArchiveroServicio.Item(1).IdEmpleado): .Parameters.Append oParameter
        .Execute
   End With
   Set oRecordset = Nothing
   Set oCommand = Nothing
   EliminarVarios = True
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function Filtrar(ByVal oDOEmpleado As dOEmpleado, lbEsConsultorioAsignado As Boolean) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sWhere As String
    sWhere = ""
    If oDOEmpleado.ApellidoPaterno <> "" Then
        sWhere = sWhere + " Empleados.ApellidoPaterno like '" + oDOEmpleado.ApellidoPaterno + "%' and "
    End If
    If oDOEmpleado.ApellidoMaterno <> "" Then
         sWhere = sWhere + " Empleados.ApellidoMaterno like '" + oDOEmpleado.ApellidoMaterno + "%' and "
    End If
    If oDOEmpleado.Nombres <> "" Then
         sWhere = sWhere + " Empleados.Nombres like '" + oDOEmpleado.Nombres + "%' and "
    End If
    If oDOEmpleado.CodigoPlanilla <> "" Then
         sWhere = sWhere + " Empleados.CodigoPlanilla ='" & oDOEmpleado.CodigoPlanilla & "' and "
    End If
    sWhere = sWhere + " ArchiveroServicio.EsConsultorioAsignado=" & IIf(lbEsConsultorioAsignado = True, "1", "0") & " and "
    
    If sWhere <> "" Then
       sWhere = Left(sWhere, Len(sWhere) - 4)
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ArchiveroServicioFiltro"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sWhere): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set Filtrar = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function FiltrarPorEmpleado(lIdEmpleado As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ArchiveroServicioXidEmpleado"
        Set oParameter = .CreateParameter("@lIdEmpleado", adInteger, adParamInput, 0, lIdEmpleado): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FiltrarPorEmpleado = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


