VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EstablecimientosNoMinsa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla EstablecimientosNoMinsa
'        Programado por: Castro W
'        Fecha: Agosto 2006
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
'   Descripción:    Inserta un registro a la tabla EstablecimientosNoMinsa
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOEstablecimientoNoMinsa) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "EstablecimientosNoMinsaAgregar"
           Set oParameter = .CreateParameter("@IdEstablecimientoNoMinsa", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 150, IIf(oTabla.nombre = "", Null, oTabla.nombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoSubsector", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSubsector = 0, Null, oTabla.IdTipoSubsector)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDistrito", adInteger, adParamInput, 0, IIf(oTabla.IdDistrito = 0, Null, oTabla.IdDistrito)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 10, IIf(oTabla.codigo = "", Null, oTabla.codigo)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdEstablecimientoNoMinsa = .Parameters("@IdEstablecimientoNoMinsa")
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
'   Descripción:    modifica un registro a la tabla EstablecimientosNoMinsa
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOEstablecimientoNoMinsa) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "EstablecimientosNoMinsaModificar"
           Set oParameter = .CreateParameter("@IdEstablecimientoNoMinsa", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoNoMinsa = 0, Null, oTabla.IdEstablecimientoNoMinsa)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 150, IIf(oTabla.nombre = "", Null, oTabla.nombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoSubsector", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSubsector = 0, Null, oTabla.IdTipoSubsector)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDistrito", adInteger, adParamInput, 0, IIf(oTabla.IdDistrito = 0, Null, oTabla.IdDistrito)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 10, IIf(oTabla.codigo = "", Null, oTabla.codigo)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla EstablecimientosNoMinsa
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOEstablecimientoNoMinsa) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "EstablecimientosNoMinsaEliminar"
           Set oParameter = .CreateParameter("@IdEstablecimientoNoMinsa", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoNoMinsa = 0, Null, oTabla.IdEstablecimientoNoMinsa)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla EstablecimientosNoMinsa
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOEstablecimientoNoMinsa) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EstablecimientosNoMinsaSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdEstablecimientoNoMinsa", adInteger, adParamInput, 0, oTabla.IdEstablecimientoNoMinsa): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdEstablecimientoNoMinsa = IIf(IsNull(oRecordset!IdEstablecimientoNoMinsa), 0, oRecordset!IdEstablecimientoNoMinsa)
       oTabla.nombre = IIf(IsNull(oRecordset!nombre), "", oRecordset!nombre)
       oTabla.IdTipoSubsector = IIf(IsNull(oRecordset!IdTipoSubsector), 0, oRecordset!IdTipoSubsector)
       oTabla.IdDistrito = IIf(IsNull(oRecordset!IdDistrito), 0, oRecordset!IdDistrito)
       oTabla.codigo = IIf(IsNull(oRecordset!codigo), "", oRecordset!codigo)
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




Function SeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EstablecimientosNoMinsaSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
Function Filtrar(ByVal oTabla As DOEstablecimientoNoMinsa, lDepartamento As Long, lProvincia As Long) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    If oTabla.nombre <> "" Then
         sWhere = sWhere + " EstablecimientosNoMinsa.Nombre like '%" + oTabla.nombre + "%' and "
    End If
    If oTabla.IdTipoSubsector <> 0 Then
         sWhere = sWhere + " EstablecimientosNoMinsa.IdTipoSubsector like '" + oTabla.IdTipoSubsector + "%' and "
    End If
    If oTabla.IdDistrito <> 0 Then
         sWhere = sWhere + " EstablecimientosNoMinsa.IdDistrito = " & oTabla.IdDistrito & " and "
    End If
    If lDepartamento <> 0 Then
         sWhere = sWhere + " Departamentos.IdDepartamento = " & lDepartamento & " and "
    End If
    If lProvincia <> 0 Then
         sWhere = sWhere + " Provincias.IdProvincia = " & lProvincia & " and "
    End If
    If sWhere <> "" Then
         sSql = sSql + " where " + Left(sWhere, Len(sWhere) - 4)
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "EstablecimientosNoMinsaFiltrar"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set Filtrar = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
