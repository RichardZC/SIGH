VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Establecimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla Establecimientos
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
'   Descripción:    Inserta un registro a la tabla Establecimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOEstablecimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "EstablecimientosAgregar"
       Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, oTabla.IdEstablecimiento): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 6, IIf(oTabla.codigo = "", Null, oTabla.codigo)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 150, IIf(oTabla.nombre = "", Null, oTabla.nombre)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdDistrito", adInteger, adParamInput, 0, IIf(oTabla.IdDistrito = 0, Null, oTabla.IdDistrito)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdTipo", adInteger, adParamInput, 0, IIf(oTabla.IdTipo = 0, Null, oTabla.IdTipo)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
   End With
 
   Insertar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description ': MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla Establecimientos
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOEstablecimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "EstablecimientosModificar"
       Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamOutput, 0): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 6, IIf(oTabla.codigo = "", Null, oTabla.codigo)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 150, IIf(oTabla.nombre = "", Null, oTabla.nombre)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdDistrito", adInteger, adParamInput, 0, IIf(oTabla.IdDistrito = 0, Null, oTabla.IdDistrito)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdTipo", adInteger, adParamInput, 0, IIf(oTabla.IdTipo = 0, Null, oTabla.IdTipo)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla Establecimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOEstablecimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "EstablecimientosEliminar"
           Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla Establecimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOEstablecimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EstablecimientosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, oTabla.IdEstablecimiento): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdTipo = IIf(IsNull(oRecordset!IdTipo), 0, oRecordset!IdTipo)
       oTabla.IdDistrito = IIf(IsNull(oRecordset!IdDistrito), 0, oRecordset!IdDistrito)
       oTabla.nombre = IIf(IsNull(oRecordset!nombre), "", oRecordset!nombre)
       oTabla.codigo = IIf(IsNull(oRecordset!codigo), "", oRecordset!codigo)
       oTabla.IdEstablecimiento = IIf(IsNull(oRecordset!IdEstablecimiento), 0, oRecordset!IdEstablecimiento)
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



Function Filtrar(ByVal oTabla As DOEstablecimiento, lDepartamento As Long, _
                lProvincia As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sOrder As String, sWhere As String
    sSql = ""
    sOrder = ""
    sWhere = ""
    If oTabla.codigo <> "" Then
        'mgaray201503
         sWhere = sWhere + " Establecimientos.Codigo = '" + _
                            FormatoCodigoRENAES(oTabla.codigo, GALENHOS) + "' and "
    End If
    If oTabla.nombre <> "" Then
         sWhere = sWhere + " Establecimientos.Nombre like '%" + oTabla.nombre + "%' and "
    End If
    If oTabla.IdDistrito <> 0 Then
         sWhere = sWhere + " Establecimientos.IdDistrito = " & oTabla.IdDistrito & " and "
    End If

    'JVG - Adicion de Nivel maximo de Establecimiento
    If oTabla.IdTipo <> 0 Then
         sWhere = sWhere + " Establecimientos.IdTipo >= " & oTabla.IdTipo & " and "
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
        .CommandText = "EstablecimientosFiltrar"
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


Function SeleccionarPorCodigo(ByVal oDoEstablecimiento As DOEstablecimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sOrder As String, sWhere As String
    sSql = ""
    sOrder = ""
    sWhere = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "EstablecimientosXcodigo"
        'mgaray201503
        Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 6, _
                                    FormatoCodigoRENAES(oDoEstablecimiento.codigo, GALENHOS)): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If Not (oRecordset.EOF And oRecordset.BOF) Then
        oDoEstablecimiento.IdEstablecimiento = oRecordset!IdEstablecimiento
        SeleccionarPorCodigo = SeleccionarPorId(oDoEstablecimiento)
   Else
        SeleccionarPorCodigo = False
   End If
   oRecordset.Close
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'---------------------------------------------------------------------------------------
' Procedure : ObtenerDatosEstablecimientoPorIdUsuario
' Author    : JVALLE
' Date      : 07/05/2012 14:48
' Purpose   : Obtiene los datos del Establecimiento dependiendo del ID del usuario
'---------------------------------------------------------------------------------------
'
Public Function ObtenerDatosEstablecimientoPorIdUsuario(IdUsuario As Long) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sOrder As String, sWhere As String
    sSql = ""
    sOrder = ""
    sWhere = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "EstablecimientosPorIdUsuario"
        Set oParameter = .CreateParameter("@IdUsuario", adInteger, adParamInput, 0, IdUsuario): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerDatosEstablecimientoPorIdUsuario = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'---------------------------------------------------------------------------------------
' Procedure : ObtenerDatosEstablecimiento
' Author    : JVALLE
' Date      : 07/05/2012 14:48
' Purpose   : Obtiene los datos del Establecimiento dependiendo del ID del usuario
'---------------------------------------------------------------------------------------
'
Public Function ObtenerDatosEstablecimientoPorId(oTabla As DOEstablecimiento) As DOEstablecimiento

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sOrder As String, sWhere As String
Dim oDoEstablecimiento As New DOEstablecimiento
    sSql = ""
    sOrder = ""
    sWhere = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "EstablecimientosXid"
        Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, oTabla.IdEstablecimiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount <> 0 Then
      While Not oRecordset.EOF
        oDoEstablecimiento.IdEstablecimiento = oRecordset!IdEstablecimiento
        oDoEstablecimiento.codigo = oRecordset!codigo
        oDoEstablecimiento.nombre = oRecordset!nombre
        oRecordset.MoveNext
      Wend
   End If
    
   Set ObtenerDatosEstablecimientoPorId = oDoEstablecimiento
   Set oCommand = Nothing
   Set oRecordset = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'mgaray20141012
Function SeleccionarTodos() As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sOrder As String, sWhere As String
    sSql = ""
    sWhere = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "EstablecimientosSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarTodos = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
