VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla Parametros
'        Programado por: Benavides M
'        Fecha: Febrero 2014
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
'   Descripción:    Inserta un registro a la tabla Parametros
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOPArametro) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ParametrosAgregar"
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 150, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 20, IIf(oTabla.codigo = "", Null, oTabla.codigo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Tipo", adVarChar, adParamInput, 20, IIf(oTabla.Tipo = "", Null, oTabla.Tipo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdParametro", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ValorFloat", adDouble, adParamInput, 0, IIf(oTabla.ValorFloat = 0, Null, oTabla.ValorFloat)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ValorInt", adInteger, adParamInput, 0, IIf(oTabla.ValorInt = 0, Null, oTabla.ValorInt)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ValorTexto", adVarChar, adParamInput, 150, IIf(oTabla.ValorTexto = "", Null, oTabla.ValorTexto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@grupo", adVarChar, adParamInput, 150, IIf(oTabla.Grupo = "", Null, oTabla.Grupo)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdParametro = .Parameters("@IdParametro")
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
'   Descripción:    modifica un registro a la tabla Parametros
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOPArametro) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ParametrosModificar"
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 150, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 20, IIf(oTabla.codigo = "", Null, oTabla.codigo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Tipo", adVarChar, adParamInput, 20, IIf(oTabla.Tipo = "", Null, oTabla.Tipo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdParametro", adInteger, adParamInput, 0, IIf(oTabla.IdParametro = 0, Null, oTabla.IdParametro)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ValorFloat", adDouble, adParamInput, 0, oTabla.ValorFloat): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ValorInt", adInteger, adParamInput, 0, oTabla.ValorInt): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ValorTexto", adVarChar, adParamInput, 150, oTabla.ValorTexto): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@grupo", adVarChar, adParamInput, 150, IIf(oTabla.Grupo = "", Null, oTabla.Grupo)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla Parametros
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOPArametro) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ParametrosEliminar"
           Set oParameter = .CreateParameter("@IdParametro", adInteger, adParamInput, 0, IIf(oTabla.IdParametro = 0, Null, oTabla.IdParametro)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla Parametros
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOPArametro) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ParametrosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdParametro", adInteger, adParamInput, 0, oTabla.IdParametro): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
       oTabla.codigo = IIf(IsNull(oRecordset!codigo), "", oRecordset!codigo)
       oTabla.Tipo = IIf(IsNull(oRecordset!Tipo), "", oRecordset!Tipo)
       oTabla.IdParametro = IIf(IsNull(oRecordset!IdParametro), 0, oRecordset!IdParametro)
       oTabla.ValorFloat = IIf(IsNull(oRecordset!ValorFloat), 0, oRecordset!ValorFloat)
       oTabla.ValorInt = IIf(IsNull(oRecordset!ValorInt), 0, oRecordset!ValorInt)
       oTabla.ValorTexto = IIf(IsNull(oRecordset!ValorTexto), "", oRecordset!ValorTexto)
       oTabla.Grupo = IIf(IsNull(oRecordset!Grupo), "", oRecordset!Grupo)
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

Function SeleccionarComparadorFechas() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarComparadorFechas = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ParametrosSeleccionarComparadorFechas"
     Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarComparadorFechas = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarTipo(sTipo As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarTipo = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ParametrosSeleccionarPorTipo"
        Set oParameter = .CreateParameter("@Tipo", adVarChar, adParamInput, 20, IIf(sTipo = "", Null, sTipo)): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarTipo = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function ObtenerCadenaConexion() As String
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   ObtenerCadenaConexion = ""
   
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ParametrosSeleccionarCadenaConexion"
     Set oRecordset = .Execute
   End With
   Set oRecordset.ActiveConnection = Nothing
   If Not (oRecordset.EOF And oRecordset.BOF) Then
        ObtenerCadenaConexion = oRecordset!ValorTexto
   End If
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
Function ObtenerCadenaConexionShape() As String
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   ObtenerCadenaConexionShape = ""
   
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ParametrosSeleccionarCadenaConexionShape"
     Set oRecordset = .Execute
   End With
   Set oRecordset.ActiveConnection = Nothing
   If Not (oRecordset.EOF And oRecordset.BOF) Then
        ObtenerCadenaConexionShape = oRecordset!ValorTexto
   End If
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


'------------------------------------------------------------------------------------
'   Función:        GenerarNroHistoria
'   Descripción:    Genera el Nro De Historia Clinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function GenerarIdGrupoMovimiento() As Long
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   GenerarIdGrupoMovimiento = 0
   With oCommand
       .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "ParametrosGenerarNroAgrupadorMovimientos"
        Set oParameter = .CreateParameter("@IdGrupoMovimiento", adInteger, adParamOutput, 0): .Parameters.Append oParameter
        .Execute
        GenerarIdGrupoMovimiento = .Parameters("@IdGrupoMovimiento")
   End With
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
    GenerarIdGrupoMovimiento = -1
    ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarIdServicioArchivoClinico() As Long
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ParametrosSeleccionarIdServicioArchivoClinico"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   SeleccionarIdServicioArchivoClinico = 0
   If Not (oRecordset.BOF And oRecordset.EOF) Then
        SeleccionarIdServicioArchivoClinico = Val(oRecordset!ValorInt)
   Else
        MsgBox "Debe configurar el código del ARCHIVO en la tabla de parámetros", vbInformation, "Error en la interface de acceso a datos"
   End If
   oRecordset.Close
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarTiposConsulta() As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ParametrosSeleccionarTiposConsulta"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarTiposConsulta = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function SeleccionarTiposConsultaPorEspecialidad(lIdEspecialidad As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarTiposConsultaPorEspecialidad = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "FactCatalogoServiciosSeleccionarTipoConsulta"
        Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, , lIdEspecialidad): .Parameters.Append oParameter
        Set oRecordset = .Execute
    End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarTiposConsultaPorEspecialidad = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
   SIGHENTIDADES.LogError ms_MensajeError
   
Exit Function

End Function

Function ObtenerCodigoDeNuevoCarnet() As Long
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ParametrosObtenerCodigoDeNuevoCarnet"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   ObtenerCodigoDeNuevoCarnet = 0
   If Not (oRecordset.BOF And oRecordset.EOF) Then
        ObtenerCodigoDeNuevoCarnet = Val(oRecordset!ValorInt)
   End If
   oRecordset.Close
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ObtenerCodigoDeDuplicadoCarnet() As Long
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ParametrosObtenerCodigoDeDuplicadoCarnet"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   ObtenerCodigoDeDuplicadoCarnet = 0
   If Not (oRecordset.BOF And oRecordset.EOF) Then
        ObtenerCodigoDeDuplicadoCarnet = Val(oRecordset!ValorInt)
   End If
   oRecordset.Close
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ObtenerCodigoDeFolder() As Long
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ParametrosObtenerCodigoDeFolder"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   ObtenerCodigoDeFolder = 0
   If Not (oRecordset.BOF And oRecordset.EOF) Then
        ObtenerCodigoDeFolder = Val(oRecordset!ValorInt)
   End If
   oRecordset.Close
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



Function SeleccionarTiposConsultaInterconsulta() As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ParametrosSeleccionarTiposConsultaInterconsulta"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarTiposConsultaInterconsulta = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ObtenerCodigoDeConsultaDeEmergencia() As Long
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ParametrosObtenerCodigoDeConsultaDeEmergencia"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   ObtenerCodigoDeConsultaDeEmergencia = 0
   If Not (oRecordset.BOF And oRecordset.EOF) Then
        ObtenerCodigoDeConsultaDeEmergencia = Val(oRecordset!ValorInt)
   End If
   oRecordset.Close
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ObtenerValorIntPorTipoYCodigo(sTipo As String, sCodigo As String) As Long
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ParametrosObtenerValorIntPorTipoYCodigo"
        Set oParameter = .CreateParameter("@sTipo", adVarChar, adParamInput, 20, sTipo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@sCodigo", adVarChar, adParamInput, 20, sCodigo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If Not (oRecordset.BOF And oRecordset.EOF) Then
        ObtenerValorIntPorTipoYCodigo = Val(oRecordset!ValorInt)
   Else
       MsgBox "Falta Parametro, tipo: " & sTipo & ",  codigo: " & sCodigo
    
   End If
   oRecordset.Close
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'***************daniel barrantes**************
'***************Devuelve Dias de un Paciente en Hospitalizacion o Emergencia
'***************
Function DiasDelPacienteEnHospitalizacionEmergencia(ldFechaIngreso As Date, lcHoraIngreso As String, _
                                                    ldFechaAlta As Date, lcHoraAlta As String, _
                                                    lcHoraEstanciaMax As String, _
                                                    Optional lbSinConsiderarHORAS As Boolean) As Integer
        Dim lnDiasEstancia As Integer
        Dim lnDiasEstanciaEnDecimales As Double
        If lbSinConsiderarHORAS = True Then
            lnDiasEstancia = DateDiff("d", ldFechaIngreso, Format(ldFechaAlta, "dd/mm/yyyy"))
            lnDiasEstanciaEnDecimales = lnDiasEstancia
        Else
            lnDiasEstanciaEnDecimales = DateDiff("h", ldFechaIngreso & " " & lcHoraIngreso, Format(ldFechaAlta, "dd/mm/yyyy") & " " & lcHoraAlta) / 24
            lnDiasEstancia = DateDiff("d", ldFechaIngreso & " " & lcHoraIngreso, Format(ldFechaAlta, "dd/mm/yyyy") & " " & lcHoraAlta)
        End If
        If lnDiasEstancia = 0 Then
            lnDiasEstancia = 1      'Estuvo un solo dia
        ElseIf lcHoraAlta > lcHoraEstanciaMax Then   '   'Paso de la 1pm se cuenta 1 dias mas
            lnDiasEstancia = lnDiasEstancia + 1
        ElseIf lnDiasEstanciaEnDecimales - lnDiasEstancia > 0 Then
            lnDiasEstancia = lnDiasEstancia + 1
        End If
        DiasDelPacienteEnHospitalizacionEmergencia = lnDiasEstancia
End Function


Function HorasDelPacienteEnHospitalizacionEmergencia(ldFechaIngreso As Date, lcHoraIngreso As String, ldFechaAlta As Date, lcHoraAlta As String) As Integer
         HorasDelPacienteEnHospitalizacionEmergencia = DateDiff("h", CDate(ldFechaIngreso & " " & lcHoraIngreso), CDate(ldFechaAlta & " " & lcHoraAlta))
End Function



'***************daniel barrantes**************
'***************Retorna un REGISTRO de la Tabla Parametro
'***************segun ID
Function SeleccionaFilaParametro(lnIdParametro As Long) As String
On Error GoTo ManejadorDeError
Dim vSeleccionaFilaParametro As String
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHENTIDADES.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ParametrosSeleccionarPorId"
        Set oParameter = .CreateParameter("@lnIdParametro", adInteger, adParamInput, 0, lnIdParametro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount > 0 Then
'       SeleccionaFilaParametro = Trim(oRecordset.Fields!ValorTexto)
       vSeleccionaFilaParametro = Trim(oRecordset.Fields!ValorTexto)
        Select Case vSeleccionaFilaParametro
            Case "s", "n"
                 SeleccionaFilaParametro = UCase(vSeleccionaFilaParametro)
            Case Else
                 SeleccionaFilaParametro = vSeleccionaFilaParametro
        End Select
   Else
       GoTo ManejadorDeError
   End If
   oRecordset.Close
   oConexion.Close
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox "Parametro: " & lnIdParametro & Chr(13) & Err.Description
End Function


Function RetornaFechaServidorSQLserver() As Date
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHENTIDADES.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RetornaFechaServidorSQL"
        Set oRecordset = .Execute
   End With
   If oRecordset.RecordCount > 0 Then
       RetornaFechaServidorSQLserver = Trim(oRecordset.Fields!FechaHoraSQL)
   Else
       GoTo ManejadorDeError
   End If
   oRecordset.Close
   oConexion.Close
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'***************daniel barrantes**************
'***************Retorna solo la fecha del Servidor en formato DD/MM/YYYY
'***************para Grabar/Mostrar Datos con esa Fecha/Hora
Function RetornaFechaServidorSQL() As String
    RetornaFechaServidorSQL = Format(RetornaFechaServidorSQLserver, SIGHENTIDADES.DevuelveFechaSoloFormato_DMY)
End Function
'***************daniel barrantes**************
'***************Retorna solo la hora del Servidor en formato HH:MM
'***************para Grabar/Mostrar Datos con esa Fecha/Hora
Function RetornaHoraServidorSQL() As String

    Dim lcRetornaFechaServidorSQL As String
    lcRetornaFechaServidorSQL = Format(RetornaFechaServidorSQLserver, SIGHENTIDADES.DevuelveHoraSoloFormato_HM)
    RetornaHoraServidorSQL = lcRetornaFechaServidorSQL
End Function

'--------------Adams Bonilla Magallanes
Function RetornaHoraServidorSQL1() As String

    RetornaHoraServidorSQL1 = Format(RetornaFechaServidorSQLserver, SIGHENTIDADES.DevuelveHoraSoloFormato_HMS)
    
End Function

Function RetornaNombreDeServidor() As String

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHENTIDADES.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RetornaNombreDeServidor"
        Set oRecordset = .Execute
   End With
   If oRecordset.RecordCount > 0 Then
       RetornaNombreDeServidor = Trim(oRecordset.Fields!nombre)
   Else
       GoTo ManejadorDeError
   End If
   oRecordset.Close
   oConexion.Close
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
'---------------

'***************daniel barrantes**************
'***************Retorna la fecha y Hora del Servidor en formato DD/MM/YYYY HH:MM
'***************para Grabar/Mostrar Datos con esa Fecha/Hora
Function RetornaFechaHoraServidorSQL() As String
    RetornaFechaHoraServidorSQL = Format(RetornaFechaServidorSQLserver, SIGHENTIDADES.DevuelveFechaSoloFormato_DMY_HM)
End Function

'***************daniel barrantes**************
'***************Retorna codigos y nombres del UBIGEO DEL HOSPITAL en un RECORDSET
Function DevuelveUbigeoDetalladoDelHospital() As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHENTIDADES.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelveUbigeoDetalladoDelHospital"
        Set oParameter = .CreateParameter("@SeleccionaFilaParametro208", adVarChar, adParamInput, 150, SeleccionaFilaParametro(208)): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set DevuelveUbigeoDetalladoDelHospital = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function RetornaLoginUsuario(idCodigoUsuario As Long) As String
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHENTIDADES.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosRetornaLoginUsuario"
        Set oParameter = .CreateParameter("@idCodigoUsuario", adInteger, adParamInput, 0, idCodigoUsuario): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   RetornaLoginUsuario = oRecordset.Fields!Usuario
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function RetornaFechaServidorSQL_AAMMDD() As String
    RetornaFechaServidorSQL_AAMMDD = Format(RetornaFechaServidorSQLserver, SIGHENTIDADES.DevuelveFechaSoloFormato_AAMMDD)
End Function


Function RetornaVersionServidorSQLserver() As String
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHENTIDADES.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RetornaVersionDeServidor"
        Set oRecordset = .Execute
   End With
   RetornaVersionServidorSQLserver = ""
   If oRecordset.RecordCount > 0 Then
       RetornaVersionServidorSQLserver = Trim(oRecordset.Fields!VersionServidor)
   End If
   oRecordset.Close
   oConexion.Close
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarPorGrupo(lcParametro As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorGrupo = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ParametrosSeleccionarPorGrupo"
     Set oParameter = .CreateParameter("@Grupo", adVarChar, adParamInput, 30, lcParametro): .Parameters.Append oParameter
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarPorGrupo = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function RetornaHoraServidorSQLserverFormatoGalenhos() As String
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHENTIDADES.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RetornaHoraServidorSQLserverFormatoGalenhos"
        Set oRecordset = .Execute
   End With
   If oRecordset.RecordCount > 0 Then
       RetornaHoraServidorSQLserverFormatoGalenhos = Trim(oRecordset.Fields!HoraSQLFormatoGalenhos)
   Else
       GoTo ManejadorDeError
   End If
   oRecordset.Close
   oConexion.Close
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



Sub DevuelveComboLlenoSegunDescripcion(ByRef CtlCombo As Object, lnIdParametro As Long)
    Dim lnFor As Integer, lnFor1 As Integer, lcTexto As String, lcTexto1 As String
    Dim mo_DOPArametro As New DOPArametro, mo_Parametros As New Parametros
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHENTIDADES.CadenaConexion
    Set mo_Parametros.Conexion = oConexion
    CtlCombo.Clear
    mo_DOPArametro.IdParametro = lnIdParametro
    If mo_Parametros.SeleccionarPorId(mo_DOPArametro) = True Then
        lcTexto = mo_DOPArametro.Descripcion
        lcTexto1 = ""
        lnFor = 4
        lnFor1 = 1
        Do While True
           If lnFor = Len(lcTexto) Then
              CtlCombo.AddItem Mid(lcTexto, lnFor1, 500)
              Exit Do
           ElseIf Mid(lcTexto, lnFor, 1) = "=" Then
               lcTexto1 = Mid(lcTexto, lnFor1, lnFor - lnFor1 - 1)
               CtlCombo.AddItem lcTexto1
               lnFor1 = lnFor - 1
           End If
           lnFor = lnFor + 1
        Loop
        For lnFor = 1 To CtlCombo.ListCount
            If Left(CtlCombo.List(lnFor - 1), 1) = Trim(mo_DOPArametro.ValorTexto) Then
               CtlCombo.ListIndex = lnFor - 1
               Exit For
            End If
        Next
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set mo_DOPArametro = Nothing
    Set mo_Parametros = Nothing
End Sub

Function SeleccionaFilaParametroValorInt(lnIdParametro As Long) As String
On Error GoTo ManejadorDeError
Dim vSeleccionaFilaParametro As String
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHENTIDADES.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ParametrosSeleccionarPorId"
        Set oParameter = .CreateParameter("@lnIdParametro", adInteger, adParamInput, 0, lnIdParametro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount > 0 Then
       vSeleccionaFilaParametro = Trim(Str(oRecordset.Fields!ValorInt))
        Select Case vSeleccionaFilaParametro
            Case "s", "n"
                 SeleccionaFilaParametroValorInt = UCase(vSeleccionaFilaParametro)
            Case Else
                 SeleccionaFilaParametroValorInt = vSeleccionaFilaParametro
        End Select
   Else
       GoTo ManejadorDeError
   End If
   oRecordset.Close
   oConexion.Close
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox "Parametro: " & lnIdParametro & Chr(13) & Err.Description
End Function

