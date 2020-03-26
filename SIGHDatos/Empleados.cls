VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Empleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla Empleados
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
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
'   Descripción:    Inserta un registro a la tabla Empleados
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOEmpleado) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oCrypKey As New CrypKey.Util
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "EmpleadosAgregar"
           Set oParameter = .CreateParameter("@Usuario", adVarChar, adParamInput, 20, IIf(oTabla.Usuario = "", Null, oTabla.Usuario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Clave", adVarChar, adParamInput, 200, IIf(oTabla.Clave = "", Null, oCrypKey.EncryptString(oTabla.Clave))): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAlta", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaAlta = 0, Null, oTabla.FechaAlta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaIngreso", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaIngreso = 0, Null, oTabla.FechaIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoPlanilla", adVarChar, adParamInput, 8, IIf(oTabla.CodigoPlanilla = "", Null, oTabla.CodigoPlanilla)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@DNI", adChar, adParamInput, 20, IIf(oTabla.DNI = "", Null, oTabla.DNI)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoEmpleado", adInteger, adParamInput, 0, IIf(oTabla.IdTipoEmpleado = 0, Null, oTabla.IdTipoEmpleado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCondicionTrabajo", adInteger, adParamInput, 0, IIf(oTabla.IdCondicionTrabajo = 0, Null, oTabla.IdCondicionTrabajo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Nombres", adVarChar, adParamInput, 50, IIf(oTabla.Nombres = "", Null, oTabla.Nombres)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 50, IIf(oTabla.ApellidoMaterno = "", Null, oTabla.ApellidoMaterno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 50, IIf(oTabla.ApellidoPaterno = "", Null, oTabla.ApellidoPaterno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamOutput, 0): .Parameters.Append oParameter
          
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
'***************daniel barrantes**************
'***************Nuevos campos
'***************
           
           Set oParameter = .CreateParameter("@LoginEstado", adInteger, adParamInput, 0, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@LoginPC", adVarChar, adParamInput, 50, ""): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idDestacado", adInteger, adParamInput, 0, IIf(oTabla.idTipoDestacado = 0, Null, oTabla.idTipoDestacado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstablecimientoExterno", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoExterno = 0, Null, oTabla.IdEstablecimientoExterno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HisCodigoDigitador", adVarChar, adParamInput, 20, IIf(oTabla.HisCodigoDigitador = "", Null, oTabla.HisCodigoDigitador)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ReniecAutorizado", adBoolean, adParamInput, 0, IIf(oTabla.ReniecAutorizado = True, 1, 0)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idTipoDocumento", adInteger, adParamInput, 0, IIf(oTabla.IdTipoDocumento = 0, Null, oTabla.IdTipoDocumento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idSupervisor", adInteger, adParamInput, 0, IIf(oTabla.IdSupervisor = 0, Null, oTabla.IdSupervisor)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsActivo", adBoolean, adParamInput, 0, IIf(oTabla.EsActivo = True, 1, 0)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idTipoSexo", adInteger, adParamInput, 0, oTabla.idTipoSexo): .Parameters.Append oParameter
           'SCCQ 26/03/2020 Cambio2 Inicio
           Set oParameter = .CreateParameter("@idPais", adInteger, adParamInput, 0, IIf(oTabla.IdPais = 0, Null, oTabla.IdPais)): .Parameters.Append oParameter
           'SCCQ 26/03/2020 Cambio2 Fin
       .Execute
           oTabla.IdEmpleado = .Parameters("@IdEmpleado")
   End With
 
   Insertar = True
   ms_MensajeError = ""
   GrabaHistorico oTabla
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Sub GrabaHistorico(oTabla As DOEmpleado)
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
    lcSql = "INSERT INTO Empleados" & _
"           (idEmpleado,ApellidoPaterno,ApellidoMaterno,Nombres" & _
"           ,DNI,CodigoPlanilla,FechaNacimiento,fecha0)" & _
"     Values" & _
"           (" & oTabla.IdEmpleado & _
"            ,'" & oTabla.ApellidoPaterno & "','" & oTabla.ApellidoMaterno & "','" & oTabla.Nombres & _
"           ','" & oTabla.DNI & "','" & oTabla.CodigoPlanilla & "','" & oTabla.FechaNacimiento & "','" & Format(Date, "dd/mm/yyyy") & _
"           ')"
   oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    oConexionMDB.Close
ErrHistorico:
    Set oRsParametros = Nothing
    Set oRsMDB = Nothing
    Set oConexionMDB = Nothing
End Sub

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla Empleados
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOEmpleado) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oCrypKey As New CrypKey.Util
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "EmpleadosModificar"
           Set oParameter = .CreateParameter("@Usuario", adVarChar, adParamInput, 20, IIf(oTabla.Usuario = "", Null, oTabla.Usuario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Clave", adVarChar, adParamInput, 200, IIf(oTabla.Clave = "", Null, oCrypKey.EncryptString(oTabla.Clave))): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaAlta", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaAlta = 0, Null, oTabla.FechaAlta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaIngreso", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaIngreso = 0, Null, oTabla.FechaIngreso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoPlanilla", adVarChar, adParamInput, 8, IIf(oTabla.CodigoPlanilla = "", Null, oTabla.CodigoPlanilla)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@DNI", adChar, adParamInput, 20, IIf(oTabla.DNI = "", Null, oTabla.DNI)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoEmpleado", adInteger, adParamInput, 0, IIf(oTabla.IdTipoEmpleado = 0, Null, oTabla.IdTipoEmpleado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCondicionTrabajo", adInteger, adParamInput, 0, IIf(oTabla.IdCondicionTrabajo = 0, Null, oTabla.IdCondicionTrabajo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Nombres", adVarChar, adParamInput, 50, IIf(oTabla.Nombres = "", Null, oTabla.Nombres)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 50, IIf(oTabla.ApellidoMaterno = "", Null, oTabla.ApellidoMaterno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 50, IIf(oTabla.ApellidoPaterno = "", Null, oTabla.ApellidoPaterno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleado = 0, Null, oTabla.IdEmpleado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria):       .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@LoginEstado", adInteger, adParamInput, 0, oTabla.LoginEstado): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@LoginPC", adVarChar, adParamInput, 50, oTabla.loginPc): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idDestacado", adInteger, adParamInput, 0, IIf(oTabla.idTipoDestacado = 0, Null, oTabla.idTipoDestacado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstablecimientoExterno", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimientoExterno = 0, Null, oTabla.IdEstablecimientoExterno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HisCodigoDigitador", adVarChar, adParamInput, 20, IIf(oTabla.HisCodigoDigitador = "", Null, oTabla.HisCodigoDigitador)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ReniecAutorizado", adBoolean, adParamInput, 0, IIf(oTabla.ReniecAutorizado = True, 1, 0)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idTipoDocumento", adInteger, adParamInput, 0, IIf(oTabla.IdTipoDocumento = 0, Null, oTabla.IdTipoDocumento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idSupervisor", adInteger, adParamInput, 0, IIf(oTabla.IdSupervisor = 0, Null, oTabla.IdSupervisor)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsActivo", adBoolean, adParamInput, 0, IIf(oTabla.EsActivo = True, 1, 0)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idTipoSexo", adInteger, adParamInput, 0, oTabla.idTipoSexo): .Parameters.Append oParameter
           'SCCQ 26/03/2020 Cambio2 Inicio
           Set oParameter = .CreateParameter("@idPais", adInteger, adParamInput, 0, IIf(oTabla.IdPais = 0, Null, oTabla.IdPais)): .Parameters.Append oParameter
           'SCCQ 26/03/2020 Cambio2 Fin
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
'   Descripción:    elimia o borra un registro a la tabla Empleados
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOEmpleado) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "EmpleadosEliminar"
           Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleado = 0, Null, oTabla.IdEmpleado)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla Empleados
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOEmpleado) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oCrypKey As New CrypKey.Util
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "EmpleadosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, oTabla.IdEmpleado): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
        SeleccionarPorId = True
        oTabla.Usuario = IIf(IsNull(oRecordset!Usuario), "", oRecordset!Usuario)
        oTabla.Clave = IIf(IsNull(oRecordset!Clave), "", oRecordset!Clave)
        If oTabla.Clave <> "" Then
            oTabla.Clave = oCrypKey.DecryptString(oTabla.Clave)
        End If
        
       oTabla.FechaAlta = IIf(IsNull(oRecordset!FechaAlta), 0, oRecordset!FechaAlta)
       oTabla.FechaIngreso = IIf(IsNull(oRecordset!FechaIngreso), 0, oRecordset!FechaIngreso)
       oTabla.CodigoPlanilla = IIf(IsNull(oRecordset!CodigoPlanilla), "", oRecordset!CodigoPlanilla)
       oTabla.DNI = IIf(IsNull(oRecordset!DNI), "", oRecordset!DNI)
       oTabla.IdTipoEmpleado = IIf(IsNull(oRecordset!IdTipoEmpleado), 0, oRecordset!IdTipoEmpleado)
       oTabla.IdCondicionTrabajo = IIf(IsNull(oRecordset!IdCondicionTrabajo), 0, oRecordset!IdCondicionTrabajo)
       oTabla.Nombres = IIf(IsNull(oRecordset!Nombres), "", oRecordset!Nombres)
       oTabla.ApellidoMaterno = IIf(IsNull(oRecordset!ApellidoMaterno), "", oRecordset!ApellidoMaterno)
       oTabla.ApellidoPaterno = IIf(IsNull(oRecordset!ApellidoPaterno), "", oRecordset!ApellidoPaterno)
       oTabla.IdEmpleado = IIf(IsNull(oRecordset!IdEmpleado), 0, oRecordset!IdEmpleado)
       
'***************daniel barrantes**************
'***************nuevos campos
'***************
       oTabla.LoginEstado = IIf(IsNull(oRecordset!LoginEstado), 0, oRecordset!LoginEstado)
       oTabla.loginPc = IIf(IsNull(oRecordset!loginPc), "", oRecordset!loginPc)
       oTabla.FechaNacimiento = IIf(IsNull(oRecordset!FechaNacimiento), 0, oRecordset!FechaNacimiento)
       oTabla.idTipoDestacado = IIf(IsNull(oRecordset!idTipoDestacado), 0, oRecordset!idTipoDestacado)
       oTabla.IdEstablecimientoExterno = IIf(IsNull(oRecordset!IdEstablecimientoExterno), 0, oRecordset!IdEstablecimientoExterno)
       oTabla.HisCodigoDigitador = IIf(IsNull(oRecordset!HisCodigoDigitador), "", oRecordset!HisCodigoDigitador)
       oTabla.ReniecAutorizado = IIf(IsNull(oRecordset!ReniecAutorizado), 0, oRecordset!ReniecAutorizado)
       oTabla.IdTipoDocumento = IIf(IsNull(oRecordset!IdTipoDocumento), 0, oRecordset!IdTipoDocumento)
       oTabla.IdSupervisor = IIf(IsNull(oRecordset!IdSupervisor), 0, oRecordset!IdSupervisor)
       oTabla.EsActivo = IIf(IsNull(oRecordset!EsActivo), 0, oRecordset!EsActivo)
       oTabla.idTipoSexo = IIf(IsNull(oRecordset!idTipoSexo), 0, oRecordset!idTipoSexo)
       'SCCQ 25/03/2020 Cambio2 Inicio
       oTabla.IdPais = IIf(IsNull(oRecordset!IdPais), 0, oRecordset!IdPais)
       'SCCQ 25/03/2020 Cambio2 Fin
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
     .CommandText = "EmpleadosSeleccionarTodos"
     Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function Filtrar(oDOEmpleado As DOEmpleado) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sOrder As String, sWhere As String
    sSql = ""
    sWhere = ""
    If oDOEmpleado.ApellidoPaterno <> "" Then
         sWhere = sWhere + " Empleados.ApellidoPaterno like '" + oDOEmpleado.ApellidoPaterno + "%' and "
    End If
    If oDOEmpleado.ApellidoMaterno <> "" Then
         sWhere = sWhere + " Empleados.ApellidoMaterno like '" + oDOEmpleado.ApellidoMaterno + "%' and "
    End If
    If oDOEmpleado.Nombres <> "" Then
         sWhere = sWhere + " Empleados.Nombres like '%" + oDOEmpleado.Nombres + "%' and "
    End If
    If oDOEmpleado.CodigoPlanilla <> "" Then
         sWhere = sWhere + " Empleados.CodigoPlanilla = '" + oDOEmpleado.CodigoPlanilla + "' and "
    End If
    If sWhere <> "" Then
         sSql = sSql + " where " + Left(sWhere, Len(sWhere) - 4)
    End If
    sSql = sSql + " order by Empleados.ApellidoPaterno, Empleados.ApellidoMaterno, Empleados.Nombres"
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosFiltrar"
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


Function SeleccionarPorCodigo(ByVal oDOEmpleado As DOEmpleado) As Boolean

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
        .CommandText = "EmpleadosXcodigoPlanilla"
        Set oParameter = .CreateParameter("@CodigoPlanilla", adVarChar, adParamInput, 8, oDOEmpleado.CodigoPlanilla): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If Not (oRecordset.EOF And oRecordset.BOF) Then
        oDOEmpleado.IdEmpleado = oRecordset!IdEmpleado
        SeleccionarPorCodigo = SeleccionarPorId(oDOEmpleado)
   Else
        SeleccionarPorCodigo = False
   End If
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ObtenerConElMismoCodigoPlanilla(ByVal oTabla As DOEmpleado) As ADODB.Recordset
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
        .CommandText = "EmpleadosXidentificadorYcodigoPlanilla"
        Set oParameter = .CreateParameter("@idEmpleado", adInteger, adParamInput, 0, oTabla.IdEmpleado): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@CodigoPlanilla", adVarChar, adParamInput, 8, oTabla.CodigoPlanilla): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerConElMismoCodigoPlanilla = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function ObtenerConElMismoUsuario(ByVal oTabla As DOEmpleado) As ADODB.Recordset
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
        .CommandText = "EmpleadosXidentificadorYusuario"
        Set oParameter = .CreateParameter("@idEmpleado", adInteger, adParamInput, 0, oTabla.IdEmpleado): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Usuario", adVarChar, adParamInput, 20, oTabla.Usuario): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerConElMismoUsuario = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function Autenticar(sUsuario As String) As ADODB.Recordset
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
        .CommandText = "EmpleadosXusuario"
        Set oParameter = .CreateParameter("@sUsuario", adVarChar, adParamInput, 20, sUsuario): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set Autenticar = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function AutenticarMaquina(sUsuario As String, nMaquina As String) As ADODB.Recordset
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
        .CommandText = "EmpleadosXusuario"
        Set oParameter = .CreateParameter("@sUsuario", adVarChar, adParamInput, 20, sUsuario): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AutenticarMaquina = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarParaExoneracionServicio(lIdCuentaAtencion As Long, lIdEmpleadoActual As Long) As Recordset
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
        .CommandText = "EmpleadosXidentificadorYCuentaExonera"
        Set oParameter = .CreateParameter("@lIdEmpleadoActua", adInteger, adParamInput, 0, lIdEmpleadoActual): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lIdCuentaAtencion", adInteger, adParamInput, 0, lIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarParaExoneracionServicio = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarParaPendientePagoServicio(lIdCuentaAtencion As Long, lIdEmpleadoActual As Long) As Recordset
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
        .CommandText = "EmpleadosXidentificadorYCuentaAutorizaPendiente"
        Set oParameter = .CreateParameter("@lIdEmpleadoActual", adInteger, adParamInput, 0, lIdEmpleadoActual): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lIdCuentaAtencion", adInteger, adParamInput, 0, lIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarParaPendientePagoServicio = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarParaModificacionDeCuentaServicio(lIdCuentaAtencion As Long, lIdEmpleadoActual As Long) As Recordset
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
        .CommandText = "EmpleadosSeleccionarParaModificacionDeCuentaServicio"
        Set oParameter = .CreateParameter("@lIdEmpleadoActual", adInteger, adParamInput, 0, lIdEmpleadoActual): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lIdCuentaAtencion", adInteger, adParamInput, 0, lIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarParaModificacionDeCuentaServicio = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'------------------------Bienes e Insumos

Function SeleccionarParaExoneracionBienesInsumos(lIdCuentaAtencion As Long, lIdEmpleadoActual As Long) As Recordset
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
        .CommandText = "EmpleadosSeleccionarParaExoneracionBienesInsumos"
        Set oParameter = .CreateParameter("@lIdEmpleadoActual", adInteger, adParamInput, 0, lIdEmpleadoActual): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lIdCuentaAtencion", adInteger, adParamInput, 0, lIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarParaExoneracionBienesInsumos = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarParaPendientePagoBienesInsumos(lIdCuentaAtencion As Long, lIdEmpleadoActual As Long) As Recordset
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
        .CommandText = "EmpleadosSeleccionarParaPendientePagoBienesInsumos"
        Set oParameter = .CreateParameter("@lIdEmpleadoActual", adInteger, adParamInput, 0, lIdEmpleadoActual): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lIdCuentaAtencion", adInteger, adParamInput, 0, lIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarParaPendientePagoBienesInsumos = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarParaModificacionDeCuentaBienesInsumos(lIdCuentaAtencion As Long, lIdEmpleadoActual As Long) As Recordset
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
        .CommandText = "EmpleadosSeleccionarParaModificacionDeCuentaBienesInsumos"
        Set oParameter = .CreateParameter("@lIdEmpleadoActual", adInteger, adParamInput, 0, lIdEmpleadoActual): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lIdCuentaAtencion", adInteger, adParamInput, 0, lIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarParaModificacionDeCuentaBienesInsumos = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ObtenerConElMismoDNI(lcDNI As String, lnIdTipoDocumento As Long) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sOrder As String, sWhere As String
Dim oConexion As New ADODB.Connection
        
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHENTIDADES.CadenaConexion
    sSql = ""
    sWhere = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosObtenerConElMismoDNI"
        Set oParameter = .CreateParameter("@lnIdTipoDocumento", adInteger, adParamInput, 0, lnIdTipoDocumento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcDNI", adVarChar, adParamInput, 20, lcDNI): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerConElMismoDNI = oRecordset
   Set oCommand = Nothing
   Set oConexion = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ObtenerConLaMismaCOLEGIATURA(lcCOLEGIATURA As String) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sOrder As String, sWhere As String
Dim oConexion As New ADODB.Connection
        
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHENTIDADES.CadenaConexion
    sSql = ""
    sWhere = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosXcolegiatura"
        Set oParameter = .CreateParameter("@lcCOLEGIATURA", adVarChar, adParamInput, 6, lcCOLEGIATURA): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerConLaMismaCOLEGIATURA = oRecordset
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



