VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Medicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla Medicos
'        Programado por: Barrantes D
'        Fecha: Mayo 2009
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
'   Descripción:    Inserta un registro a la tabla Medicos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOMedico) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "MedicosAgregar"
           Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, oTabla.IdEmpleado): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Colegiatura", adChar, adParamInput, 6, oTabla.Colegiatura): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@LoteHis", adVarChar, adParamInput, 2, IIf(oTabla.LoteHis = "", Null, oTabla.LoteHis)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idColegioHIS", adVarChar, adParamInput, 2, IIf(oTabla.idColegioHis = "", Null, oTabla.idColegioHis)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@rne", adVarChar, adParamInput, 50, IIf(oTabla.Rne = "", Null, oTabla.Rne)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@egresado", adBoolean, adParamInput, 0, IIf(oTabla.Egresado = True, 1, 0)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdMedico = .Parameters("@IdMedico")
   End With
 
   Insertar = True
   ms_MensajeError = ""
 GrabaHistorico oTabla
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla Medicos
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOMedico) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "MedicosModificar"
       Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, oTabla.IdEmpleado): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Colegiatura", adChar, adParamInput, 6, oTabla.Colegiatura): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, oTabla.IdMedico): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@LoteHis", adVarChar, adParamInput, 2, IIf(oTabla.LoteHis = "", Null, oTabla.LoteHis)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@idColegioHIS", adVarChar, adParamInput, 2, IIf(oTabla.idColegioHis = "", Null, oTabla.idColegioHis)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@rne", adVarChar, adParamInput, 50, IIf(oTabla.Rne = "", Null, oTabla.Rne)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@egresado", adBoolean, adParamInput, 0, IIf(oTabla.Egresado = True, 1, 0)): .Parameters.Append oParameter
       
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


Sub GrabaHistorico(oTabla As DOMedico)
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
    lcSql = "INSERT INTO Medicos" & _
"           (IdMedico" & _
"           ,Colegiatura" & _
"           ,IdEmpleado" & _
"           ,LoteHIS" & _
"           ,idColegioHIS" & _
"           ,rne" & _
"           ,egresado,fecha0)" & _
"     Values" & _
"           (" & _
"           " & oTabla.IdMedico & _
"           ,'" & oTabla.Colegiatura & _
"           '," & oTabla.IdEmpleado & _
"           ,'" & oTabla.LoteHis & _
"           ','" & oTabla.idColegioHis & _
"           ','" & oTabla.Rne & _
"           '," & IIf(oTabla.Egresado = True, "1", "0") & ",'" & Format(Date, "dd/mm/yyyy") & _
"') "
   oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    oConexionMDB.Close
ErrHistorico:
    Set oRsParametros = Nothing
    Set oRsMDB = Nothing
    Set oConexionMDB = Nothing
End Sub
'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla Medicos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOMedico) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "MedicosEliminar"
       Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, oTabla.IdMedico): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla Medicos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOMedico) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "MedicosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, oTabla.IdMedico): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdEmpleado = oRecordset("IdEmpleado")
       oTabla.Colegiatura = oRecordset("Colegiatura")
       oTabla.IdMedico = oRecordset("IdMedico")
       oTabla.LoteHis = IIf(IsNull(oRecordset("LoteHis")), "", oRecordset("LoteHis"))
       oTabla.idColegioHis = IIf(IsNull(oRecordset("idColegioHis")), "", oRecordset("idColegioHis"))
       oTabla.Rne = IIf(IsNull(oRecordset("rne")), "", oRecordset("rne"))
       oTabla.Egresado = IIf(IsNull(oRecordset!Egresado), 0, oRecordset!Egresado)
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

Function SeleccionarPorCodigo(ByVal oDOMedico As DOMedico, ByVal oDOEmpleado As DOEmpleado) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
   
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "MedicosXcodigoPlanilla"
        Set oParameter = .CreateParameter("@CodigoPlanilla", adVarChar, adParamInput, 1000, oDOEmpleado.CodigoPlanilla): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If Not (oRecordset.EOF And oRecordset.BOF) Then
        oDOMedico.IdMedico = oRecordset!IdMedico
        SeleccionarPorCodigo = SeleccionarPorId(oDOMedico)
   Else
        SeleccionarPorCodigo = False
   End If
   oRecordset.Close
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
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
     .CommandText = "MedicosSeleccionarTodos"
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

'Function FiltrarPorDptosYEspecialidad(IdDepartamento As Long, IdEspecialidad As Long) As ADODB.Recordset
'On Error GoTo ManejadorDeError
'Dim oRecordset As New Recordset
'Dim oCommand As New ADODB.Command
'Dim oParameter As ADODB.Parameter
'Dim sSql As String, sWhere As String, lcSql As String
'    sSql = ""
'    sWhere = ""
'    lcSql = ""
'    If IdDepartamento > 0 And IdEspecialidad > 0 Then
'         lcSql = lcSql & " where Especialidades.IdDepartamento = " & IdDepartamento & _
'                 "     and Especialidades.IdEspecialidad = " & IdEspecialidad & _
'                 " order by Nombre"
'    ElseIf IdDepartamento > 0 And IdEspecialidad = 0 Then
'         lcSql = lcSql & " where Especialidades.IdDepartamento = " & IdDepartamento & _
'                 " order by Nombre"
'    ElseIf IdDepartamento = 0 And IdEspecialidad > 0 Then
'         lcSql = lcSql & " where Especialidades.IdEspecialidad = " & IdEspecialidad & _
'                 " order by Nombre"
'    ElseIf IdDepartamento = 0 And IdEspecialidad = 0 Then
'         lcSql = lcSql & " order by Nombre"
'    End If
'
'    With oCommand
'        .CommandType = adCmdStoredProc
'        Set .ActiveConnection = mo_Conexion
'        .CommandTimeout = 150
'        .CommandText = "MedicosPorFiltro"
'        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcSql): .Parameters.Append oParameter
'        Set oRecordset = .Execute
'        Set oRecordset.ActiveConnection = Nothing
'   End With
'   Set FiltrarPorDptosYEspecialidad = oRecordset
'   Set oCommand = Nothing
'   Exit Function
'ManejadorDeError:
'    MsgBox Err.Description
'End Function
Function FiltrarPorDptosYEspecialidadEsActivo(IdDepartamento As Long, IdEspecialidad As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String, lcSql As String
    sSql = ""
    sWhere = ""
    lcSql = ""
    
    'yamill palomino
    If IdDepartamento > 0 And IdEspecialidad > 0 Then
         lcSql = lcSql & " where EsActivo = 1 and Especialidades.IdDepartamento = " & IdDepartamento & _
                 "     and Especialidades.IdEspecialidad = " & IdEspecialidad & _
                 " order by Nombre"
    ElseIf IdDepartamento > 0 And IdEspecialidad = 0 Then
         lcSql = lcSql & " where EsActivo = 1 and Especialidades.IdDepartamento = " & IdDepartamento & _
                 " order by Nombre"
    ElseIf IdDepartamento = 0 And IdEspecialidad > 0 Then
         lcSql = lcSql & " where EsActivo = 1 and Especialidades.IdEspecialidad = " & IdEspecialidad & _
                 " order by Nombre"
    ElseIf IdDepartamento = 0 And IdEspecialidad = 0 Then
         lcSql = lcSql & " where EsActivo = 1  order by Nombre"
    End If
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "MedicosPorFiltro"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FiltrarPorDptosYEspecialidadEsActivo = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
'mgaray20141014
Function FiltrarPorDptosYEspecialidadEsActivoConEspecialidad(IdDepartamento As Long, IdEspecialidad As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String, lcSql As String
    sSql = ""
    sWhere = ""
    lcSql = ""
    
    'yamill palomino
    If IdDepartamento > 0 And IdEspecialidad > 0 Then
         lcSql = lcSql & " where EsActivo = 1 and Especialidades.IdDepartamento = " & IdDepartamento & _
                 "     and Especialidades.IdEspecialidad = " & IdEspecialidad & _
                 " order by Nombre"
    ElseIf IdDepartamento > 0 And IdEspecialidad = 0 Then
         lcSql = lcSql & " where EsActivo = 1 and Especialidades.IdDepartamento = " & IdDepartamento & _
                 " order by Nombre"
    ElseIf IdDepartamento = 0 And IdEspecialidad > 0 Then
         lcSql = lcSql & " where EsActivo = 1 and Especialidades.IdEspecialidad = " & IdEspecialidad & _
                 " order by Nombre"
    ElseIf IdDepartamento = 0 And IdEspecialidad = 0 Then
         lcSql = lcSql & " where EsActivo = 1  order by Nombre"
    End If
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "MedicosPorFiltroConEspecialidad"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FiltrarPorDptosYEspecialidadEsActivoConEspecialidad = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function FiltrarPorProgramacion(IdDepartamento As Long, IdEspecialidad As Long, lnIdServicio As Long, daFecha As Date) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String, lcSql As String
    sSql = ""
    sWhere = ""
    lcSql = ""
    'mgaray201503
    If lnIdServicio > 0 Then
       lcSql = lcSql & " WHERE dbo.ProgramacionMedica.Fecha =CONVERT(DATETIME,'" & Format(daFecha, "dd/mm/yyyy") & "',103)" & _
                       "       and dbo.ProgramacionMedica.idServicio=" & lnIdServicio & _
                       " ORDER BY  dbo.Servicios.Nombre"
    ElseIf IdEspecialidad > 0 Then
       lcSql = lcSql & " WHERE dbo.ProgramacionMedica.Fecha =CONVERT(DATETIME,'" & Format(daFecha, "dd/mm/yyyy") & "',103)" & _
                       "       and dbo.ProgramacionMedica.IdEspecialidad=" & IdEspecialidad & _
                       "       and dbo.ProgramacionMedica.IdDepartamento=" & IdDepartamento & _
                       " ORDER BY  dbo.Servicios.Nombre"
    ElseIf IdDepartamento > 0 Then
       lcSql = lcSql & " WHERE dbo.ProgramacionMedica.Fecha =CONVERT(DATETIME,'" & Format(daFecha, "dd/mm/yyyy") & "',103)" & _
                       "       and dbo.ProgramacionMedica.IdDepartamento=" & IdDepartamento & _
                       " ORDER BY  dbo.Servicios.Nombre"
    Else
       lcSql = lcSql & " WHERE dbo.ProgramacionMedica.Fecha =CONVERT(DATETIME,'" & Format(daFecha, "dd/mm/yyyy") & "',103)" & _
                       " ORDER BY  dbo.Servicios.Nombre"
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "MedicosFiltrarPorProgramacion"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FiltrarPorProgramacion = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function Filtrar(oDOMedico As DOMedico, oDOEmpleado As DOEmpleado, lIdEspecialidad As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String, lcSql As String
    sSql = ""
    sWhere = ""
    lcSql = ""
    If oDOEmpleado.apellidopaterno <> "" Then
         sWhere = sWhere + " Empleados.ApellidoPaterno like '" + oDOEmpleado.apellidopaterno + "%' and "
    End If
    If oDOEmpleado.apellidomaterno <> "" Then
         sWhere = sWhere + " Empleados.ApellidoMaterno like '" + oDOEmpleado.apellidomaterno + "%' and "
    End If
    If oDOEmpleado.Nombres <> "" Then
         sWhere = sWhere + " Empleados.Nombres like '%" + oDOEmpleado.Nombres + "%' and "
    End If
    If oDOEmpleado.CodigoPlanilla <> "" Then
         sWhere = sWhere + " Empleados.CodigoPlanilla = '" + oDOEmpleado.CodigoPlanilla + "' and "
    End If
    If lIdEspecialidad <> 0 Then
         sWhere = sWhere + " MedicosEspecialidad.IdEspecialidad = " & lIdEspecialidad & " and "
    End If
    If sWhere <> "" Then
         sSql = sSql + " where " + Left(sWhere, Len(sWhere) - 4)
    End If
    sSql = sSql + " order by Empleados.ApellidoPaterno, Empleados.ApellidoMaterno, Empleados.Nombres"
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "MedicosFiltrar"
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


Function ObtenerDepartamento(IdMedico As Long) As Long
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   ObtenerDepartamento = 0
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "MedicosObtenerDepartamento"
     Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IdMedico): .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamOutput, 0): .Parameters.Append oParameter
     .Execute
     
     ObtenerDepartamento = .Parameters("@IdDepartamento")
      
   End With
   
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function FiltrarPorDptosYEspecialidad(IdDepartamento As Long, IdEspecialidad As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String, lcSql As String
    sSql = ""
    sWhere = ""
    lcSql = ""
    If IdDepartamento > 0 And IdEspecialidad > 0 Then
         lcSql = lcSql & " where and Especialidades.IdDepartamento = " & IdDepartamento & _
                 "     and Especialidades.IdEspecialidad = " & IdEspecialidad & _
                 " order by Nombre"
    ElseIf IdDepartamento > 0 And IdEspecialidad = 0 Then
         lcSql = lcSql & " where Especialidades.IdDepartamento = " & IdDepartamento & _
                 " order by Nombre"
    ElseIf IdDepartamento = 0 And IdEspecialidad > 0 Then
         lcSql = lcSql & " where Especialidades.IdEspecialidad = " & IdEspecialidad & _
                 " order by Nombre"
    ElseIf IdDepartamento = 0 And IdEspecialidad = 0 Then
         lcSql = lcSql & " order by Nombre"
    End If
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "MedicosPorFiltro"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FiltrarPorDptosYEspecialidad = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
