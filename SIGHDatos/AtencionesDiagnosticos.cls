VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionesDiagnosticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla AtencionesDiagnosticos
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
'   Descripción:    Inserta un registro a la tabla AtencionesDiagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOAtencionDiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesDiagnosticosAgregar"
           Set oParameter = .CreateParameter("@IdSubclasificacionDx", adInteger, adParamInput, 0, IIf(oTabla.IdSubClasificacionDX = 0, Null, oTabla.IdSubClasificacionDX)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdClasificacionDx", adInteger, adParamInput, 0, IIf(oTabla.IdClasificacionDx = 0, Null, oTabla.IdClasificacionDx)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionDiagnostico", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@labConfHIS", adVarChar, adParamInput, 3, IIf(oTabla.labConfHIS = "", Null, oTabla.labConfHIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@GrupoHIS", adInteger, adParamInput, 0, oTabla.GrupoHIS): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SubGrupoHIS", adInteger, adParamInput, 0, oTabla.SubGrupoHIS): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdAtencionDiagnostico = .Parameters("@IdAtencionDiagnostico")
   End With
 
   Insertar = True
   ms_MensajeError = ""
   GrabaHistorico oTabla
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function


Sub GrabaHistorico(oTabla As DOAtencionDiagnostico)
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
    '
    lcSql = "select fechaIngreso from atenciones where idatencion=" & oTabla.idAtencion
    If oRsParametros.State = 1 Then oRsParametros.Close
    oRsParametros.Open lcSql, mo_Conexion, adOpenKeyset, adLockOptimistic
    '
    
    lcSql = "INSERT INTO AtencionesDiagnosticos" & _
"           (IdAtencion" & _
"           ,IdClasificacionDx" & _
"           ,IdDiagnostico" & _
"           ,IdSubclasificacionDx" & _
"           ,labConfHIS" & _
"           ,GrupoHIS" & _
"           ,SubGrupoHIS,fecha0)" & _
"     Values (" & _
"            " & oTabla.idAtencion & _
"            ," & oTabla.IdClasificacionDx & _
"            ," & oTabla.IdDiagnostico & _
"            ," & oTabla.IdSubClasificacionDX & _
"            ,'" & oTabla.labConfHIS & _
"            '," & oTabla.GrupoHIS & _
"            ," & oTabla.SubGrupoHIS & ",'" & Format(oRsParametros!FechaIngreso, "dd/mm/yyyy") & _
"            ')"
   oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    oConexionMDB.Close
ErrHistorico:
    Set oRsParametros = Nothing
    Set oRsMDB = Nothing
    Set oConexionMDB = Nothing
End Sub

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla AtencionesDiagnosticos
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtencionDiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesDiagnosticosModificar"
           Set oParameter = .CreateParameter("@IdSubclasificacionDx", adInteger, adParamInput, 0, IIf(oTabla.IdSubClasificacionDX = 0, Null, oTabla.IdSubClasificacionDX)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdClasificacionDx", adInteger, adParamInput, 0, IIf(oTabla.IdClasificacionDx = 0, Null, oTabla.IdClasificacionDx)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionDiagnostico = 0, Null, oTabla.IdAtencionDiagnostico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@GrupoHIS", adInteger, adParamInput, 0, oTabla.GrupoHIS): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SubGrupoHIS", adInteger, adParamInput, 0, oTabla.SubGrupoHIS): .Parameters.Append oParameter
           
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
'   Descripción:    elimia o borra un registro a la tabla AtencionesDiagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtencionDiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesDiagnosticosEliminar"
           Set oParameter = .CreateParameter("@IdAtencionDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionDiagnostico = 0, Null, oTabla.IdAtencionDiagnostico)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla AtencionesDiagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtencionDiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtencionesDiagnosticosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdAtencionDiagnostico", adInteger, adParamInput, 0, oTabla.IdAtencionDiagnostico): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdSubClasificacionDX = IIf(IsNull(oRecordset!IdSubClasificacionDX), 0, oRecordset!IdSubClasificacionDX)
       oTabla.IdClasificacionDx = IIf(IsNull(oRecordset!IdClasificacionDx), 0, oRecordset!IdClasificacionDx)
       oTabla.IdDiagnostico = IIf(IsNull(oRecordset!IdDiagnostico), 0, oRecordset!IdDiagnostico)
       oTabla.IdAtencionDiagnostico = IIf(IsNull(oRecordset!IdAtencionDiagnostico), 0, oRecordset!IdAtencionDiagnostico)
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       oTabla.GrupoHIS = IIf(IsNull(oRecordset!GrupoHIS), 0, oRecordset!GrupoHIS)
       oTabla.SubGrupoHIS = IIf(IsNull(oRecordset!SubGrupoHIS), 0, oRecordset!SubGrupoHIS)
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

Function EliminarDiagnosticosPorAtencion(lIdAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesDiagnosticosEliminarXIdAtencion"
        Set oParameter = .CreateParameter("@lIdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        .Execute
   End With
   EliminarDiagnosticosPorAtencion = True
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarPorAtencion(lIdAtencion As Long, lIdTipoDiagnostico As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorAtencion = Nothing
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtencionesDiagnosticosSeleccionarPorAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Tipodiagnostico", adInteger, adParamInput, 0, lIdTipoDiagnostico): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
'   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorAtencion = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function



Function SeleccionarComplicaciones(lIdAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarComplicaciones = Nothing
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtencionesDiagnosticosSeleccionarComplicaciones"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarComplicaciones = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function ActualizarComplicaciones(oComplicaciones As Collection, lIdAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oDOAtencionDiagnostico As New DOAtencionDiagnostico
 
   ActualizarComplicaciones = False
   
   'elimina los complicaciones
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesDiagnosticosEliminarPorIdAtencion"
        Set oParameter = .CreateParameter("@lIdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        .Execute
   End With
   
   If Not oComplicaciones Is Nothing Then
        For Each oDOAtencionDiagnostico In oComplicaciones
            oDOAtencionDiagnostico.idAtencion = lIdAtencion
            If Not Insertar(oDOAtencionDiagnostico) Then
                Exit Function
            End If
         Next
    End If
    
   ActualizarComplicaciones = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function EliminarComplicaciones(lIdAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
 
   EliminarComplicaciones = False
   
   'elimina los diagnosticos

   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesDiagnosticosEliminarPorIdAtencion"
        Set oParameter = .CreateParameter("@lIdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        .Execute
   End With

   
   EliminarComplicaciones = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function ActualizarDiagnosticosAtencion(oDiagnosticos As Collection, lIdAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oDOAtencionDiagnostico As New DOAtencionDiagnostico
 
   ActualizarDiagnosticosAtencion = False
   
   'elimina todos los dx menos los de interconsultas

   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesDiagnosticosEliminarXIdAtencion"
        Set oParameter = .CreateParameter("@lIdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        .Execute
   End With
   
   If Not oDiagnosticos Is Nothing Then
        For Each oDOAtencionDiagnostico In oDiagnosticos
            oDOAtencionDiagnostico.idAtencion = lIdAtencion
            If Not Insertar(oDOAtencionDiagnostico) Then
                Exit Function
            End If
         Next
    End If
    
   ActualizarDiagnosticosAtencion = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   MsgBox Err.Description

   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function EliminarDiagnosticosDeAtencion(lIdAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
 
   EliminarDiagnosticosDeAtencion = False
   
   'elimina los diagnosticos
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesDiagnosticosEliminarXIdAtencion"
        Set oParameter = .CreateParameter("@lIdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        .Execute
   End With
   
   EliminarDiagnosticosDeAtencion = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarDiagnosticosDeEgreso(lIdAtencion As Long) As ADODB.Recordset 'scrafet reporte 5
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarDiagnosticosDeEgreso = Nothing
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtencionesDiagnosticosEgresoSeleccionarPorAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarDiagnosticosDeEgreso = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarTodosPorIdAtencion(lnIdAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesDiagnosticosSeleccionarTodosPorIdAtencion"
        Set oParameter = .CreateParameter("@lnIdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarTodosPorIdAtencion = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'Actualizado 15102014
Function SeleccionarPorIdAtencionIdDx(ByVal oTabla As DOAtencionDiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorIdAtencionIdDx = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtencionesDiagnosticosSeleccionarPorIdAtencionIdDx"
       Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oTabla.idAtencion): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, oTabla.IdDiagnostico): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorIdAtencionIdDx = True
       oTabla.IdSubClasificacionDX = IIf(IsNull(oRecordset!IdSubClasificacionDX), 0, oRecordset!IdSubClasificacionDX)
       oTabla.IdClasificacionDx = IIf(IsNull(oRecordset!IdClasificacionDx), 0, oRecordset!IdClasificacionDx)
       oTabla.IdDiagnostico = IIf(IsNull(oRecordset!IdDiagnostico), 0, oRecordset!IdDiagnostico)
       oTabla.IdAtencionDiagnostico = IIf(IsNull(oRecordset!IdAtencionDiagnostico), 0, oRecordset!IdAtencionDiagnostico)
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
   Else
       SeleccionarPorIdAtencionIdDx = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'mgaray20141015

Function ModificarTodosCampos(ByVal oTabla As DOAtencionDiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   ModificarTodosCampos = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesDiagnosticosModificarTodosCampos"
           Set oParameter = .CreateParameter("@IdSubclasificacionDx", adInteger, adParamInput, 0, IIf(oTabla.IdSubClasificacionDX = 0, Null, oTabla.IdSubClasificacionDX)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdClasificacionDx", adInteger, adParamInput, 0, IIf(oTabla.IdClasificacionDx = 0, Null, oTabla.IdClasificacionDx)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencionDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdAtencionDiagnostico = 0, Null, oTabla.IdAtencionDiagnostico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@labConfHIS", adVarChar, adParamInput, 3, IIf(oTabla.labConfHIS = "", Null, oTabla.labConfHIS)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
 
   ModificarTodosCampos = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Public Function SeleccionPorIdAtencionDiagnosticoLab(ByVal oTabla As DOAtencionDiagnostico) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesDiagnosticosSeleccionPorIdAtencionDiagnosticoLab"
        Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@labConfHIS", adVarChar, adParamInput, 3, IIf(oTabla.labConfHIS = "", Null, Trim(oTabla.labConfHIS))): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set SeleccionPorIdAtencionDiagnosticoLab = oRecordset
    Set oCommand = Nothing
    Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

