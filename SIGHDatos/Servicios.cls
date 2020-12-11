VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Servicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla Servicios
'        Programado por: Castro W
'        Fecha: Marzo 2006
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
'   Descripción:    Inserta un registro a la tabla Servicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOServicio) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ServiciosAgregar"
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SVG", adVarChar, adParamInput, 8000, IIf(oTabla.SVG = "", Null, oTabla.SVG)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 6, IIf(oTabla.codigo = "", Null, oTabla.codigo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 50, IIf(oTabla.nombre = "", Null, oTabla.nombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, IIf(oTabla.IdEspecialidad = 0, Null, oTabla.IdEspecialidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, IIf(oTabla.IdTipoServicio = 0, Null, oTabla.IdTipoServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@soloTipoSexo", adInteger, adParamInput, 0, oTabla.soloTipoSexo): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@maximaEdad", adInteger, adParamInput, 0, oTabla.maximaEdad): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@codigoServicioSEM", adVarChar, adParamInput, 10, IIf(oTabla.codigoServicioSEM = "", Null, oTabla.codigoServicioSEM)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ubicacionSEM", adVarChar, adParamInput, 5, IIf(oTabla.ubicacionSEM = "", Null, oTabla.ubicacionSEM)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@codigoServicioHIS", adVarChar, adParamInput, 6, IIf(oTabla.codigoServicioHIS = "", Null, oTabla.codigoServicioHIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CostoCeroCE", adVarChar, adParamInput, 1, IIf(oTabla.CostoCeroCE = "", Null, oTabla.CostoCeroCE)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@minimaEdad", adInteger, adParamInput, 0, oTabla.minimaEdad): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Triaje", adBoolean, adParamInput, 0, IIf(oTabla.Triaje = True, 1, 0)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsObservacionEmergencia", adBoolean, adParamInput, 0, IIf(oTabla.EsObservacionEmergencia = True, 1, 0)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@UsaModuloNinoSano", adBoolean, adParamInput, 0, IIf(oTabla.UsaModuloNinoSano = True, 1, 0)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@UsaModuloMaterno", adBoolean, adParamInput, 0, IIf(oTabla.UsaModuloMaterno = True, 1, 0)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@UsaGalenHos", adBoolean, adParamInput, 0, IIf(oTabla.UsaGalenHos = True, 1, 0)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TipoEdad", adInteger, adParamInput, 0, oTabla.TipoEdad): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@UsaFua", adBoolean, adParamInput, 0, IIf(oTabla.UsaFUA = True, 1, 0)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@codigoServicioSuSalud", adVarChar, adParamInput, 7, IIf(oTabla.codigoServicioSuSalud = "", Null, oTabla.codigoServicioSuSalud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@codigoServicioFUA", adVarChar, adParamInput, 6, IIf(oTabla.codigoServicioFUA = "", Null, oTabla.codigoServicioFUA)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FuaTipoAnexo2015", adInteger, adParamInput, 0, IIf(oTabla.FuaTipoAnexo2015 = 0, Null, oTabla.FuaTipoAnexo2015)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@MaxCuposCitasAdelantadas", adInteger, adParamInput, 0, oTabla.MaxCuposCitasAdelantadas): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@MaxCuposAdicionales", adInteger, adParamInput, 0, oTabla.MaxCuposAdicionales): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@MaxCuposCitasHoySIS", adInteger, adParamInput, 0, oTabla.MaxCuposCitasHoySIS): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@MaxCuposCitasAdelandatasSIS", adInteger, adParamInput, 0, oTabla.MaxCuposCitasAdelandatasSIS): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@usaNacimiento", adInteger, adParamInput, 0, oTabla.usaNacimiento): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@emergenciaCorrelativo", adVarChar, adParamInput, 10, IIf(oTabla.emergenciaCorrelativo = "", Null, oTabla.emergenciaCorrelativo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NoUsaMensajeTexto", adInteger, adParamInput, 0, oTabla.NoUsaMensajeTexto): .Parameters.Append oParameter
       
       
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdServicio = .Parameters("@IdServicio")
   End With
 
   Insertar = True
   ms_MensajeError = ""
  GrabaHistorico oTabla
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Sub GrabaHistorico(oTabla As DOServicio)
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
    lcSql = "INSERT INTO Servicios" & _
"           (Nombre" & _
"           ,IdEspecialidad" & _
"           ,IdTipoServicio" & _
"           ,Codigo" & _
"           ,soloTipoSexo" & _
"           ,maximaEdad" & _
"           ,MinimaEdad" & _
"           ,EsObservacionEmergencia" & _
"           ,TipoEdad" & _
"           ,IdServicio,fecha0)" & _
"     Values" & _
"           (" & _
"           '" & oTabla.nombre & _
"           '," & oTabla.IdEspecialidad & _
"           ," & oTabla.IdTipoServicio & _
"           ,'" & oTabla.codigo & _
"           '," & oTabla.soloTipoSexo & _
"           ," & oTabla.maximaEdad & _
"           ," & oTabla.minimaEdad & _
"           ," & IIf(oTabla.EsObservacionEmergencia = True, "1", "0") & _
"           ," & oTabla.TipoEdad & _
"           ," & oTabla.IdServicio & ",'" & Format(Date, "dd/mm/yyyy") & "'" & _
"           )"
           
    oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    oConexionMDB.Close
ErrHistorico:
    Set oRsParametros = Nothing
    Set oRsMDB = Nothing
    Set oConexionMDB = Nothing
End Sub



'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla Servicios
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOServicio) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ServiciosModificar"
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SVG", adVarChar, adParamInput, 8000, IIf(oTabla.SVG = "", Null, oTabla.SVG)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 6, IIf(oTabla.codigo = "", Null, oTabla.codigo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 50, IIf(oTabla.nombre = "", Null, oTabla.nombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, IIf(oTabla.IdEspecialidad = 0, Null, oTabla.IdEspecialidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, IIf(oTabla.IdTipoServicio = 0, Null, oTabla.IdTipoServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@soloTipoSexo", adInteger, adParamInput, 0, oTabla.soloTipoSexo): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@maximaEdad", adInteger, adParamInput, 0, oTabla.maximaEdad): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@codigoServicioSEM", adVarChar, adParamInput, 10, IIf(oTabla.codigoServicioSEM = "", Null, oTabla.codigoServicioSEM)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ubicacionSEM", adVarChar, adParamInput, 5, IIf(oTabla.ubicacionSEM = "", Null, oTabla.ubicacionSEM)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@codigoServicioHIS", adVarChar, adParamInput, 6, IIf(oTabla.codigoServicioHIS = "", Null, oTabla.codigoServicioHIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CostoCeroCE", adVarChar, adParamInput, 1, IIf(oTabla.CostoCeroCE = "", Null, oTabla.CostoCeroCE)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@minimaEdad", adInteger, adParamInput, 0, oTabla.minimaEdad): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, oTabla.IdEstado): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Triaje", adBoolean, adParamInput, 0, IIf(oTabla.Triaje = True, 1, 0)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsObservacionEmergencia", adBoolean, adParamInput, 0, IIf(oTabla.EsObservacionEmergencia = True, 1, 0)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@UsaModuloNinoSano", adBoolean, adParamInput, 0, IIf(oTabla.UsaModuloNinoSano = True, 1, 0)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@UsaModuloMaterno", adBoolean, adParamInput, 0, IIf(oTabla.UsaModuloMaterno = True, 1, 0)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@UsaGalenHos", adBoolean, adParamInput, 0, IIf(oTabla.UsaGalenHos = True, 1, 0)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TipoEdad", adInteger, adParamInput, 0, oTabla.TipoEdad): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@UsaFua", adBoolean, adParamInput, 0, IIf(oTabla.UsaFUA = True, 1, 0)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@codigoServicioSuSalud", adVarChar, adParamInput, 7, IIf(oTabla.codigoServicioSuSalud = "", Null, oTabla.codigoServicioSuSalud)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@codigoServicioFUA", adVarChar, adParamInput, 6, IIf(oTabla.codigoServicioFUA = "", Null, oTabla.codigoServicioFUA)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FuaTipoAnexo2015", adInteger, adParamInput, 0, IIf(oTabla.FuaTipoAnexo2015 = 0, Null, oTabla.FuaTipoAnexo2015)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@MaxCuposCitasAdelantadas", adInteger, adParamInput, 0, oTabla.MaxCuposCitasAdelantadas): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@MaxCuposAdicionales", adInteger, adParamInput, 0, oTabla.MaxCuposAdicionales): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@MaxCuposCitasHoySIS", adInteger, adParamInput, 0, oTabla.MaxCuposCitasHoySIS): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@MaxCuposCitasAdelandatasSIS", adInteger, adParamInput, 0, oTabla.MaxCuposCitasAdelandatasSIS): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@usaNacimiento", adInteger, adParamInput, 0, oTabla.usaNacimiento): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@emergenciaCorrelativo", adVarChar, adParamInput, 10, IIf(oTabla.emergenciaCorrelativo = "", Null, oTabla.emergenciaCorrelativo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NoUsaMensajeTexto", adInteger, adParamInput, 0, oTabla.NoUsaMensajeTexto): .Parameters.Append oParameter

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
'   Descripción:    elimia o borra un registro a la tabla Servicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOServicio) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ServiciosEliminar"
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
 
   Eliminar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox Err.Number & " " + Err.Description, vbCritical, "Avise a soporte tecnico"
Exit Function
End Function
Sub ActualizaTabla(ByRef oTabla As DOServicio, oRecordset As Recordset)
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       oTabla.SVG = IIf(IsNull(oRecordset!SVG), "", oRecordset!SVG)
       oTabla.codigo = IIf(IsNull(oRecordset!codigo), "", oRecordset!codigo)
       oTabla.IdServicio = IIf(IsNull(oRecordset!IdServicio), 0, oRecordset!IdServicio)
       oTabla.nombre = IIf(IsNull(oRecordset!nombre), "", oRecordset!nombre)
       oTabla.IdEspecialidad = IIf(IsNull(oRecordset!IdEspecialidad), 0, oRecordset!IdEspecialidad)
       oTabla.IdTipoServicio = IIf(IsNull(oRecordset!IdTipoServicio), 0, oRecordset!IdTipoServicio)
       oTabla.soloTipoSexo = IIf(IsNull(oRecordset!soloTipoSexo), 0, oRecordset!soloTipoSexo)
       oTabla.maximaEdad = IIf(IsNull(oRecordset!maximaEdad), 0, oRecordset!maximaEdad)
       oTabla.codigoServicioSEM = IIf(IsNull(oRecordset!codigoServicioSEM), "", oRecordset!codigoServicioSEM)
       oTabla.ubicacionSEM = IIf(IsNull(oRecordset!ubicacionSEM), "", oRecordset!ubicacionSEM)
       oTabla.codigoServicioHIS = IIf(IsNull(oRecordset!codigoServicioHIS), "", oRecordset!codigoServicioHIS)
       oTabla.CostoCeroCE = IIf(IsNull(oRecordset!CostoCeroCE), "", oRecordset!CostoCeroCE)
       oTabla.minimaEdad = IIf(IsNull(oRecordset!minimaEdad), 0, oRecordset!minimaEdad)
       oTabla.IdEstado = IIf(IsNull(oRecordset!IdEstado), 0, oRecordset!IdEstado)
       oTabla.Triaje = IIf(IsNull(oRecordset!Triaje), 0, oRecordset!Triaje)
       oTabla.EsObservacionEmergencia = IIf(IsNull(oRecordset!EsObservacionEmergencia), 0, oRecordset!EsObservacionEmergencia)
       oTabla.UsaModuloNinoSano = IIf(IsNull(oRecordset!UsaModuloNinoSano), 0, oRecordset!UsaModuloNinoSano)
       oTabla.UsaModuloMaterno = IIf(IsNull(oRecordset!UsaModuloMaterno), 0, oRecordset!UsaModuloMaterno)
       oTabla.UsaGalenHos = IIf(IsNull(oRecordset!UsaGalenHos), 0, oRecordset!UsaGalenHos)
       oTabla.TipoEdad = IIf(IsNull(oRecordset!TipoEdad), 0, oRecordset!TipoEdad)
       oTabla.UsaFUA = IIf(IsNull(oRecordset!UsaFUA), 0, oRecordset!UsaFUA)
       oTabla.codigoServicioFUA = IIf(IsNull(oRecordset!codigoServicioFUA), "", oRecordset!codigoServicioFUA)
       oTabla.codigoServicioSuSalud = IIf(IsNull(oRecordset!codigoServicioSuSalud), "", oRecordset!codigoServicioSuSalud)
       oTabla.FuaTipoAnexo2015 = IIf(IsNull(oRecordset!FuaTipoAnexo2015), 0, oRecordset!FuaTipoAnexo2015)
       oTabla.MaxCuposAdicionales = IIf(IsNull(oRecordset!MaxCuposAdicionales), 0, oRecordset!MaxCuposAdicionales)
       oTabla.MaxCuposCitasAdelantadas = IIf(IsNull(oRecordset!MaxCuposCitasAdelantadas), 0, oRecordset!MaxCuposCitasAdelantadas)
       oTabla.MaxCuposCitasHoySIS = IIf(IsNull(oRecordset!MaxCuposCitasHoySIS), 0, oRecordset!MaxCuposCitasHoySIS)
       oTabla.MaxCuposCitasAdelandatasSIS = IIf(IsNull(oRecordset!MaxCuposCitasAdelandatasSIS), 0, oRecordset!MaxCuposCitasAdelandatasSIS)
       oTabla.usaNacimiento = IIf(IsNull(oRecordset!usaNacimiento), 0, oRecordset!usaNacimiento)
       If oTabla.IdTipoServicio = 2 Then
       oTabla.emergenciaCorrelativo = IIf(IsNull(oRecordset!emergenciaCorrelativo), Trim(Str(Year(Date))) & "00001", oRecordset!emergenciaCorrelativo)
       Else
       oTabla.emergenciaCorrelativo = ""
       End If
       oTabla.NoUsaMensajeTexto = IIf(IsNull(oRecordset!NoUsaMensajeTexto), 0, oRecordset!NoUsaMensajeTexto)

End Sub
'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla Servicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOServicio) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ServiciosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, oTabla.IdServicio): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       ActualizaTabla oTabla, oRecordset
   Else
       SeleccionarPorId = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:

   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox Err.Number & " " + Err.Description, vbInformation, "Avise a soporte tecnico"
Exit Function
End Function



Function SeleccionarPorCodigo(ByVal oTabla As DOServicio) As Boolean

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ServiciosXcodigo"
        Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 6, Trim(oTabla.codigo)): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If Not (oRecordset.EOF And oRecordset.BOF) Then
        ActualizaTabla oTabla, oRecordset
        SeleccionarPorCodigo = True
   Else
        SeleccionarPorCodigo = False
   End If
   oRecordset.Close
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarPorTipo(IdTipoServicio As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorTipo = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ServiciosSeleccionarPorTipo"
     Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, IdTipoServicio): .Parameters.Append oParameter
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarPorTipo = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
Function SeleccionarPorTipoV2(IdTipoServicio As Long, lnTipoEstados As sghTipoEstados) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorTipoV2 = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ServiciosSeleccionarPorTipoV2"
     Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, IdTipoServicio): .Parameters.Append oParameter
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   If lnTipoEstados <> sghFiltraAnuladosYactivos Then
      oRecordset.Filter = "idEstado=" & lnTipoEstados
   End If
   Set SeleccionarPorTipoV2 = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'SCCQ 20/02/2020 Cambio7 Incio
Function SeleccionarCEDisponibles(IdEspecialidad As Long, lnHoraInicio As String, lnHoraFin As String, lnFecha As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 Dim fnicio As String
 Dim ffin As String
   Set SeleccionarCEDisponibles = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ServiciosSeleccionarCEDisponiblesXFecha"
     Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, IdEspecialidad): .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@HoraInicio", adChar, adParamInput, 5, IIf(lnHoraInicio = "", Null, lnHoraInicio)): .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@HoraFin", adChar, adParamInput, 5, IIf(lnHoraFin = "", Null, lnHoraFin)): .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, Format(lnFecha, SIGHENTIDADES.DevuelveFechaSoloFormato_DMY)): .Parameters.Append oParameter
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarCEDisponibles = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'SCCQ 20/02/2020 Cambio7 Fin
Function ServiciosSeleccionarConsultoriosPorEspecialidadDEBB(IdEspecialidad As Long, lnTipoEstados As sghTipoEstados) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set ServiciosSeleccionarConsultoriosPorEspecialidadDEBB = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ServiciosSeleccionarConsultoriosPorEspecialidad"
     Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, IdEspecialidad): .Parameters.Append oParameter
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   If lnTipoEstados <> sghFiltraAnuladosYactivos Then
      oRecordset.Filter = "idEstado=" & lnTipoEstados
   End If
   
   Set ServiciosSeleccionarConsultoriosPorEspecialidadDEBB = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarPorTipoServicioDptoEspecialidad(IdTipoServicio As Long, IdDepartamento As Long, IdEspecialidad As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorTipoServicioDptoEspecialidad = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ServiciosSeleccionarPorTipoServicioDptoEspecialidad"
     Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, IdTipoServicio): .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamInput, 0, IdDepartamento): .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, IdEspecialidad): .Parameters.Append oParameter
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarPorTipoServicioDptoEspecialidad = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarPorTipoServicioYDpto(IdTipoServicio As Long, IdDepartamento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorTipoServicioYDpto = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ServiciosSeleccionarPorTipoServicioYDpto"
     Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, IdTipoServicio): .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@IdDepartamento", adInteger, adParamInput, 0, IdDepartamento): .Parameters.Append oParameter
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarPorTipoServicioYDpto = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function Filtrar(ByVal oTabla As DOServicio, lDepartamentoHospital As Long, lnTipoEstado As sghTipoEstados) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    If oTabla.IdTipoServicio <> 0 Then
         sWhere = sWhere + " Servicios.IdTipoServicio = " & oTabla.IdTipoServicio & " and "
    End If
    If oTabla.codigo <> 0 Then
         sWhere = sWhere + " Servicios.Codigo = '" & oTabla.codigo & "' and "
    End If
    If oTabla.nombre <> "" Then
         sWhere = sWhere + " Servicios.Nombre like '%" + oTabla.nombre + "%' and "
    End If
    If lDepartamentoHospital <> 0 Then
         sWhere = sWhere + " DepartamentosHospital.IdDepartamento = " & lDepartamentoHospital & " and "
    End If
    If lnTipoEstado <> sghFiltraAnuladosYactivos Then
        sWhere = sWhere + " Servicios.idEstado=" & lnTipoEstado & " and "
    End If
    If sWhere <> "" Then
         sSql = sSql + " where " + Left(sWhere, Len(sWhere) - 4)
    End If
    'yamill palomino
    If sghOrdenServicio.sghPorEspecialidad Then
       sSql = sSql + " order by Especialidades.Nombre, Servicios.Nombre"
    Else
       sSql = sSql + " order by Servicios.Nombre, Especialidades.Nombre"
    End If
'    sSql = sSql + " order by Especialidades.Nombre, Servicios.Nombre"
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ServiciosFiltrar"
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

Function ObtenerConElMismoCodigo(ByVal oTabla As DOServicio) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ServiciosXcodigoServicio"
        Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 6, oTabla.codigo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, oTabla.IdServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerConElMismoCodigo = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ObtenerConElMismoNombre(ByVal oTabla As DOServicio) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ServiciosXcodigoServicio"
        Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 50, oTabla.nombre): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, oTabla.IdServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerConElMismoNombre = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function SeleccionarIdDepartamento(ByVal lIdServicio As Long) As Long
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ServiciosFiltrarDepartamentoHosp"
        Set oParameter = .CreateParameter("@lIdServicio", adInteger, adParamInput, 0, lIdServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If Not (oRecordset.EOF And oRecordset.BOF) Then
        SeleccionarIdDepartamento = oRecordset!IdDepartamento
   Else
        SeleccionarIdDepartamento = 0
   End If
   oRecordset.Close
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarPorIdArchivero(lIdEmpleado As Long) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ServiciosXidEmpleado"
        Set oParameter = .CreateParameter("@lIdEmpleado", adInteger, adParamInput, 0, lIdEmpleado): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorIdArchivero = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



Sub Tabla_FactPuntosCarga_Actualiza()
On Error GoTo ErrAct
     Dim oRsServ As New Recordset
     Dim oRsPtoCarga As New Recordset
     Dim lcHE As String
     Dim oDOFactPuntoCarga As New DOFactPuntoCarga, oFactPuntosCarga As New FactPuntosCarga
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Set oFactPuntosCarga.Conexion = mo_Conexion
     Set oRsPtoCarga = oFactPuntosCarga.SeleccionarTodos
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ServiciosSoloEmergenciaHospitalizacion"
        Set oRsServ = .Execute
        Set oRsServ.ActiveConnection = Nothing
     End With
     If oRsServ.RecordCount > 0 Then
        'mo_Conexion.BeginTrans
        oRsServ.MoveFirst
        Do While Not oRsServ.EOF
           lcHE = IIf(oRsServ.Fields!IdTipoServicio = 3, " (H)", " (E)")
           oRsPtoCarga.MoveFirst
           oRsPtoCarga.Find "idServicio=" & oRsServ.Fields!IdServicio
           If oRsPtoCarga.EOF Then
              oDOFactPuntoCarga.IdPuntoCarga = oRsServ.Fields!IdServicio + 500
              oDOFactPuntoCarga.IdServicio = oRsServ.Fields!IdServicio
              oDOFactPuntoCarga.IdUPS = 5
              oDOFactPuntoCarga.Descripcion = Left(Trim(oRsServ.Fields!nombre) & lcHE, 50)
              If oFactPuntosCarga.Insertar(oDOFactPuntoCarga) = False Then
                 GoTo ErrAct
              End If
           Else
              oDOFactPuntoCarga.IdPuntoCarga = oRsPtoCarga.Fields!IdPuntoCarga
              If oFactPuntosCarga.SeleccionarPorId(oDOFactPuntoCarga) = False Then
                 GoTo ErrAct
              End If
              oDOFactPuntoCarga.IdUPS = 5
              oDOFactPuntoCarga.Descripcion = Left(Trim(oRsServ.Fields!nombre) & lcHE, 50)
              If oFactPuntosCarga.Modificar(oDOFactPuntoCarga) = False Then
                 GoTo ErrAct
              End If
           End If
           oRsServ.MoveNext
        Loop
        'mo_Conexion.CommitTrans
     End If
     oRsServ.Close
     oRsPtoCarga.Close
     Set oDOFactPuntoCarga = Nothing
     Set oFactPuntosCarga = Nothing
     Set oRsServ = Nothing
     Set oRsPtoCarga = Nothing
     Set oCommand = Nothing
     Exit Sub
ErrAct:
   MsgBox Err.Description
End Sub

Sub Tabla_FactPuntosCarga_Elimina(lnIdServicio As Long)
    'Elimina Punto de Carga(FactPuntosCarga) y ServiciosVsPuntosCarga(FactCatalogoServiciosPtos)
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim lcSql As String
    Dim oFactPuntosCarga As New FactPuntosCarga
    Set oFactPuntosCarga.Conexion = mo_Conexion
    Set oRsTmp1 = oFactPuntosCarga.SeleccionarTodos
    oRsTmp1.Filter = "idServicio=" & lnIdServicio
    If oRsTmp1.RecordCount > 0 Then
       'mo_Conexion.BeginTrans
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
           With oCommand
               .CommandType = adCmdStoredProc
               Set .ActiveConnection = mo_Conexion
               .CommandTimeout = 150
               .CommandText = "FactCatalogoServiciosPtosEliminaXidPuntoCarga"
               Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, oRsTmp1.Fields!IdPuntoCarga): .Parameters.Append oParameter
               .Execute
           End With
           Set oCommand = Nothing
           Set oParameter = Nothing
           With oCommand
               .CommandType = adCmdStoredProc
               Set .ActiveConnection = mo_Conexion
               .CommandTimeout = 150
               .CommandText = "FactPuntosCargaEliminar"
               Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, oRsTmp1.Fields!IdPuntoCarga): .Parameters.Append oParameter
               Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 0): .Parameters.Append oParameter
               .Execute
           End With
           Set oCommand = Nothing
           Set oParameter = Nothing
           oRsTmp1.MoveNext
       Loop
      ' mo_Conexion.CommitTrans
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Exit Sub
ErrAct:
   MsgBox Err.Description
End Sub



Function SeleccionarPorIdArchiveroTipoServicio(lIdEmpleado As Long, lcTipoServicio As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String, sSQL1 As String, sSQL2 As String
    sSql = ""
    sWhere = ""
    If lIdEmpleado > 0 Then
        sSQL1 = " Where ArchiveroServicio.IdEmpleado = " & lIdEmpleado
    End If
    sSQL2 = ""
    If lcTipoServicio <> "" Then
       If sSQL1 = "" Then
          sSQL2 = " Where Servicios.idTipoServicio in " & lcTipoServicio
       Else
          sSQL2 = " and Servicios.idTipoServicio in " & lcTipoServicio
       End If
    End If
    sSql = sSql + sSQL1 + sSQL2 + " order by Servicios.Nombre "
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ServiciosFiltrarPorIdArchiveroTipoServicio"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorIdArchiveroTipoServicio = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function DevuelveEspecialidadDelServicio(lnIdServicio As Long) As Long
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ServiciosDevuelveEspecialidad"
        Set oParameter = .CreateParameter("@lnIdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   DevuelveEspecialidadDelServicio = 0
   If oRecordset.RecordCount > 0 Then
       DevuelveEspecialidadDelServicio = oRecordset.Fields!IdEspecialidad
   End If
   oRecordset.Close
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


