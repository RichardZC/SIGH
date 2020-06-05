VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Diagnosticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla Diagnosticos
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
'   Descripción:    Inserta un registro a la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DODiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "DiagnosticosAgregar"
           Set oParameter = .CreateParameter("@Intrahospitalario", adBoolean, adParamInput, 0, IIf(oTabla.Intrahospitalario = 0, Null, oTabla.Intrahospitalario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 250, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoCIE9", adChar, adParamInput, 5, IIf(oTabla.CodigoCIE9 = "", Null, oTabla.CodigoCIE9)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoCIE10", adChar, adParamInput, 7, IIf(oTabla.CodigoCIE10 = "", Null, oTabla.CodigoCIE10)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoExportacion", adChar, adParamInput, 5, IIf(oTabla.CodigoExportacion = "", Null, oTabla.CodigoExportacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.idTipoSexo = 0, Null, oTabla.idTipoSexo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Morbilidad", adBoolean, adParamInput, 0, IIf(oTabla.Morbilidad = 0, Null, oTabla.Morbilidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCategoria", adInteger, adParamInput, 0, IIf(oTabla.IdCategoria = 0, Null, oTabla.IdCategoria)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Restriccion", adBoolean, adParamInput, 0, IIf(oTabla.Restriccion = 0, Null, oTabla.Restriccion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EdadMaxDias", adInteger, adParamInput, 0, IIf(oTabla.EdadMaxDias = 0, Null, oTabla.EdadMaxDias)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EdadMinDias", adInteger, adParamInput, 0, IIf(oTabla.EdadMinDias = 0, Null, oTabla.EdadMinDias)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoCIE2004", adChar, adParamInput, 7, IIf(oTabla.CodigoCIE2004 = "", Null, oTabla.CodigoCIE2004)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCapitulo", adInteger, adParamInput, 0, IIf(oTabla.IdCapitulo = 0, Null, oTabla.IdCapitulo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGrupo", adInteger, adParamInput, 0, IIf(oTabla.IdGrupo = 0, Null, oTabla.IdGrupo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Gestacion", adBoolean, adParamInput, 0, IIf(oTabla.Gestacion = 0, Null, oTabla.Gestacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@DescripcionMINSA", adVarChar, adParamInput, 250, IIf(oTabla.DescripcionMINSA = "", Null, oTabla.DescripcionMINSA)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@codigoCIEsinPto", adChar, adParamInput, 7, IIf(oTabla.codigoCIEsinPto = "", Null, oTabla.codigoCIEsinPto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaInicioVigencia", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaInicioVigencia = 0, Null, oTabla.FechaInicioVigencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsActivo", adBoolean, adParamInput, 0, IIf(oTabla.EsActivo = 0, False, oTabla.EsActivo)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdDiagnostico = .Parameters("@IdDiagnostico")
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
'   Descripción:    modifica un registro a la tabla Diagnosticos
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DODiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "DiagnosticosModificar"
           Set oParameter = .CreateParameter("@Intrahospitalario", adBoolean, adParamInput, 0, IIf(oTabla.Intrahospitalario = 0, Null, oTabla.Intrahospitalario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 250, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoCIE9", adChar, adParamInput, 5, IIf(oTabla.CodigoCIE9 = "", Null, oTabla.CodigoCIE9)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoCIE10", adChar, adParamInput, 7, IIf(oTabla.CodigoCIE10 = "", Null, oTabla.CodigoCIE10)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoExportacion", adChar, adParamInput, 5, IIf(oTabla.CodigoExportacion = "", Null, oTabla.CodigoExportacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.idTipoSexo = 0, Null, oTabla.idTipoSexo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Morbilidad", adBoolean, adParamInput, 0, IIf(oTabla.Morbilidad = 0, Null, oTabla.Morbilidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCategoria", adInteger, adParamInput, 0, IIf(oTabla.IdCategoria = 0, Null, oTabla.IdCategoria)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Restriccion", adBoolean, adParamInput, 0, IIf(oTabla.Restriccion = 0, Null, oTabla.Restriccion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EdadMaxDias", adInteger, adParamInput, 0, IIf(oTabla.EdadMaxDias = 0, Null, oTabla.EdadMaxDias)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EdadMinDias", adInteger, adParamInput, 0, IIf(oTabla.EdadMinDias = 0, Null, oTabla.EdadMinDias)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoCIE2004", adChar, adParamInput, 7, IIf(oTabla.CodigoCIE2004 = "", Null, oTabla.CodigoCIE2004)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCapitulo", adInteger, adParamInput, 0, IIf(oTabla.IdCapitulo = 0, Null, oTabla.IdCapitulo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGrupo", adInteger, adParamInput, 0, IIf(oTabla.IdGrupo = 0, Null, oTabla.IdGrupo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Gestacion", adBoolean, adParamInput, 0, IIf(oTabla.Gestacion = 0, Null, oTabla.Gestacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@DescripcionMINSA", adVarChar, adParamInput, 250, IIf(oTabla.DescripcionMINSA = "", Null, oTabla.DescripcionMINSA)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@codigoCIEsinPto", adChar, adParamInput, 7, IIf(oTabla.codigoCIEsinPto = "", Null, oTabla.codigoCIEsinPto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaInicioVigencia", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaInicioVigencia = 0, Null, oTabla.FechaInicioVigencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsActivo", adBoolean, adParamInput, 0, IIf(oTabla.EsActivo = 0, False, oTabla.EsActivo)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DODiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "DiagnosticosEliminar"
           Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DODiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "DiagnosticosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, oTabla.IdDiagnostico): .Parameters.Append oParameter
     Set oRecordset = .Execute
     Set .ActiveConnection = Nothing
     Set oRecordset.ActiveConnection = Nothing
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.Intrahospitalario = IIf(IsNull(oRecordset!Intrahospitalario), 0, oRecordset!Intrahospitalario)
       oTabla.IdDiagnostico = IIf(IsNull(oRecordset!IdDiagnostico), 0, oRecordset!IdDiagnostico)
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
       oTabla.CodigoCIE9 = IIf(IsNull(oRecordset!CodigoCIE9), "", oRecordset!CodigoCIE9)
       oTabla.CodigoCIE10 = IIf(IsNull(oRecordset!CodigoCIE10), "", oRecordset!CodigoCIE10)
       oTabla.CodigoExportacion = IIf(IsNull(oRecordset!CodigoExportacion), "", oRecordset!CodigoExportacion)
       oTabla.idTipoSexo = IIf(IsNull(oRecordset!idTipoSexo), 0, oRecordset!idTipoSexo)
       oTabla.Morbilidad = IIf(IsNull(oRecordset!Morbilidad), 0, oRecordset!Morbilidad)
       oTabla.IdCategoria = IIf(IsNull(oRecordset!IdCategoria), 0, oRecordset!IdCategoria)
       oTabla.Restriccion = IIf(IsNull(oRecordset!Restriccion), 0, oRecordset!Restriccion)
       oTabla.EdadMaxDias = IIf(IsNull(oRecordset!EdadMaxDias), 0, oRecordset!EdadMaxDias)
       oTabla.EdadMinDias = IIf(IsNull(oRecordset!EdadMinDias), 0, oRecordset!EdadMinDias)
       oTabla.CodigoCIE2004 = IIf(IsNull(oRecordset!CodigoCIE2004), "", oRecordset!CodigoCIE2004)
       oTabla.IdCapitulo = IIf(IsNull(oRecordset!IdCapitulo), 0, oRecordset!IdCapitulo)
       oTabla.IdGrupo = IIf(IsNull(oRecordset!IdGrupo), 0, oRecordset!IdGrupo)
       oTabla.Gestacion = IIf(IsNull(oRecordset!Gestacion), 0, oRecordset!Gestacion)
       oTabla.DescripcionMINSA = IIf(IsNull(oRecordset!DescripcionMINSA), "", oRecordset!DescripcionMINSA)
       oTabla.codigoCIEsinPto = IIf(IsNull(oRecordset!codigoCIEsinPto), "", oRecordset!codigoCIEsinPto)
       oTabla.FechaInicioVigencia = IIf(IsNull(oRecordset!FechaInicioVigencia), 0, oRecordset!FechaInicioVigencia)
       oTabla.EsActivo = IIf(IsNull(oRecordset!EsActivo), False, oRecordset!EsActivo)
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
     .CommandText = "DiagnosticosSeleccionarTodos"
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

Function Filtrar(oDoDiagnostico As DODiagnostico, lbSoloMuestraDxGalenHos As Boolean, lbUSAcodigoCIEsinPto As Boolean) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sOrder As String, sWhere As String
    sSql = ""
    sOrder = ""
       sWhere = ""
'
'       If lbUSAcodigoCIEsinPto = True Then
'           If oDoDiagnostico.codigoCIEsinPto <> "" Then
'               sWhere = sWhere + " codigoCIEsinPto like '" + oDoDiagnostico.codigoCIEsinPto + "%' and "
'           End If
'       Else
'           If oDoDiagnostico.CodigoCIE2004 <> "" Then
'              sWhere = sWhere + " CodigoCIE2004 like '" + oDoDiagnostico.CodigoCIE2004 + "%' and "
'           End If
'       End If
'       If oDoDiagnostico.Descripcion <> "" And oDoDiagnostico.Descripcion <> "%%" Then
'            If Left(oDoDiagnostico.Descripcion, 1) <> "%" Then
'               oDoDiagnostico.Descripcion = Trim(oDoDiagnostico.Descripcion) & "%"
'            End If
'            sWhere = sWhere + " Descripcion like '" + oDoDiagnostico.Descripcion + "' and "
'       End If
'       If lbSoloMuestraDxGalenHos = True Then
'           sWhere = sWhere + " not (descripcionMINSA is null) and "
'       End If
'       If sWhere <> "" Then
'            sWhere = " Where " & Left(sWhere, Len(sWhere) - 4)
'       End If
'
'       sWhere = sWhere + " order by  Descripcion,CodigoCIE2004 "
    sWhere = getStringWhereForFilter(oDoDiagnostico, lbSoloMuestraDxGalenHos, lbUSAcodigoCIEsinPto, False)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "DiagnosticosFiltrar"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sWhere): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lbUSAcodigoCIEsinPto", adBoolean, adParamInput, 0, lbUSAcodigoCIEsinPto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set Filtrar = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'mgaray09
Function FiltrarSoloActivos(oDoDiagnostico As DODiagnostico, lbSoloMuestraDxGalenHos As Boolean, lbUSAcodigoCIEsinPto As Boolean) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sOrder As String, sWhere As String
    sSql = ""
    sWhere = getStringWhereForFilter(oDoDiagnostico, lbSoloMuestraDxGalenHos, lbUSAcodigoCIEsinPto, True)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "DiagnosticosFiltrar"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sWhere): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lbUSAcodigoCIEsinPto", adBoolean, adParamInput, 0, lbUSAcodigoCIEsinPto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FiltrarSoloActivos = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


'mgaray09
Private Function getStringWhereForFilter(oDoDiagnostico As DODiagnostico, lbSoloMuestraDxGalenHos As Boolean, lbUSAcodigoCIEsinPto As Boolean, _
            MostrarSoloActivos As Boolean)
    Dim sSql As String, sOrder As String, sWhere As String
    sSql = ""
    sOrder = ""
       sWhere = ""

       If lbUSAcodigoCIEsinPto = True Then
           If oDoDiagnostico.codigoCIEsinPto <> "" Then
               sWhere = sWhere + " codigoCIEsinPto like '" + oDoDiagnostico.codigoCIEsinPto + "%' and "
           End If
       Else
           If oDoDiagnostico.CodigoCIE2004 <> "" Then
              sWhere = sWhere + " CodigoCIE2004 like '" + oDoDiagnostico.CodigoCIE2004 + "%' and "
           End If
       End If
       If oDoDiagnostico.Descripcion <> "" And oDoDiagnostico.Descripcion <> "%%" Then
            If Left(oDoDiagnostico.Descripcion, 1) <> "%" Then
               oDoDiagnostico.Descripcion = "%" & Trim(oDoDiagnostico.Descripcion) & "%"
            End If
            
            sWhere = sWhere + " Descripcion like '" + oDoDiagnostico.Descripcion + "' and "
            'sWhere = sWhere + " dbo.cleanSpecialChar(Descripcion) like dbo.cleanSpecialChar('" + oDoDiagnostico.Descripcion + "') and " 'Actualizado Yamill Palomino 16/10/2014
       End If
       If lbSoloMuestraDxGalenHos = True Then
           sWhere = sWhere + " not (descripcionMINSA is null) and "
       End If
       
       If MostrarSoloActivos = True Then
           sWhere = sWhere + " EsActivo = 1 and "
       End If
       
       If sWhere <> "" Then
            sWhere = " Where " & Left(sWhere, Len(sWhere) - 4)
       End If
'SCCQ 03/04/2020 Cambio2 Inicio
sWhere = sWhere + " and CIEHISWEB=1"
'SCCQ 03/04/2020 Cambio2 Fin
       sWhere = sWhere + " order by  Descripcion,CodigoCIE2004 "

       getStringWhereForFilter = sWhere
End Function

Function SeleccionarPorCodigoCIE2004(ByVal oTabla As DODiagnostico, lbSoloMuestraDxGalenHos As Boolean) As Boolean

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sOrder As String
    sSql = ""
    sOrder = ""
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "DiagnosticosSeleccionarPorCodigoCIE2004"
        Set oParameter = .CreateParameter("@CodigoCIE2004", adVarChar, adParamInput, 7, oTabla.CodigoCIE2004): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lbSoloMuestraDxGalenHos", adBoolean, adParamInput, 0, lbSoloMuestraDxGalenHos): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
    If Not (oRecordset.EOF And oRecordset.BOF) Then
        oTabla.IdDiagnostico = oRecordset!IdDiagnostico
        SeleccionarPorCodigoCIE2004 = SeleccionarPorId(oTabla)
    Else
        SeleccionarPorCodigoCIE2004 = False
    End If
    oRecordset.Close
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'mgaray09
Function validarEliminar(ByVal oTabla As DODiagnostico) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New ADODB.Recordset
   
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "DiagnosticosEliminarValidar"
        Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico)): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    
    Set oCommand = Nothing
    Set validarEliminar = oRecordset
    ms_MensajeError = ""
    Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
