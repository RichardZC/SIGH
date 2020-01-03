VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SunatParteDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: Sunat_ResumenDiario
'        Autor: Daniel Enrique Barrantes
'        Fecha: 28/02/2019 07:04:40 p.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim mo_Conexion As ADODB.Connection
Dim ms_MensajeError As String
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla Sunat_ResumenDiario
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoSunatResumenDia) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "Sunat_ResumenDiarioAgregar"
   Set oParameter = .CreateParameter("@id", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DctoRDI", adVarChar, adParamInput, 50, IIf(oTabla.DctoRDI = "", Null, oTabla.DctoRDI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fechaEmision", adVarChar, adParamInput, 10, IIf(oTabla.fechaEmision = "", Null, oTabla.fechaEmision))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fechaResumen", adVarChar, adParamInput, 10, IIf(oTabla.fechaResumen = "", Null, oTabla.fechaResumen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@tipoDcto", adVarChar, adParamInput, 2, IIf(oTabla.tipoDcto = "", Null, oTabla.tipoDcto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SerieDocumento", adVarChar, adParamInput, 15, IIf(oTabla.SerieDocumento = "", Null, oTabla.SerieDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteTipoDoc", adVarChar, adParamInput, 1, IIf(oTabla.PacienteTipoDoc = "", Null, oTabla.PacienteTipoDoc))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteNumeroDoc", adVarChar, adParamInput, 8, IIf(oTabla.PacienteNumeroDoc = "", Null, oTabla.PacienteNumeroDoc))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Moneda", adVarChar, adParamInput, 3, IIf(oTabla.Moneda = "", Null, oTabla.Moneda))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OpeGravadas", adVarChar, adParamInput, 15, IIf(oTabla.OpeGravadas = "", Null, oTabla.OpeGravadas))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OpeExoneradas", adVarChar, adParamInput, 15, IIf(oTabla.OpeExoneradas = "", Null, oTabla.OpeExoneradas))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OpeInafectas", adVarChar, adParamInput, 15, IIf(oTabla.OpeInafectas = "", Null, oTabla.OpeInafectas))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OpeGratuitas", adVarChar, adParamInput, 15, IIf(oTabla.OpeGratuitas = "", Null, oTabla.OpeGratuitas))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OtrosCargos", adVarChar, adParamInput, 15, IIf(oTabla.OtrosCargos = "", Null, oTabla.OtrosCargos))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TotalIsc", adVarChar, adParamInput, 15, IIf(oTabla.TotalIsc = "", Null, oTabla.TotalIsc))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TotalIgv", adVarChar, adParamInput, 15, IIf(oTabla.TotalIgv = "", Null, oTabla.TotalIgv))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TotalOtros", adVarChar, adParamInput, 15, IIf(oTabla.TotalOtros = "", Null, oTabla.TotalOtros))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ImporteVenta", adVarChar, adParamInput, 15, IIf(oTabla.ImporteVenta = "", Null, oTabla.ImporteVenta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ModTipoDcto", adVarChar, adParamInput, 5, IIf(oTabla.ModTipoDcto = "", Null, oTabla.ModTipoDcto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ModSerie", adVarChar, adParamInput, 5, IIf(oTabla.ModSerie = "", Null, oTabla.ModSerie))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ModDocumento", adVarChar, adParamInput, 10, oTabla.ModDocumento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PercepRegimen", adVarChar, adParamInput, 15, IIf(oTabla.PercepRegimen = "", Null, oTabla.PercepRegimen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PercepPorcen", adVarChar, adParamInput, 15, IIf(oTabla.PercepPorcen = "", Null, oTabla.PercepPorcen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PercepBaseImp", adVarChar, adParamInput, 15, IIf(oTabla.PercepBaseImp = "", Null, oTabla.PercepBaseImp))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PercepMonto", adVarChar, adParamInput, 15, IIf(oTabla.PercepMonto = "", Null, oTabla.PercepMonto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PercepCobrar", adVarChar, adParamInput, 15, IIf(oTabla.PercepCobrar = "", Null, oTabla.PercepCobrar))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Estado", adVarChar, adParamInput, 1, IIf(oTabla.Estado = "", Null, oTabla.Estado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DctoSunat", adVarChar, adParamInput, 50, IIf(oTabla.DctoSunat = "", Null, oTabla.DctoSunat))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.Id = .Parameters("@id")
End With
 
Insertar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla Sunat_ResumenDiario
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoSunatResumenDia) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "Sunat_ResumenDiarioModificar"
   Set oParameter = .CreateParameter("@id", adInteger, adParamInput, 0, IIf(oTabla.Id = 0, Null, oTabla.Id))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DctoRDI", adVarChar, adParamInput, 50, IIf(oTabla.DctoRDI = "", Null, oTabla.DctoRDI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fechaEmision", adVarChar, adParamInput, 10, IIf(oTabla.fechaEmision = "", Null, oTabla.fechaEmision))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fechaResumen", adVarChar, adParamInput, 10, IIf(oTabla.fechaResumen = "", Null, oTabla.fechaResumen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@tipoDcto", adVarChar, adParamInput, 2, IIf(oTabla.tipoDcto = "", Null, oTabla.tipoDcto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SerieDocumento", adVarChar, adParamInput, 15, IIf(oTabla.SerieDocumento = "", Null, oTabla.SerieDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteTipoDoc", adVarChar, adParamInput, 1, IIf(oTabla.PacienteTipoDoc = "", Null, oTabla.PacienteTipoDoc))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PacienteNumeroDoc", adVarChar, adParamInput, 8, IIf(oTabla.PacienteNumeroDoc = "", Null, oTabla.PacienteNumeroDoc))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Moneda", adVarChar, adParamInput, 3, IIf(oTabla.Moneda = "", Null, oTabla.Moneda))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OpeGravadas", adVarChar, adParamInput, 15, IIf(oTabla.OpeGravadas = "", Null, oTabla.OpeGravadas))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OpeExoneradas", adVarChar, adParamInput, 15, IIf(oTabla.OpeExoneradas = "", Null, oTabla.OpeExoneradas))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OpeInafectas", adVarChar, adParamInput, 15, IIf(oTabla.OpeInafectas = "", Null, oTabla.OpeInafectas))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OpeGratuitas", adVarChar, adParamInput, 15, IIf(oTabla.OpeGratuitas = "", Null, oTabla.OpeGratuitas))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OtrosCargos", adVarChar, adParamInput, 15, IIf(oTabla.OtrosCargos = "", Null, oTabla.OtrosCargos))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TotalIsc", adVarChar, adParamInput, 15, IIf(oTabla.TotalIsc = "", Null, oTabla.TotalIsc))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TotalIgv", adVarChar, adParamInput, 15, IIf(oTabla.TotalIgv = "", Null, oTabla.TotalIgv))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TotalOtros", adVarChar, adParamInput, 15, IIf(oTabla.TotalOtros = "", Null, oTabla.TotalOtros))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ImporteVenta", adVarChar, adParamInput, 15, IIf(oTabla.ImporteVenta = "", Null, oTabla.ImporteVenta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ModTipoDcto", adVarChar, adParamInput, 5, IIf(oTabla.ModTipoDcto = "", Null, oTabla.ModTipoDcto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ModSerie", adVarChar, adParamInput, 5, IIf(oTabla.ModSerie = "", Null, oTabla.ModSerie))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ModDocumento", adVarChar, adParamInput, 10, oTabla.ModDocumento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PercepRegimen", adVarChar, adParamInput, 15, IIf(oTabla.PercepRegimen = "", Null, oTabla.PercepRegimen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PercepPorcen", adVarChar, adParamInput, 15, IIf(oTabla.PercepPorcen = "", Null, oTabla.PercepPorcen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PercepBaseImp", adVarChar, adParamInput, 15, IIf(oTabla.PercepBaseImp = "", Null, oTabla.PercepBaseImp))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PercepMonto", adVarChar, adParamInput, 15, IIf(oTabla.PercepMonto = "", Null, oTabla.PercepMonto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PercepCobrar", adVarChar, adParamInput, 15, IIf(oTabla.PercepCobrar = "", Null, oTabla.PercepCobrar))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Estado", adVarChar, adParamInput, 1, IIf(oTabla.Estado = "", Null, oTabla.Estado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DctoSunat", adVarChar, adParamInput, 50, IIf(oTabla.DctoSunat = "", Null, oTabla.DctoSunat))
   .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla Sunat_ResumenDiario
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoSunatResumenDia) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "Sunat_ResumenDiarioEliminar"
   Set oParameter = .CreateParameter("@id", adInteger, adParamInput, 0, oTabla.Id)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Eliminar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla Sunat_ResumenDiario
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoSunatResumenDia) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "Sunat_ResumenDiarioSeleccionarPorId"
     Set oParameter = .CreateParameter("@id", adInteger, adParamInput, 0, oTabla.Id)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.Id = IIf(IsNull(oRecordset!Id), 0, oRecordset!Id)
       oTabla.DctoRDI = IIf(IsNull(oRecordset!DctoRDI), "", oRecordset!DctoRDI)
       oTabla.fechaEmision = IIf(IsNull(oRecordset!fechaEmision), "", oRecordset!fechaEmision)
       oTabla.fechaResumen = IIf(IsNull(oRecordset!fechaResumen), "", oRecordset!fechaResumen)
       oTabla.tipoDcto = IIf(IsNull(oRecordset!tipoDcto), "", oRecordset!tipoDcto)
       oTabla.SerieDocumento = IIf(IsNull(oRecordset!SerieDocumento), "", oRecordset!SerieDocumento)
       oTabla.PacienteTipoDoc = IIf(IsNull(oRecordset!PacienteTipoDoc), "", oRecordset!PacienteTipoDoc)
       oTabla.PacienteNumeroDoc = IIf(IsNull(oRecordset!PacienteNumeroDoc), "", oRecordset!PacienteNumeroDoc)
       oTabla.Moneda = IIf(IsNull(oRecordset!Moneda), "", oRecordset!Moneda)
       oTabla.OpeGravadas = IIf(IsNull(oRecordset!OpeGravadas), "", oRecordset!OpeGravadas)
       oTabla.OpeExoneradas = IIf(IsNull(oRecordset!OpeExoneradas), "", oRecordset!OpeExoneradas)
       oTabla.OpeInafectas = IIf(IsNull(oRecordset!OpeInafectas), "", oRecordset!OpeInafectas)
       oTabla.OpeGratuitas = IIf(IsNull(oRecordset!OpeGratuitas), "", oRecordset!OpeGratuitas)
       oTabla.OtrosCargos = IIf(IsNull(oRecordset!OtrosCargos), "", oRecordset!OtrosCargos)
       oTabla.TotalIsc = IIf(IsNull(oRecordset!TotalIsc), "", oRecordset!TotalIsc)
       oTabla.TotalIgv = IIf(IsNull(oRecordset!TotalIgv), "", oRecordset!TotalIgv)
       oTabla.TotalOtros = IIf(IsNull(oRecordset!TotalOtros), "", oRecordset!TotalOtros)
       oTabla.ImporteVenta = IIf(IsNull(oRecordset!ImporteVenta), "", oRecordset!ImporteVenta)
       oTabla.ModTipoDcto = IIf(IsNull(oRecordset!ModTipoDcto), "", oRecordset!ModTipoDcto)
       oTabla.ModSerie = IIf(IsNull(oRecordset!ModSerie), "", oRecordset!ModSerie)
       oTabla.ModDocumento = IIf(IsNull(oRecordset!ModDocumento), "", oRecordset!ModDocumento)
       oTabla.PercepRegimen = IIf(IsNull(oRecordset!PercepRegimen), "", oRecordset!PercepRegimen)
       oTabla.PercepPorcen = IIf(IsNull(oRecordset!PercepPorcen), "", oRecordset!PercepPorcen)
       oTabla.PercepBaseImp = IIf(IsNull(oRecordset!PercepBaseImp), "", oRecordset!PercepBaseImp)
       oTabla.PercepMonto = IIf(IsNull(oRecordset!PercepMonto), "", oRecordset!PercepMonto)
       oTabla.PercepCobrar = IIf(IsNull(oRecordset!PercepCobrar), "", oRecordset!PercepCobrar)
       oTabla.Estado = IIf(IsNull(oRecordset!Estado), "", oRecordset!Estado)
       oTabla.DctoSunat = IIf(IsNull(oRecordset!DctoSunat), "", oRecordset!DctoSunat)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoSunatResumenDia)
   oTabla.Id = 0
   oTabla.DctoRDI = ""
   oTabla.fechaEmision = ""
   oTabla.fechaResumen = ""
   oTabla.tipoDcto = ""
   oTabla.SerieDocumento = ""
   oTabla.PacienteTipoDoc = ""
   oTabla.PacienteNumeroDoc = ""
   oTabla.Moneda = ""
   oTabla.OpeGravadas = ""
   oTabla.OpeExoneradas = ""
   oTabla.OpeInafectas = ""
   oTabla.OpeGratuitas = ""
   oTabla.OtrosCargos = ""
   oTabla.TotalIsc = ""
   oTabla.TotalIgv = ""
   oTabla.TotalOtros = ""
   oTabla.ImporteVenta = ""
   oTabla.ModTipoDcto = ""
   oTabla.ModSerie = ""
   oTabla.ModDocumento = ""
   oTabla.PercepRegimen = ""
   oTabla.PercepPorcen = ""
   oTabla.PercepBaseImp = ""
   oTabla.PercepMonto = ""
   oTabla.PercepCobrar = ""
   oTabla.Estado = ""
   oTabla.DctoSunat = ""
End Sub
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: Sunat_ResumenDiario
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: Sunat_ResumenDiario
'        Autor:   DAniel Barrantes B.
'        Fecha: 28/02/2019 07:04:45 p.m.
'------------------------------------------------------------------------------------
 

