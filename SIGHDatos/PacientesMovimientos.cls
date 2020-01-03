VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PacientesMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - HFG
'        Aplicativo: SisGalenPlus v.3
'        Programa: código autogenerado para la clase: PacientesMovimientos
'        Programado por: Daniel Enrique Barrantes
'        Fecha: 27/02/2015 04:42:07 p.m.
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
'   Descripción:    Inserta un registro a la tabla PacientesMovimientos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOPacienteMovimientos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PacientesMovimientosAgregar"
   Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Peso", adCurrency, adParamInput, 0, oTabla.Peso)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Talla", adCurrency, adParamInput, 0, oTabla.Talla)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.idDxNutricional = 0, Null, oTabla.idDxNutricional))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrafXedadEnMeses", adInteger, adParamInput, 0, IIf(oTabla.GrafXedadEnMeses = 0, Null, oTabla.GrafXedadEnMeses))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrafYpercentilTE", adInteger, adParamInput, 0, IIf(oTabla.GrafYpercentilTE = 0, Null, oTabla.GrafYpercentilTE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrafYpercentilPT", adInteger, adParamInput, 0, IIf(oTabla.GrafYpercentilPT = 0, Null, oTabla.GrafYpercentilPT))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrafYpercentilPE", adInteger, adParamInput, 0, IIf(oTabla.GrafYpercentilPE = 0, Null, oTabla.GrafYpercentilPE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ZetaPT", adCurrency, adParamInput, 0, oTabla.ZetaPT)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ZetaTE", adCurrency, adParamInput, 0, oTabla.ZetaTE)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ZetaPE", adCurrency, adParamInput, 0, oTabla.ZetaPE)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Hemoglobina", adCurrency, adParamInput, 0, IIf(oTabla.Hemoglobina = 0, Null, oTabla.Hemoglobina))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Parasitosis", adVarChar, adParamInput, 2, IIf(oTabla.Parasitosis = "", Null, oTabla.Parasitosis))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
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
'   Descripción:    modifica un registro a la tabla PacientesMovimientos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOPacienteMovimientos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PacientesMovimientosModificar"
   Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Peso", adCurrency, adParamInput, 0, oTabla.Peso)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Talla", adCurrency, adParamInput, 0, oTabla.Talla)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.idDxNutricional = 0, Null, oTabla.idDxNutricional))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrafXedadEnMeses", adInteger, adParamInput, 0, IIf(oTabla.GrafXedadEnMeses = 0, Null, oTabla.GrafXedadEnMeses))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrafYpercentilTE", adInteger, adParamInput, 0, IIf(oTabla.GrafYpercentilTE = 0, Null, oTabla.GrafYpercentilTE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrafYpercentilPT", adInteger, adParamInput, 0, IIf(oTabla.GrafYpercentilPT = 0, Null, oTabla.GrafYpercentilPT))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrafYpercentilPE", adInteger, adParamInput, 0, IIf(oTabla.GrafYpercentilPE = 0, Null, oTabla.GrafYpercentilPE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ZetaPT", adCurrency, adParamInput, 0, oTabla.ZetaPT)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ZetaTE", adCurrency, adParamInput, 0, oTabla.ZetaTE)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ZetaPE", adCurrency, adParamInput, 0, oTabla.ZetaPE)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Hemoglobina", adCurrency, adParamInput, 0, IIf(oTabla.Hemoglobina = 0, Null, oTabla.Hemoglobina))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Parasitosis", adVarChar, adParamInput, 2, IIf(oTabla.Parasitosis = "", Null, oTabla.Parasitosis))
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
'   Descripción:    elimia o borra un registro a la tabla PacientesMovimientos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOPacienteMovimientos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PacientesMovimientosEliminar"
   Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oTabla.IdCuentaAtencion)
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
'   Descripción:    Selecciona un unico registro de la tabla PacientesMovimientos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOPacienteMovimientos) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "PacientesMovimientosSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oTabla.IdCuentaAtencion)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.Peso = oRecordset!Peso
       oTabla.Talla = oRecordset!Talla
       oTabla.idDxNutricional = IIf(IsNull(oRecordset!idDxNutricional), 0, oRecordset!idDxNutricional)
       oTabla.GrafXedadEnMeses = IIf(IsNull(oRecordset!GrafXedadEnMeses), 0, oRecordset!GrafXedadEnMeses)
       oTabla.GrafYpercentilTE = IIf(IsNull(oRecordset!GrafYpercentilTE), 0, oRecordset!GrafYpercentilTE)
       oTabla.GrafYpercentilPT = IIf(IsNull(oRecordset!GrafYpercentilPT), 0, oRecordset!GrafYpercentilPT)
       oTabla.GrafYpercentilPE = IIf(IsNull(oRecordset!GrafYpercentilPE), 0, oRecordset!GrafYpercentilPE)
       oTabla.ZetaPT = oRecordset!ZetaPT
       oTabla.ZetaTE = oRecordset!ZetaTE
       oTabla.ZetaPE = oRecordset!ZetaPE
       oTabla.Hemoglobina = IIf(IsNull(oRecordset!Hemoglobina), 0, oRecordset!Hemoglobina)
       oTabla.Parasitosis = IIf(IsNull(oRecordset!Parasitosis), "", oRecordset!Parasitosis)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOPacienteMovimientos)
   oTabla.IdCuentaAtencion = 0
   oTabla.Peso = 0
   oTabla.Talla = 0
   oTabla.idDxNutricional = 0
   oTabla.GrafXedadEnMeses = 0
   oTabla.GrafYpercentilTE = 0
   oTabla.GrafYpercentilPT = 0
   oTabla.GrafYpercentilPE = 0
   oTabla.ZetaPT = 0
   oTabla.ZetaTE = 0
   oTabla.ZetaPE = 0
   oTabla.Hemoglobina = 0
   oTabla.Parasitosis = ""
End Sub








'Function PacientesMovimientosActualizar(ByVal oTabla As DOPacienteMovimientos, lbEsEliminarRegistro As Boolean) As Boolean
'On Error GoTo ManejadorDeError
'Dim oCommand As New ADODB.Command
'Dim oParameter As ADODB.Parameter
'
'PacientesMovimientosActualizar = False
'With oCommand
'   .CommandType = adCmdStoredProc
'   Set .ActiveConnection = mo_Conexion
'   .CommandText = "PacientesMovimientosActualizar"
'   Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, oTabla.Fecha)
'   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente)
'   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@idListBarItem", adInteger, adParamInput, 0, oTabla.idListBarItem)
'   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@Peso", adCurrency, adParamInput, 0, oTabla.Peso)
'   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@Talla", adCurrency, adParamInput, 0, oTabla.Talla)
'   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@idDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.idDiagnostico = 0, Null, oTabla.idDiagnostico))
'   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@GrafXedadEnMeses", adInteger, adParamInput, 0, IIf(oTabla.GrafXedadEnMeses = 0, Null, oTabla.GrafXedadEnMeses))
'   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@GrafYpercentilTE", adInteger, adParamInput, 0, IIf(oTabla.GrafYpercentilTE = 0, Null, oTabla.GrafYpercentilTE))
'   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@GrafYpercentilPT", adInteger, adParamInput, 0, IIf(oTabla.GrafYpercentilPT = 0, Null, oTabla.GrafYpercentilPT))
'   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@GrafYpercentilPE", adInteger, adParamInput, 0, IIf(oTabla.GrafYpercentilPE = 0, Null, oTabla.GrafYpercentilPE))
'   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@ZetaPT", adCurrency, adParamInput, 0, oTabla.ZetaPT)
'   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@ZetaTE", adCurrency, adParamInput, 0, oTabla.ZetaTE)
'   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@ZetaPE", adCurrency, adParamInput, 0, oTabla.ZetaPE)
'   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@Hemoglobina", adCurrency, adParamInput, 0, IIf(oTabla.Hemoglobina = 0, Null, oTabla.Hemoglobina))
'   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@Parasitosis", adVarChar, adParamInput, 2, IIf(oTabla.Parasitosis = "", Null, oTabla.Parasitosis))
'   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@esEliminarRegistro", adInteger, adParamInput, 0, lbEsEliminarRegistro)
'   .Parameters.Append oParameter
'   .Execute
'End With
'
'PacientesMovimientosActualizar = True
'ms_MensajeError = ""
'
'Exit Function
'ManejadorDeError:
'   ms_MensajeError = Err.Number & " " + Err.Description
'   'Resume
'Exit Function
'End Function
