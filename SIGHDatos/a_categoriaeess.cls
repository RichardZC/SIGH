VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_categoriaeess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla a_categoriaeess de la db sis
'        Programado por: Garay M
'        Fecha: Marzo 2015
'
'------------------------------------------------------------------------------------
Option Explicit

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
'   Función:        SeleccionarPorCatAberviatura
'   Descripción:    Selecciona un unico registro de la tabla a_presentaciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorCatAbreviatura(ByVal oTabla As DOa_categoriaeess) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorCatAbreviatura = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "a_categoriaeessSeleccionarPorCatAbreviatura"
     Set oParameter = .CreateParameter("@Cat_Abreviatura", adVarChar, adParamInput, 10, oTabla.cat_Abreviatura)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.cat_IdCategoriaEESS = IIf(IsNull(oRecordset!cat_IdCategoriaEESS), "", oRecordset!cat_IdCategoriaEESS)
       oTabla.cat_Descripcion = IIf(IsNull(oRecordset!cat_Descripcion), "", oRecordset!cat_Descripcion)
       oTabla.cat_Abreviatura = IIf(IsNull(oRecordset!cat_Abreviatura), "", oRecordset!cat_Abreviatura)
       oTabla.cat_Nivel = IIf(IsNull(oRecordset!cat_Nivel), 0, oRecordset!cat_Nivel)
       oTabla.cat_IdEstado = oRecordset!cat_IdEstado
        SeleccionarPorCatAbreviatura = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Sub SetDefaults(ByVal oTabla As DOa_categoriaeess)
   oTabla.cat_IdCategoriaEESS = ""
   oTabla.cat_Descripcion = ""
   oTabla.cat_Abreviatura = ""
   oTabla.cat_Nivel = 0
   oTabla.cat_IdEstado = ""
End Sub

Public Function ObtenerValorDefectoIdCategoriaEESS()
    ObtenerValorDefectoIdCategoriaEESS = "00"
End Function
