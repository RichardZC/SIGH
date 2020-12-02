VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoPadronNominal_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla PadronNominal_Detalle
'        Programado por: Palomino Y
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String
Dim ml_IdPaNomDetalle As Long
Dim ml_IdTipoDoc As Long
Dim ml_NumDocumento As Long
Dim ms_HistClinica As String
Dim ms_ApellidoPaterno As String
Dim ms_ApellidoMaterno As String
Dim ms_Nombres As String
Dim ml_idSexo As Long
Dim ms_FecNacimiento As String
Dim ml_IdTipoSeguro As Long
Dim ms_NumAfiliacion As String
Dim ms_FecEvaluacion As String
Dim ms_Peso As String
Dim ms_Talla As String
Dim ml_IdDiagNutricional As Long
Dim ml_CodRenaes As Long
Dim ml_IdDiagPE As Long
Dim ml_IdDiagPT As Long
Dim ml_IdDiagTE As Long
Dim ml_Hemoglobina  As Long
Dim ml_Heces As String

Property Let Heces(lValue As String)
   ml_Heces = lValue
End Property
Property Get Heces() As String
   Heces = ml_Heces
End Property
Property Let Hemoglobina(lValue As Long)
   ml_Hemoglobina = lValue
End Property
Property Get Hemoglobina() As Long
   Hemoglobina = ml_Hemoglobina
End Property
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let IdPaNomDetalle(iValue As Long)
   ml_IdPaNomDetalle = iValue
End Property
Property Get IdPaNomDetalle() As Long
   IdPaNomDetalle = ml_IdPaNomDetalle
End Property
Property Let IdTipoDoc(iValue As Long)
   ml_IdTipoDoc = iValue
End Property
Property Get IdTipoDoc() As Long
   IdTipoDoc = ml_IdTipoDoc
End Property
Property Let NumDocumento(iValue As Long)
   ml_NumDocumento = iValue
End Property
Property Get NumDocumento() As Long
   NumDocumento = ml_NumDocumento
End Property
Property Let HistClinica(sValue As String)
   ms_HistClinica = sValue
End Property
Property Get HistClinica() As String
   HistClinica = ms_HistClinica
End Property
Property Let ApellidoPaterno(sValue As String)
   ms_ApellidoPaterno = sValue
End Property
Property Get ApellidoPaterno() As String
   ApellidoPaterno = ms_ApellidoPaterno
End Property
Property Let ApellidoMaterno(sValue As String)
   ms_ApellidoMaterno = sValue
End Property
Property Get ApellidoMaterno() As String
   ApellidoMaterno = ms_ApellidoMaterno
End Property
Property Let Nombres(sValue As String)
   ms_Nombres = sValue
End Property
Property Get Nombres() As String
   Nombres = ms_Nombres
End Property
Property Let idSexo(iValue As Long)
   ml_idSexo = iValue
End Property
Property Get idSexo() As Long
   idSexo = ml_idSexo
End Property
Property Let FecNacimiento(sValue As String)
   ms_FecNacimiento = sValue
End Property
Property Get FecNacimiento() As String
   FecNacimiento = ms_FecNacimiento
End Property
Property Let IdTipoSeguro(iValue As Long)
   ml_IdTipoSeguro = iValue
End Property
Property Get IdTipoSeguro() As Long
   IdTipoSeguro = ml_IdTipoSeguro
End Property
Property Let NumAfiliacion(sValue As String)
   ms_NumAfiliacion = sValue
End Property
Property Get NumAfiliacion() As String
   NumAfiliacion = ms_NumAfiliacion
End Property
Property Let FecEvaluacion(sValue As String)
   ms_FecEvaluacion = sValue
End Property
Property Get FecEvaluacion() As String
   FecEvaluacion = ms_FecEvaluacion
End Property
Property Let Peso(sValue As String)
   ms_Peso = sValue
End Property
Property Get Peso() As String
   Peso = ms_Peso
End Property
Property Let Talla(sValue As String)
   ms_Talla = sValue
End Property
Property Get Talla() As String
   Talla = ms_Talla
End Property
Property Let IdDiagNutricional(iValue As Long)
   ml_IdDiagNutricional = iValue
End Property
Property Get IdDiagNutricional() As Long
   IdDiagNutricional = ml_IdDiagNutricional
End Property

Property Let CodRenaes(iValue As Long)
   ml_CodRenaes = iValue
End Property
Property Get CodRenaes() As Long
   CodRenaes = ml_CodRenaes
End Property
Property Let IdDiagPE(iValue As Long)
   ml_IdDiagPE = iValue
End Property
Property Get IdDiagPE() As Long
   IdDiagPE = ml_IdDiagPE
End Property
Property Let IdDiagPT(iValue As Long)
   ml_IdDiagPT = iValue
End Property
Property Get IdDiagPT() As Long
   IdDiagPT = ml_IdDiagPT
End Property
Property Let IdDiagTE(iValue As Long)
   ml_IdDiagTE = iValue
End Property
Property Get IdDiagTE() As Long
   IdDiagTE = ml_IdDiagTE
End Property
