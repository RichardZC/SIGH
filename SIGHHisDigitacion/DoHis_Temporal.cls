VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoHis_Temporal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Estructura para la tabla HIS_TEMPORAL
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String
Dim ms_Codigo1 As String
Dim ms_Codigo2 As String
Dim ms_Codigo3 As String
Dim ms_Codigo4 As String
Dim ms_Codigo5 As String
Dim ms_Codigo6 As String
Dim ms_LabConf1 As String
Dim ms_LabConf2 As String
Dim ms_LabConf3 As String
Dim ms_LabConf4 As String
Dim ms_LabConf5 As String
Dim ms_LabConf6 As String
Dim ml_Diagnost1 As Long
Dim ml_Diagnost2 As Long
Dim ml_Diagnost3 As Long
Dim ml_Diagnost4 As Long
Dim ml_Diagnost5 As Long
Dim ml_Diagnost6 As Long
Dim ml_Edad As Long
Dim ml_TipoEdad As Long
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
Property Let Codigo1(sValue As String)
   ms_Codigo1 = sValue
End Property
Property Get Codigo1() As String
   Codigo1 = ms_Codigo1
End Property
Property Let Codigo2(sValue As String)
   ms_Codigo2 = sValue
End Property
Property Get Codigo2() As String
   Codigo2 = ms_Codigo2
End Property
Property Let Codigo3(sValue As String)
   ms_Codigo3 = sValue
End Property
Property Get Codigo3() As String
   Codigo3 = ms_Codigo3
End Property
Property Let Codigo4(sValue As String)
   ms_Codigo4 = sValue
End Property
Property Get Codigo4() As String
   Codigo4 = ms_Codigo4
End Property
Property Let Codigo5(sValue As String)
   ms_Codigo5 = sValue
End Property
Property Get Codigo5() As String
   Codigo5 = ms_Codigo5
End Property
Property Let Codigo6(sValue As String)
   ms_Codigo6 = sValue
End Property
Property Get Codigo6() As String
   Codigo6 = ms_Codigo6
End Property
Property Let LabConf1(sValue As String)
   ms_LabConf1 = sValue
End Property
Property Get LabConf1() As String
   LabConf1 = ms_LabConf1
End Property
Property Let LabConf2(sValue As String)
   ms_LabConf2 = sValue
End Property
Property Get LabConf2() As String
   LabConf2 = ms_LabConf2
End Property
Property Let LabConf3(sValue As String)
   ms_LabConf3 = sValue
End Property
Property Get LabConf3() As String
   LabConf3 = ms_LabConf3
End Property
Property Let LabConf4(sValue As String)
   ms_LabConf4 = sValue
End Property
Property Get LabConf4() As String
   LabConf4 = ms_LabConf4
End Property
Property Let LabConf5(sValue As String)
   ms_LabConf5 = sValue
End Property
Property Get LabConf5() As String
   LabConf5 = ms_LabConf5
End Property
Property Let LabConf6(sValue As String)
   ms_LabConf6 = sValue
End Property
Property Get LabConf6() As String
   LabConf6 = ms_LabConf6
End Property
Property Let Diagnost1(iValue As Long)
   ml_Diagnost1 = iValue
End Property
Property Get Diagnost1() As Long
   Diagnost1 = ml_Diagnost1
End Property
Property Let Diagnost2(iValue As Long)
   ml_Diagnost2 = iValue
End Property
Property Get Diagnost2() As Long
   Diagnost2 = ml_Diagnost2
End Property
Property Let Diagnost3(iValue As Long)
   ml_Diagnost3 = iValue
End Property
Property Get Diagnost3() As Long
   Diagnost3 = ml_Diagnost3
End Property
Property Let Diagnost4(iValue As Long)
   ml_Diagnost4 = iValue
End Property
Property Get Diagnost4() As Long
   Diagnost4 = ml_Diagnost4
End Property
Property Let Diagnost5(iValue As Long)
   ml_Diagnost5 = iValue
End Property
Property Get Diagnost5() As Long
   Diagnost5 = ml_Diagnost5
End Property
Property Let Diagnost6(iValue As Long)
   ml_Diagnost6 = iValue
End Property
Property Get Diagnost6() As Long
   Diagnost6 = ml_Diagnost6
End Property
Property Let Edad(iValue As Long)
   ml_Edad = iValue
End Property
Property Get Edad() As Long
   Edad = ml_Edad
End Property
Property Let TipoEdad(iValue As Long)
   ml_TipoEdad = iValue
End Property
Property Get TipoEdad() As Long
   TipoEdad = ml_TipoEdad
End Property
