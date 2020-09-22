VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFarm_FormDetL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: farm_formdetl
'        Autor: Daniel Enrique Barrantes
'        Fecha: 26/04/2019 10:38:37 a.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ms_CODIGO_EJE As String
Dim ms_CODIGO_PRE As String
Dim ms_TIPSUM As String
Dim ms_ANNOMES As String
Dim ms_CODIGO_MED As String
Dim ms_LOTE As String
Dim mda_FECHVTO As Date
Dim md_SALDO As Long
Dim ms_SIT As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let CODIGO_EJE(sValue As String)
   ms_CODIGO_EJE = sValue
End Property
Property Get CODIGO_EJE() As String
   CODIGO_EJE = ms_CODIGO_EJE
End Property
Property Let CODIGO_PRE(sValue As String)
   ms_CODIGO_PRE = sValue
End Property
Property Get CODIGO_PRE() As String
   CODIGO_PRE = ms_CODIGO_PRE
End Property
Property Let TIPSUM(sValue As String)
   ms_TIPSUM = sValue
End Property
Property Get TIPSUM() As String
   TIPSUM = ms_TIPSUM
End Property
Property Let ANNOMES(sValue As String)
   ms_ANNOMES = sValue
End Property
Property Get ANNOMES() As String
   ANNOMES = ms_ANNOMES
End Property
Property Let CODIGO_MED(sValue As String)
   ms_CODIGO_MED = sValue
End Property
Property Get CODIGO_MED() As String
   CODIGO_MED = ms_CODIGO_MED
End Property
Property Let LOTE(sValue As String)
   ms_LOTE = sValue
End Property
Property Get LOTE() As String
   LOTE = ms_LOTE
End Property
Property Let FECHVTO(daValue As Date)
   mda_FECHVTO = daValue
End Property
Property Get FECHVTO() As Date
   FECHVTO = mda_FECHVTO
End Property
Property Let SALDO(dValue As Long)
   md_SALDO = dValue
End Property
Property Get SALDO() As Long
   SALDO = md_SALDO
End Property
Property Let SIT(sValue As String)
   ms_SIT = sValue
End Property
Property Get SIT() As String
   SIT = ms_SIT
End Property

'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: farm_formdetl
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: farm_formdetl
'        Autor:   DAniel Barrantes B.
'        Fecha: 26/04/2019 10:38:37 a.m.
'------------------------------------------------------------------------------------
 

