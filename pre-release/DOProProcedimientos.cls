VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOProProcedimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Procedimientos
'        Programado por: Castro W
'        Fecha: Diciembre 2004
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdPrograma As Long
Dim ml_IdProCabecera As Long
Dim ml_IdControl As Long
Dim ml_IdDiagnostico As Long
Dim ml_IdProducto As Long
Dim ml_IdResultado As Long
Dim mc_labConfHIS As String

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdPrograma(iValue As Long)
   ml_IdPrograma = iValue
End Property
Property Get IdPrograma() As Long
   IdPrograma = ml_IdPrograma
End Property
Property Let IdProCabecera(iValue As Long)
   ml_IdProCabecera = iValue
End Property
Property Get IdProCabecera() As Long
   IdProCabecera = ml_IdProCabecera
End Property
Property Let IdControl(iValue As Long)
   ml_IdControl = iValue
End Property
Property Get IdControl() As Long
   IdControl = ml_IdControl
End Property
Property Let IdDiagnostico(iValue As Long)
   ml_IdDiagnostico = iValue
End Property
Property Get IdDiagnostico() As Long
   IdDiagnostico = ml_IdDiagnostico
End Property
Property Let IdProducto(iValue As Long)
   ml_IdProducto = iValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property
Property Let IdResultado(iValue As Long)
   ml_IdResultado = iValue
End Property
Property Get IdResultado() As Long
   IdResultado = ml_IdResultado
End Property
Property Let labConfHIS(iValue As String)
   mc_labConfHIS = iValue
End Property
Property Get labConfHIS() As String
   labConfHIS = mc_labConfHIS
End Property
