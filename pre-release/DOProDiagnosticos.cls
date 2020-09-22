VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOProDiagnosticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla ProDiagnosticos
'        Programado por: Barrantes D
'        Fecha: Diciembre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdPrograma As Long
Dim ml_IdProCabecera As Long
Dim ml_IdControl As Long
Dim ml_IdDiagnostico As Long
Dim mb_Principal As Boolean
Dim mc_labConfHIS As String
Dim ml_IdSubClasificacionDX As Long

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
Property Let Principal(bValue As Boolean)
   mb_Principal = bValue
End Property
Property Get Principal() As Boolean
   Principal = mb_Principal
End Property
Property Let labConfHIS(bValue As String)
   mc_labConfHIS = bValue
End Property
Property Get labConfHIS() As String
   labConfHIS = mc_labConfHIS
End Property
Property Let IdSubClasificacionDX(lValue As Long)
   ml_IdSubClasificacionDX = lValue
End Property
Property Get IdSubClasificacionDX() As Long
   IdSubClasificacionDX = ml_IdSubClasificacionDX
End Property

