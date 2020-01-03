VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtencionDiagnostico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtencionesDiagnostico
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------

Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdSubClasificacionDX As Long
Dim ml_IdClasificacionDx As Long
Dim ml_IdDiagnostico As Long
Dim ml_IdAtencionDiagnostico As Long
Dim ml_IdAtencion As Long
Dim ms_LabConfHIS As String
Dim ml_GrupoHIS As Long
Dim ml_SubGrupoHIS As Long

Property Get GrupoHIS() As Long
   GrupoHIS = ml_GrupoHIS
End Property
Property Let GrupoHIS(lValue As Long)
   ml_GrupoHIS = lValue
End Property
Property Get SubGrupoHIS() As Long
   SubGrupoHIS = ml_SubGrupoHIS
End Property
Property Let SubGrupoHIS(lValue As Long)
   ml_SubGrupoHIS = lValue
End Property

Property Get labConfHIS() As String
   labConfHIS = ms_LabConfHIS
End Property
Property Let labConfHIS(sValue As String)
   ms_LabConfHIS = sValue
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdSubClasificacionDX(lValue As Long)
   ml_IdSubClasificacionDX = lValue
End Property
Property Get IdSubClasificacionDX() As Long
   IdSubClasificacionDX = ml_IdSubClasificacionDX
End Property
Property Let IdClasificacionDx(lValue As Long)
   ml_IdClasificacionDx = lValue
End Property
Property Get IdClasificacionDx() As Long
   IdClasificacionDx = ml_IdClasificacionDx
End Property
Property Let IdDiagnostico(lValue As Long)
   ml_IdDiagnostico = lValue
End Property
Property Get IdDiagnostico() As Long
   IdDiagnostico = ml_IdDiagnostico
End Property
Property Let IdAtencionDiagnostico(lValue As Long)
   ml_IdAtencionDiagnostico = lValue
End Property
Property Get IdAtencionDiagnostico() As Long
   IdAtencionDiagnostico = ml_IdAtencionDiagnostico
End Property
Property Let IdAtencion(lValue As Long)
   ml_IdAtencion = lValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property


