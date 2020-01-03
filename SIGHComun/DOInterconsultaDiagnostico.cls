VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOInterconsultaDiagnostico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla InterconsultasDiagnosticos
'        Programado por: Castro W
'        Fecha: Abril 2005
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdInterconsultaDiagnostico As Long
Dim ml_IdSubClasificacionDX As Long
Dim ml_IdClasificacionDx As Long
Dim ml_IdDiagnostico As Long
Dim ml_IdInterconsulta As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdInterconsultaDiagnostico(lValue As Long)
   ml_IdInterconsultaDiagnostico = lValue
End Property
Property Get IdInterconsultaDiagnostico() As Long
   IdInterconsultaDiagnostico = ml_IdInterconsultaDiagnostico
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
Property Let IdInterconsulta(lValue As Long)
   ml_IdInterconsulta = lValue
End Property
Property Get IdInterconsulta() As Long
   IdInterconsulta = ml_IdInterconsulta
End Property


