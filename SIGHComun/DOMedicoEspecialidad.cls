VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOMedicoEspecialidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla MedicosEspecialidad
'        Programado por: Barrantes D.
'        Fecha: Mayo 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_idMedico As Long
Dim ml_IdEspecialidad As Long
Dim ml_IdMedicoEspecialidad As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let idMedico(lValue As Long)
   ml_idMedico = lValue
End Property
Property Get idMedico() As Long
   idMedico = ml_idMedico
End Property
Property Let IdEspecialidad(lValue As Long)
   ml_IdEspecialidad = lValue
End Property
Property Get IdEspecialidad() As Long
   IdEspecialidad = ml_IdEspecialidad
End Property
Property Let IdMedicoEspecialidad(lValue As Long)
   ml_IdMedicoEspecialidad = lValue
End Property
Property Get IdMedicoEspecialidad() As Long
   IdMedicoEspecialidad = ml_IdMedicoEspecialidad
End Property


