VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOEnfermeria_Variables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Enfermeria_Variables
'        Programado por: Cachay F
'        Fecha: Agosto 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idCuentaAtencion As Long
Dim ml_IdVisita As Long
Dim ml_IdVariable As Long
Dim ms_VariableDato As String

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idCuentaAtencion(iValue As Long)
   ml_idCuentaAtencion = iValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property
Property Let IdVisita(iValue As Long)
   ml_IdVisita = iValue
End Property
Property Get IdVisita() As Long
   IdVisita = ml_IdVisita
End Property
Property Let IdVariable(iValue As Long)
   ml_IdVariable = iValue
End Property
Property Get IdVariable() As Long
   IdVariable = ml_IdVariable
End Property
Property Let VariableDato(sValue As String)
   ms_VariableDato = sValue
End Property
Property Get VariableDato() As String
   VariableDato = ms_VariableDato
End Property
