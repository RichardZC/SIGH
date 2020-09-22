VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOMedico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Medicos
'        Programado por: Barrantes D.
'        Fecha: Mayo 2013
'
'------------------------------------------------------------------------------------

Option Explicit


Dim ml_Auditoria As Long
Dim ml_IdEmpleado As Long
Dim ms_Colegiatura As String
Dim ml_idMedico As Long
Dim ms_LoteHis As String
Dim ms_idColegioHis As String
Dim ms_rne As String
Dim ms_egresado As Boolean
Property Let Egresado(sValue As Boolean)
   ms_egresado = sValue
End Property
Property Get Egresado() As Boolean
   Egresado = ms_egresado
End Property
Property Let Rne(sValue As String)
   ms_rne = sValue
End Property
Property Get Rne() As String
   Rne = ms_rne
End Property
Property Let idColegioHis(sValue As String)
   ms_idColegioHis = sValue
End Property
Property Get idColegioHis() As String
   idColegioHis = ms_idColegioHis
End Property
Property Let LoteHis(sValue As String)
   ms_LoteHis = sValue
End Property
Property Get LoteHis() As String
   LoteHis = ms_LoteHis
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdEmpleado(lValue As Long)
   ml_IdEmpleado = lValue
End Property
Property Get IdEmpleado() As Long
   IdEmpleado = ml_IdEmpleado
End Property
Property Let Colegiatura(sValue As String)
   ms_Colegiatura = sValue
End Property
Property Get Colegiatura() As String
   Colegiatura = ms_Colegiatura
End Property
Property Let idMedico(lValue As Long)
   ml_idMedico = lValue
End Property
Property Get idMedico() As Long
   idMedico = ml_idMedico
End Property


