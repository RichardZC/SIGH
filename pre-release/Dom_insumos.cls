VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dom_insumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla m_insumos
'        Programado por: Barrantes D.
'        Fecha: Mayo 2013
'
'------------------------------------------------------------------------------------

Option Explicit


Dim ml_IdUsuarioAuditoria As Long
Dim ms_ins_CodIns As String
Dim ms_ins_Nombre As String
Dim ms_ins_FormaFarmaceutica As String
Dim ms_ins_Presen As String
Dim ms_ins_Concen As String
Dim mv_ins_Costo As Variant
Dim ms_ins_Observacion As String
Dim ms_ins_Petitorio As String
Dim mda_ins_FecBaja As Date
Dim ms_ins_DocBaja As String
Dim ms_ins_IdEstado As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let ins_CodIns(sValue As String)
   ms_ins_CodIns = sValue
End Property
Property Get ins_CodIns() As String
   ins_CodIns = ms_ins_CodIns
End Property
Property Let ins_Nombre(sValue As String)
   ms_ins_Nombre = sValue
End Property
Property Get ins_Nombre() As String
   ins_Nombre = ms_ins_Nombre
End Property
Property Let ins_FormaFarmaceutica(sValue As String)
   ms_ins_FormaFarmaceutica = sValue
End Property
Property Get ins_FormaFarmaceutica() As String
   ins_FormaFarmaceutica = ms_ins_FormaFarmaceutica
End Property
Property Let ins_Presen(sValue As String)
   ms_ins_Presen = sValue
End Property
Property Get ins_Presen() As String
   ins_Presen = ms_ins_Presen
End Property
Property Let ins_Concen(sValue As String)
   ms_ins_Concen = sValue
End Property
Property Get ins_Concen() As String
   ins_Concen = ms_ins_Concen
End Property
Property Let ins_Costo(vValue As Variant)
   mv_ins_Costo = vValue
End Property
Property Get ins_Costo() As Variant
   ins_Costo = mv_ins_Costo
End Property
Property Let ins_Observacion(sValue As String)
   ms_ins_Observacion = sValue
End Property
Property Get ins_Observacion() As String
   ins_Observacion = ms_ins_Observacion
End Property
Property Let ins_Petitorio(sValue As String)
   ms_ins_Petitorio = sValue
End Property
Property Get ins_Petitorio() As String
   ins_Petitorio = ms_ins_Petitorio
End Property
Property Let ins_FecBaja(daValue As Date)
   mda_ins_FecBaja = daValue
End Property
Property Get ins_FecBaja() As Date
   ins_FecBaja = mda_ins_FecBaja
End Property
Property Let ins_DocBaja(sValue As String)
   ms_ins_DocBaja = sValue
End Property
Property Get ins_DocBaja() As String
   ins_DocBaja = ms_ins_DocBaja
End Property
Property Let ins_IdEstado(sValue As String)
   ms_ins_IdEstado = sValue
End Property
Property Get ins_IdEstado() As String
   ins_IdEstado = ms_ins_IdEstado
End Property

