VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dom_medicamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla m_medicamentos
'        Programado por: Barrantes D.
'        Fecha: Mayo 2013
'
'------------------------------------------------------------------------------------

Option Explicit


Dim ml_IdUsuarioAuditoria As Long
Dim ms_med_CodMed As String
Dim ms_med_Nombre As String
Dim ms_med_FormaFarmaceutica As String
Dim ms_med_Presen As String
Dim ms_med_Concen As String
Dim mv_med_Costo As Variant
Dim ms_med_Petitorio As String
Dim ms_med_Petitorio2005 As String
Dim ms_med_Petitorio2010 As String
Dim mda_med_FecBaja As Date
Dim ms_med_FFDigemid As String
Dim ms_med_IdEstado As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let med_CodMed(sValue As String)
   ms_med_CodMed = sValue
End Property
Property Get med_CodMed() As String
   med_CodMed = ms_med_CodMed
End Property
Property Let med_Nombre(sValue As String)
   ms_med_Nombre = sValue
End Property
Property Get med_Nombre() As String
   med_Nombre = ms_med_Nombre
End Property
Property Let med_FormaFarmaceutica(sValue As String)
   ms_med_FormaFarmaceutica = sValue
End Property
Property Get med_FormaFarmaceutica() As String
   med_FormaFarmaceutica = ms_med_FormaFarmaceutica
End Property
Property Let med_Presen(sValue As String)
   ms_med_Presen = sValue
End Property
Property Get med_Presen() As String
   med_Presen = ms_med_Presen
End Property
Property Let med_Concen(sValue As String)
   ms_med_Concen = sValue
End Property
Property Get med_Concen() As String
   med_Concen = ms_med_Concen
End Property
Property Let med_Costo(vValue As Variant)
   mv_med_Costo = vValue
End Property
Property Get med_Costo() As Variant
   med_Costo = mv_med_Costo
End Property
Property Let med_Petitorio(sValue As String)
   ms_med_Petitorio = sValue
End Property
Property Get med_Petitorio() As String
   med_Petitorio = ms_med_Petitorio
End Property
Property Let med_Petitorio2005(sValue As String)
   ms_med_Petitorio2005 = sValue
End Property
Property Get med_Petitorio2005() As String
   med_Petitorio2005 = ms_med_Petitorio2005
End Property
Property Let med_Petitorio2010(sValue As String)
   ms_med_Petitorio2010 = sValue
End Property
Property Get med_Petitorio2010() As String
   med_Petitorio2010 = ms_med_Petitorio2010
End Property
Property Let med_FecBaja(daValue As Date)
   mda_med_FecBaja = daValue
End Property
Property Get med_FecBaja() As Date
   med_FecBaja = mda_med_FecBaja
End Property
Property Let med_FFDigemid(sValue As String)
   ms_med_FFDigemid = sValue
End Property
Property Get med_FFDigemid() As String
   med_FFDigemid = ms_med_FFDigemid
End Property
Property Let med_IdEstado(sValue As String)
   ms_med_IdEstado = sValue
End Property
Property Get med_IdEstado() As String
   med_IdEstado = ms_med_IdEstado
End Property



