VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoLabGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla LabGrupos
'        Programado por: Bonilla A
'        Fecha: Mayo 2009
'
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idGrupo As Long
Dim ms_NombreGrupo As String
Dim ms_SiglasGrupo As String
Dim ml_idCargo As Long

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let idGrupo(iValue As Long)
   ml_idGrupo = iValue
End Property
Property Get idGrupo() As Long
   idGrupo = ml_idGrupo
End Property
Property Let NombreGrupo(sValue As String)
   ms_NombreGrupo = sValue
End Property
Property Get NombreGrupo() As String
   NombreGrupo = ms_NombreGrupo
End Property
Property Let SiglasGrupo(sValue As String)
   ms_SiglasGrupo = sValue
End Property
Property Get SiglasGrupo() As String
   SiglasGrupo = ms_SiglasGrupo
End Property
Property Let idCargo(iValue As Long)
   ml_idCargo = iValue
End Property
Property Get idCargo() As Long
   idCargo = ml_idCargo
End Property

