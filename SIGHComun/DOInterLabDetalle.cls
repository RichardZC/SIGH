VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOInterLabDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: InterLabDetalle
'        Autor: Daniel Enrique Barrantes
'        Fecha: 12/05/2018 09:32:48 a.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_identificador As Long
Dim ms_cptCodigo As String
Dim ms_cpt As String
Dim ms_DNI As String
Dim mda_FechaHoraResultado As Date
Dim ms_Grupo As String
Dim ms_item As String
Dim ms_itemResultado As String
Dim ms_itemReferencia As String
Dim ms_itemMetodo As String
Dim ml_idEstado As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let identificador(iValue As Long)
   ml_identificador = iValue
End Property
Property Get identificador() As Long
   identificador = ml_identificador
End Property
Property Let cptCodigo(sValue As String)
   ms_cptCodigo = sValue
End Property
Property Get cptCodigo() As String
   cptCodigo = ms_cptCodigo
End Property
Property Let cpt(sValue As String)
   ms_cpt = sValue
End Property
Property Get cpt() As String
   cpt = ms_cpt
End Property
Property Let DNI(sValue As String)
   ms_DNI = sValue
End Property
Property Get DNI() As String
   DNI = ms_DNI
End Property
Property Let FechaHoraResultado(daValue As Date)
   mda_FechaHoraResultado = daValue
End Property
Property Get FechaHoraResultado() As Date
   FechaHoraResultado = mda_FechaHoraResultado
End Property
Property Let Grupo(sValue As String)
   ms_Grupo = sValue
End Property
Property Get Grupo() As String
   Grupo = ms_Grupo
End Property
Property Let item(sValue As String)
   ms_item = sValue
End Property
Property Get item() As String
   item = ms_item
End Property
Property Let itemResultado(sValue As String)
   ms_itemResultado = sValue
End Property
Property Get itemResultado() As String
   itemResultado = ms_itemResultado
End Property
Property Let itemReferencia(sValue As String)
   ms_itemReferencia = sValue
End Property
Property Get itemReferencia() As String
   itemReferencia = ms_itemReferencia
End Property
Property Let itemMetodo(sValue As String)
   ms_itemMetodo = sValue
End Property
Property Get itemMetodo() As String
   itemMetodo = ms_itemMetodo
End Property
Property Let idEstado(iValue As Long)
   ml_idEstado = iValue
End Property
Property Get idEstado() As Long
   idEstado = ml_idEstado
End Property
