VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOImagTipoModalidadSala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla ImagTipoModalidadSala
'        Programado por: Garay M
'        Fecha: Noviembre 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdTipoModalidadSala As Long
Dim ms_TipoModalidadSala As String
Dim ms_Codigo As String
Dim mb_EsActivo As Boolean
Dim mda_FechaCrea As Date
Dim mda_FechaEdita As Date
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdTipoModalidadSala(iValue As Long)
   ml_IdTipoModalidadSala = iValue
End Property
Property Get IdTipoModalidadSala() As Long
   IdTipoModalidadSala = ml_IdTipoModalidadSala
End Property
Property Let TipoModalidadSala(sValue As String)
   ms_TipoModalidadSala = sValue
End Property
Property Get TipoModalidadSala() As String
   TipoModalidadSala = ms_TipoModalidadSala
End Property
Property Let Codigo(sValue As String)
   ms_Codigo = sValue
End Property
Property Get Codigo() As String
   Codigo = ms_Codigo
End Property
Property Let EsActivo(bValue As Boolean)
   mb_EsActivo = bValue
End Property
Property Get EsActivo() As Boolean
   EsActivo = mb_EsActivo
End Property
Property Let FechaCrea(daValue As Date)
   mda_FechaCrea = daValue
End Property
Property Get FechaCrea() As Date
   FechaCrea = mda_FechaCrea
End Property
Property Let FechaEdita(daValue As Date)
   mda_FechaEdita = daValue
End Property
Property Get FechaEdita() As Date
   FechaEdita = mda_FechaEdita
End Property
