VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtenInteHCRespuestaPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtenInteGrupoHCRespuestaPaciente
'        Programado por: Garay M
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdGrupoHCPaciente As Long
Dim ml_IdPaciente As Long
Dim ml_ItemRespuesta As Long
Dim ms_ValorTexto As String
Dim mc_ValorNumero As Currency
Dim mda_ValorFecha As Date
Dim mc_ValorNumeroFin As Currency
Dim mda_ValorFechaFin As Date
Dim ms_ValorEspecificacion As String
Dim mb_EsActivo As Boolean
'Comodin no usar en transacciones con la entidad de dominio
Dim ml_IdPregunta As Long

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let IdGrupoHCPaciente(iValue As Long)
   ml_IdGrupoHCPaciente = iValue
End Property
Property Get IdGrupoHCPaciente() As Long
   IdGrupoHCPaciente = ml_IdGrupoHCPaciente
End Property
Property Let IdPaciente(iValue As Long)
   ml_IdPaciente = iValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let ItemRespuesta(iValue As Long)
   ml_ItemRespuesta = iValue
End Property
Property Get ItemRespuesta() As Long
   ItemRespuesta = ml_ItemRespuesta
End Property
Property Let ValorTexto(sValue As String)
   ms_ValorTexto = sValue
End Property
Property Get ValorTexto() As String
   ValorTexto = ms_ValorTexto
End Property
Property Let ValorNumero(cValue As Currency)
   mc_ValorNumero = cValue
End Property
Property Get ValorNumero() As Currency
   ValorNumero = mc_ValorNumero
End Property
Property Let ValorFecha(daValue As Date)
   mda_ValorFecha = daValue
End Property
Property Get ValorFecha() As Date
   ValorFecha = mda_ValorFecha
End Property
Property Let ValorNumeroFin(cValue As Currency)
   mc_ValorNumeroFin = cValue
End Property
Property Get ValorNumeroFin() As Currency
   ValorNumeroFin = mc_ValorNumeroFin
End Property
Property Let ValorFechaFin(daValue As Date)
   mda_ValorFechaFin = daValue
End Property
Property Get ValorFechaFin() As Date
   ValorFechaFin = mda_ValorFechaFin
End Property
Property Let ValorEspecificacion(sValue As String)
   ms_ValorEspecificacion = sValue
End Property
Property Get ValorEspecificacion() As String
   ValorEspecificacion = ms_ValorEspecificacion
End Property
Property Let EsActivo(bValue As Boolean)
   mb_EsActivo = bValue
End Property
Property Get EsActivo() As Boolean
   EsActivo = mb_EsActivo
End Property

Property Let IdPregunta(iValue As Long)
   ml_IdPregunta = iValue
End Property
Property Get IdPregunta() As Long
   IdPregunta = ml_IdPregunta
End Property
