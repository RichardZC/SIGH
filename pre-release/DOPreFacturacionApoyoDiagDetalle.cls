VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtencionApoyoDiagDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtencionApoyoDiagnotiscoDetalle
'        Programado por: Garay M
'        Fecha: Agosto 2014
'
'------------------------------------------------------------------------------------

Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdFacturacionServicio As Long
Dim ms_HoraResultado As String
Dim mda_FechaResultado As Date
Dim ml_IdServicioRealiza As Long
Dim ml_IdProcedimiento As Long
Dim ml_IdAtencionApoyoDx As Long
Dim ml_IdAtencionApoyoDetalle As Long
Dim ms_EstadoRegistro As String


Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdFacturacionServicio(lValue As Long)
   ml_IdFacturacionServicio = lValue
End Property
Property Get IdFacturacionServicio() As Long
   IdFacturacionServicio = ml_IdFacturacionServicio
End Property
Property Let HoraResultado(sValue As String)
   ms_HoraResultado = sValue
End Property
Property Get HoraResultado() As String
   HoraResultado = ms_HoraResultado
End Property
Property Let FechaResultado(daValue As Date)
   mda_FechaResultado = daValue
End Property
Property Get FechaResultado() As Date
   FechaResultado = mda_FechaResultado
End Property
Property Let IdServicioRealiza(lValue As Long)
   ml_IdServicioRealiza = lValue
End Property
Property Get IdServicioRealiza() As Long
   IdServicioRealiza = ml_IdServicioRealiza
End Property
Property Let IdProcedimiento(lValue As Long)
   ml_IdProcedimiento = lValue
End Property
Property Get IdProcedimiento() As Long
   IdProcedimiento = ml_IdProcedimiento
End Property
Property Let IdAtencionApoyoDx(lValue As Long)
   ml_IdAtencionApoyoDx = lValue
End Property
Property Get IdAtencionApoyoDx() As Long
   IdAtencionApoyoDx = ml_IdAtencionApoyoDx
End Property
Property Let IdAtencionApoyoDetalle(lValue As Long)
   ml_IdAtencionApoyoDetalle = lValue
End Property
Property Get IdAtencionApoyoDetalle() As Long
   IdAtencionApoyoDetalle = ml_IdAtencionApoyoDetalle
End Property

Property Let EstadoRegistro(sValue As String)
   ms_EstadoRegistro = sValue
End Property
Property Get EstadoRegistro() As String
   EstadoRegistro = ms_EstadoRegistro
End Property
