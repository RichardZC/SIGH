VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DODiagnostico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Diagnosticos
'        Programado por: Castro W
'        Fecha: Agosto 2004
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim mb_Intrahospitalario As Boolean
Dim ml_IdDiagnostico As Long
Dim ms_Descripcion As String
Dim ms_CodigoCIE9 As String
Dim ms_CodigoCIE10 As String
Dim ms_CodigoExportacion As String
Dim ml_idTipoSexo As Long
Dim mb_Morbilidad As Boolean
Dim ml_IdCategoria As Long
Dim mb_Restriccion As Boolean
Dim ml_EdadMaxDias As Long
Dim ml_EdadMinDias As Long
Dim ms_CodigoCIE2004 As String
Dim ml_IdCapitulo As Long
Dim ml_idGrupo As Long
Dim mb_Gestacion As Boolean
Dim ms_DescripcionMINSA As String
Dim ms_codigoCIEsinPto As String
Dim md_FechaInicioVigencia As Date
Dim mb_EsActivo As Boolean

Property Let codigoCIEsinPto(sValue As String)
   ms_codigoCIEsinPto = sValue
End Property
Property Get codigoCIEsinPto() As String
   codigoCIEsinPto = ms_codigoCIEsinPto
End Property

Property Let DescripcionMINSA(sValue As String)
   ms_DescripcionMINSA = sValue
End Property
Property Get DescripcionMINSA() As String
   DescripcionMINSA = ms_DescripcionMINSA
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let Intrahospitalario(bValue As Boolean)
   mb_Intrahospitalario = bValue
End Property
Property Get Intrahospitalario() As Boolean
   Intrahospitalario = mb_Intrahospitalario
End Property
Property Let IdDiagnostico(lValue As Long)
   ml_IdDiagnostico = lValue
End Property
Property Get IdDiagnostico() As Long
   IdDiagnostico = ml_IdDiagnostico
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
Property Let CodigoCIE9(sValue As String)
   ms_CodigoCIE9 = sValue
End Property
Property Get CodigoCIE9() As String
   CodigoCIE9 = ms_CodigoCIE9
End Property
Property Let CodigoCIE10(sValue As String)
   ms_CodigoCIE10 = sValue
End Property
Property Get CodigoCIE10() As String
   CodigoCIE10 = ms_CodigoCIE10
End Property
Property Let CodigoExportacion(sValue As String)
   ms_CodigoExportacion = sValue
End Property
Property Get CodigoExportacion() As String
   CodigoExportacion = ms_CodigoExportacion
End Property
Property Let idTipoSexo(lValue As Long)
   ml_idTipoSexo = lValue
End Property
Property Get idTipoSexo() As Long
   idTipoSexo = ml_idTipoSexo
End Property
Property Let Morbilidad(bValue As Boolean)
   mb_Morbilidad = bValue
End Property
Property Get Morbilidad() As Boolean
   Morbilidad = mb_Morbilidad
End Property
Property Let IdCategoria(lValue As Long)
   ml_IdCategoria = lValue
End Property
Property Get IdCategoria() As Long
   IdCategoria = ml_IdCategoria
End Property
Property Let Restriccion(bValue As Boolean)
   mb_Restriccion = bValue
End Property
Property Get Restriccion() As Boolean
   Restriccion = mb_Restriccion
End Property
Property Let EdadMaxDias(lValue As Long)
   ml_EdadMaxDias = lValue
End Property
Property Get EdadMaxDias() As Long
   EdadMaxDias = ml_EdadMaxDias
End Property
Property Let EdadMinDias(lValue As Long)
   ml_EdadMinDias = lValue
End Property
Property Get EdadMinDias() As Long
   EdadMinDias = ml_EdadMinDias
End Property
Property Let CodigoCIE2004(sValue As String)
   ms_CodigoCIE2004 = sValue
End Property
Property Get CodigoCIE2004() As String
   CodigoCIE2004 = ms_CodigoCIE2004
End Property
Property Let IdCapitulo(lValue As Long)
   ml_IdCapitulo = lValue
End Property
Property Get IdCapitulo() As Long
   IdCapitulo = ml_IdCapitulo
End Property
Property Let idGrupo(lValue As Long)
   ml_idGrupo = lValue
End Property
Property Get idGrupo() As Long
   idGrupo = ml_idGrupo
End Property
Property Let Gestacion(bValue As Boolean)
   mb_Gestacion = bValue
End Property
Property Get Gestacion() As Boolean
   Gestacion = mb_Gestacion
End Property

Property Let EsActivo(bValue As Boolean)
   mb_EsActivo = bValue
End Property
Property Get EsActivo() As Boolean
   EsActivo = mb_EsActivo
End Property

Property Let FechaInicioVigencia(dValue As Date)
   md_FechaInicioVigencia = dValue
End Property
Property Get FechaInicioVigencia() As Date
   FechaInicioVigencia = md_FechaInicioVigencia
End Property
