VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoPerinatalAtencionCred1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla PerinatalAtencionCred1
'        Programado por: Garay M
'        Fecha: Setiembre 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idPerinatalAtencion As Long
Dim ml_idModulo As Long
Dim mb_EstimulacionTemprana As Boolean
Dim mb_AlimentacionComplementaria As Boolean
Dim mb_LactanciaMaterna As Boolean
Dim mb_PersonalSalud As Boolean
Dim mb_DemandaIndividual As Boolean
Dim mb_MujerEdadReproductiva As Boolean
Dim mb_MujerGestante As Boolean
Dim ml_IdAtencion As Long
Dim mb_LactanciaMaternaComp As Boolean
Property Let LactanciaMaternaComp(bValue As Boolean)
   mb_LactanciaMaternaComp = bValue
End Property
Property Get LactanciaMaternaComp() As Boolean
   LactanciaMaternaComp = mb_LactanciaMaternaComp
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idPerinatalAtencion(iValue As Long)
   ml_idPerinatalAtencion = iValue
End Property
Property Get idPerinatalAtencion() As Long
   idPerinatalAtencion = ml_idPerinatalAtencion
End Property
Property Let idModulo(iValue As Long)
   ml_idModulo = iValue
End Property
Property Get idModulo() As Long
   idModulo = ml_idModulo
End Property
Property Let EstimulacionTemprana(bValue As Boolean)
   mb_EstimulacionTemprana = bValue
End Property
Property Get EstimulacionTemprana() As Boolean
   EstimulacionTemprana = mb_EstimulacionTemprana
End Property
Property Let AlimentacionComplementaria(bValue As Boolean)
   mb_AlimentacionComplementaria = bValue
End Property
Property Get AlimentacionComplementaria() As Boolean
   AlimentacionComplementaria = mb_AlimentacionComplementaria
End Property
Property Let LactanciaMaterna(bValue As Boolean)
   mb_LactanciaMaterna = bValue
End Property
Property Get LactanciaMaterna() As Boolean
   LactanciaMaterna = mb_LactanciaMaterna
End Property
Property Let PersonalSalud(bValue As Boolean)
   mb_PersonalSalud = bValue
End Property
Property Get PersonalSalud() As Boolean
   PersonalSalud = mb_PersonalSalud
End Property
Property Let DemandaIndividual(bValue As Boolean)
   mb_DemandaIndividual = bValue
End Property
Property Get DemandaIndividual() As Boolean
   DemandaIndividual = mb_DemandaIndividual
End Property
Property Let MujerEdadReproductiva(bValue As Boolean)
   mb_MujerEdadReproductiva = bValue
End Property
Property Get MujerEdadReproductiva() As Boolean
   MujerEdadReproductiva = mb_MujerEdadReproductiva
End Property
Property Let MujerGestante(bValue As Boolean)
   mb_MujerGestante = bValue
End Property
Property Get MujerGestante() As Boolean
   MujerGestante = mb_MujerGestante
End Property
Property Let IdAtencion(iValue As Long)
   ml_IdAtencion = iValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property
