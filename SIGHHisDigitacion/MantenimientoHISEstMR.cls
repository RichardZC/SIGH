VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MantenimientoHISEstMR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para mantenimiento de HIS microRed
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mi_Opcion As sghOpciones
Dim ml_IdUsuario As Long
Dim ml_IdEstablecimiento As Long
Dim ms_NombreEstablecimiento As String
Dim ms_CodigoEstablecimiento As String
Dim mo_RcsListaEstablecimientos As New Recordset

Dim mo_FormularioHISEstabMR As New frmMantenimientoEstMR

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Get Opcion() As sghOpciones
    Opcion = mi_Opcion
End Property

Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property

Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property

Property Let IdEstablecimiento(lValue As Long)
   ml_IdEstablecimiento = lValue
End Property

Property Get IdEstablecimiento() As Long
   IdEstablecimiento = ml_IdEstablecimiento
End Property

Property Let NombreEstablecimiento(lValue As String)
   ms_NombreEstablecimiento = lValue
End Property

Property Get NombreEstablecimiento() As String
   NombreEstablecimiento = ms_NombreEstablecimiento
End Property

Property Let CodigoEstablecimiento(lValue As String)
   ms_CodigoEstablecimiento = lValue
End Property

Property Get CodigoEstablecimiento() As String
   CodigoEstablecimiento = ms_CodigoEstablecimiento
End Property

Sub MostrarFormulario()
    mo_FormularioHISEstabMR.Opcion = mi_Opcion
    mo_FormularioHISEstabMR.IdUsuario = ml_IdUsuario
    mo_FormularioHISEstabMR.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    mo_FormularioHISEstabMR.lcNombrePc = mo_lcNombrePc
    mo_FormularioHISEstabMR.IdEstablecimiento = ml_IdEstablecimiento
    mo_FormularioHISEstabMR.NombreEstablecimiento = ms_NombreEstablecimiento
    mo_FormularioHISEstabMR.CodigoEstablecimiento = ms_CodigoEstablecimiento
    mo_FormularioHISEstabMR.Show 1
    Set mo_FormularioHISEstabMR = Nothing
End Sub

