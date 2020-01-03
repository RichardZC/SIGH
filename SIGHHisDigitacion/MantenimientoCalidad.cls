VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MantenimientoCalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para mantenimiento de calidad
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim mo_DetallePadronInicial As New frmPadronNominal
Dim mi_Opcion As sghOpciones
Dim ml_IdUsuario As Long
Dim ml_IdEstablecimiento As Long
Dim ml_IdPadNominal As Long

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

Property Get IdPadNominal() As Long
    IdPadNominal = ml_IdPadNominal
End Property

Property Let IdPadNominal(lValue As Long)
   ml_IdPadNominal = lValue
End Property

Property Let BotonPresionado(oValue As sghBotonDetallePresionado)
   mi_BotonPresionado = oValue
End Property

Sub MostrarFormulario()
    mo_DetallePadronInicial.Opcion = mi_Opcion
    mo_DetallePadronInicial.IdUsuario = ml_IdUsuario
    mo_DetallePadronInicial.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    mo_DetallePadronInicial.lcNombrePc = mo_lcNombrePc
    mo_DetallePadronInicial.IdPadNominal = ml_IdPadNominal
    mo_DetallePadronInicial.Show 1
    mi_BotonPresionado = mo_DetallePadronInicial.BotonPresionado
End Sub

