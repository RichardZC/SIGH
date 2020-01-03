VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CamaDetalleProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Mantenimiento de CAMAS para Hospitalización y Emergencia
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_CamaDetalle As New CamaDetalle
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let Opcion(lValue As sghOpciones)
    mo_CamaDetalle.Opcion = lValue
End Property
Property Get Opcion() As sghOpciones
    Opcion = mo_CamaDetalle.Opcion
End Property
Property Let IdCama(lValue As Long)
    mo_CamaDetalle.IdCama = lValue
End Property
Property Get IdCama() As Long
    IdCama = mo_CamaDetalle.IdCama
End Property
Property Let IdTipoServicio(lValue As Long)
   mo_CamaDetalle.cmbIdTipoServicio.BoundText = lValue
End Property
Property Get IdTipoServicio() As Long
   IdTipoServicio = mo_CamaDetalle.cmbIdTipoServicio.BoundText
End Property
Property Let IdServicio(lValue As Long)
   mo_CamaDetalle.txtIdServicioPropietario.Tag = lValue
   mo_CamaDetalle.txtIdServicioUbicacionActual.Tag = lValue
End Property
Property Get IdServicio() As Long
   IdServicio = mo_CamaDetalle.txtIdServicioPropietario.Tag
End Property
Property Let CodigoServicio(lValue As String)
   mo_CamaDetalle.txtIdServicioPropietario = lValue
   mo_CamaDetalle.txtIdServicioUbicacionActual = lValue
End Property
Property Get CodigoServicio() As String
   CodigoServicio = mo_CamaDetalle.txtIdServicioPropietario
End Property
Property Let NombreServicio(lValue As String)
   mo_CamaDetalle.lblNombreServicioPropietario = lValue
   mo_CamaDetalle.lblNombreServicioUbicacionActual = lValue
End Property
Property Get NombreServicio() As String
   NombreServicio = mo_CamaDetalle.lblNombreServicioPropietario
End Property
Property Get ConfirmoOperacion() As Integer
    ConfirmoOperacion = mo_CamaDetalle.ConfirmoOperacion
End Property
Property Let IdUsuario(lValue As Long)
   mo_CamaDetalle.IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = mo_CamaDetalle.IdUsuario
End Property

Sub MostrarDialogo(lnIdTipoServicio As Long)
    mo_CamaDetalle.lnIdTablaLISTBARITEMS = 303
    
    mo_CamaDetalle.lcNombrePc = "."
    mo_CamaDetalle.IdTipoServicioActual = lnIdTipoServicio
    mo_CamaDetalle.Show 1
End Sub

