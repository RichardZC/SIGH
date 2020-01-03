VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ingresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Ingresos de Insumos
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_IdMovimiento As New LabIngresos
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim ml_IdMovimiento As Long
Dim ml_IdPuntoCarga As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Sub MostrarFormulario()
  mo_IdMovimiento.Opcion = mi_Opcion
  mo_IdMovimiento.idUsuario = ml_idUsuario
  mo_IdMovimiento.IdMovimiento = ml_IdMovimiento
  mo_IdMovimiento.IdPuntoCarga = ml_IdPuntoCarga
  mo_IdMovimiento.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
  mo_IdMovimiento.lcNombrePc = mo_lcNombrePc
  mo_IdMovimiento.Show 1
  Set mo_IdMovimiento = Nothing
End Sub

Property Get Opcion() As sghOpciones
  Opcion = mi_Opcion
End Property

Property Let idUsuario(lValue As Long)
  ml_idUsuario = lValue
End Property

Property Let Opcion(lValue As sghOpciones)
  mi_Opcion = lValue
End Property

Property Let IdMovimiento(lValue As Long)
  ml_IdMovimiento = lValue
End Property

Property Let IdPuntoCarga(lValue As Long)
  ml_IdPuntoCarga = lValue
End Property



Sub MuestraResultadoDelExamen(ml_IdPruebaSeleccionada As String, ml_NombrePruebaSeleccionada As String, _
                              ml_nombrePaciente As String, ml_idOrden As Long, _
                              ml_idPaciente As Long, ml_nombreMedico As String, _
                              ml_areaTrabajo As Long, ml_idOrdenLab As Long, _
                              lnIdTipoSexo As Long, lbNoMuestraBotonGrabar As Boolean, _
                              Optional grdListaDetalleRS As Recordset, Optional ldFechaNacim As Date)
    Dim mo_Bioquimica As New SIGHLaboratorio.bioquimica
    Dim mo_Citopatologia As New SIGHLaboratorio.citopatologia
    Dim mo_Hematologia  As New SIGHLaboratorio.hematologia
    Dim mo_Inmunoserologia  As New SIGHLaboratorio.inmunoserologia
    Dim mo_Microbiologia As New SIGHLaboratorio.microbiologia
    Dim mo_Parasitologia As New SIGHLaboratorio.parasitologia
    Dim mo_Patologia As New SIGHLaboratorio.patologia
    Dim mo_Urianalisis As New SIGHLaboratorio.urianalisis
    Dim mo_BancoSangre As New SIGHLaboratorio.bancoSangre
    Dim mo_ReglasLaboratorio As New ReglasLaboratorio
    Dim ml_CodigoPruebaSeleccionada As String, ldFechaNacimiento As Date
    
'  ml_CodigoPruebaSeleccionada = mo_ReglasLaboratorio.LabAveriguaCodigoPrueba(ml_IdPruebaSeleccionada)
'  ldFechaNacimiento = Date
  
    ml_CodigoPruebaSeleccionada = mo_ReglasLaboratorio.LabAveriguaCodigoPrueba(ml_IdPruebaSeleccionada)
    If Year(ldFechaNacim) = 1899 Then 'Yamill estuvo aqui
      ldFechaNacimiento = Date
    Else
      ldFechaNacimiento = ldFechaNacim
    End If
    If Left(ml_CodigoPruebaSeleccionada, 3) = "BQM" Then
  'Bioquímica
    mo_Bioquimica.idUsuario = sighentidades.Usuario
    mo_Bioquimica.idOrden = ml_idOrden
    mo_Bioquimica.idPrueba = ml_IdPruebaSeleccionada
    mo_Bioquimica.nombrePrueba = ml_NombrePruebaSeleccionada
    mo_Bioquimica.idPaciente = ml_idPaciente
    mo_Bioquimica.nombreMedico = ml_nombreMedico
    mo_Bioquimica.nombrePaciente = ml_nombrePaciente
    mo_Bioquimica.AreaTrabajo = ml_areaTrabajo
    mo_Bioquimica.CodigoPruebaSeleccionada = ml_CodigoPruebaSeleccionada
    If Not (grdListaDetalleRS Is Nothing) Then
       mo_Bioquimica.DetalleOrden = grdListaDetalleRS  'grdListaOrdenesDetalle.DataSource
    End If
    mo_Bioquimica.idOrdenLab = ml_idOrdenLab
    mo_Bioquimica.FechaNacimiento = ldFechaNacimiento
    mo_Bioquimica.idTipoSexo = lnIdTipoSexo
    mo_Bioquimica.NoMuestraBotonGrabar = lbNoMuestraBotonGrabar
    mo_Bioquimica.MostrarFormulario
  ElseIf Left(ml_CodigoPruebaSeleccionada, 3) = "HEM" Then
  'Hematología
    mo_Hematologia.idUsuario = sighentidades.Usuario
    mo_Hematologia.idOrden = ml_idOrden
    mo_Hematologia.idPrueba = ml_IdPruebaSeleccionada
    mo_Hematologia.nombrePrueba = ml_NombrePruebaSeleccionada
    mo_Hematologia.idPaciente = ml_idPaciente
    mo_Hematologia.nombreMedico = ml_nombreMedico
    mo_Hematologia.nombrePaciente = ml_nombrePaciente
    mo_Hematologia.AreaTrabajo = ml_areaTrabajo
    mo_Hematologia.CodigoPruebaSeleccionada = ml_CodigoPruebaSeleccionada
    If Not (grdListaDetalleRS Is Nothing) Then
       mo_Hematologia.DetalleOrden = grdListaDetalleRS 'grdListaOrdenesDetalle.DataSource
    End If
    mo_Hematologia.idOrdenLab = ml_idOrdenLab
    mo_Hematologia.FechaNacimiento = ldFechaNacimiento
    mo_Hematologia.idTipoSexo = lnIdTipoSexo
    mo_Hematologia.NoMuestraBotonGrabar = lbNoMuestraBotonGrabar
    mo_Hematologia.MostrarFormulario
  ElseIf Left(ml_CodigoPruebaSeleccionada, 3) = "INM" Then
  'Inmunoserología
    mo_Inmunoserologia.idUsuario = sighentidades.Usuario
    mo_Inmunoserologia.idOrden = ml_idOrden
    mo_Inmunoserologia.idPrueba = ml_IdPruebaSeleccionada
    mo_Inmunoserologia.nombrePrueba = ml_NombrePruebaSeleccionada
    mo_Inmunoserologia.idPaciente = ml_idPaciente
    mo_Inmunoserologia.nombreMedico = ml_nombreMedico
    mo_Inmunoserologia.nombrePaciente = ml_nombrePaciente
    mo_Inmunoserologia.AreaTrabajo = ml_areaTrabajo
    mo_Inmunoserologia.CodigoPruebaSeleccionada = ml_CodigoPruebaSeleccionada
    If Not (grdListaDetalleRS Is Nothing) Then
        mo_Inmunoserologia.DetalleOrden = grdListaDetalleRS 'grdListaOrdenesDetalle.DataSource
    End If
    mo_Inmunoserologia.idOrdenLab = ml_idOrdenLab
    mo_Inmunoserologia.FechaNacimiento = ldFechaNacimiento
    mo_Inmunoserologia.idTipoSexo = lnIdTipoSexo
    mo_Inmunoserologia.NoMuestraBotonGrabar = lbNoMuestraBotonGrabar
    mo_Inmunoserologia.MostrarFormulario
  ElseIf Left(ml_CodigoPruebaSeleccionada, 3) = "MIC" Then
  'Microbiología
    mo_Microbiologia.idUsuario = sighentidades.Usuario
    mo_Microbiologia.idOrden = ml_idOrden
    mo_Microbiologia.idPrueba = ml_IdPruebaSeleccionada
    mo_Microbiologia.nombrePrueba = ml_NombrePruebaSeleccionada
    mo_Microbiologia.idPaciente = ml_idPaciente
    mo_Microbiologia.nombreMedico = ml_nombreMedico
    mo_Microbiologia.nombrePaciente = ml_nombrePaciente
    mo_Microbiologia.AreaTrabajo = ml_areaTrabajo
    mo_Microbiologia.CodigoPruebaSeleccionada = ml_CodigoPruebaSeleccionada
    If Not (grdListaDetalleRS Is Nothing) Then
       mo_Microbiologia.DetalleOrden = grdListaDetalleRS 'grdListaOrdenesDetalle.DataSource
    End If
    mo_Microbiologia.idOrdenLab = ml_idOrdenLab
    mo_Microbiologia.FechaNacimiento = ldFechaNacimiento
    mo_Microbiologia.idTipoSexo = lnIdTipoSexo
    mo_Microbiologia.NoMuestraBotonGrabar = lbNoMuestraBotonGrabar
    mo_Microbiologia.MostrarFormulario
  ElseIf Left(ml_CodigoPruebaSeleccionada, 3) = "COA" Then
  'Parasitología
    mo_Parasitologia.idUsuario = sighentidades.Usuario
    mo_Parasitologia.idOrden = ml_idOrden
    mo_Parasitologia.idPrueba = ml_IdPruebaSeleccionada
    mo_Parasitologia.nombrePrueba = ml_NombrePruebaSeleccionada
    mo_Parasitologia.idPaciente = ml_idPaciente
    mo_Parasitologia.nombreMedico = ml_nombreMedico
    mo_Parasitologia.nombrePaciente = ml_nombrePaciente
    mo_Parasitologia.AreaTrabajo = ml_areaTrabajo
    mo_Parasitologia.CodigoPruebaSeleccionada = ml_CodigoPruebaSeleccionada
    If Not (grdListaDetalleRS Is Nothing) Then
       mo_Parasitologia.DetalleOrden = grdListaDetalleRS 'grdListaOrdenesDetalle.DataSource
    End If
    mo_Parasitologia.idOrdenLab = ml_idOrdenLab
    mo_Parasitologia.FechaNacimiento = ldFechaNacimiento
    mo_Parasitologia.idTipoSexo = lnIdTipoSexo
    mo_Parasitologia.NoMuestraBotonGrabar = lbNoMuestraBotonGrabar
    mo_Parasitologia.MostrarFormulario
  ElseIf Left(ml_CodigoPruebaSeleccionada, 3) = "ANA" Then
  'Urianálisis
    mo_Urianalisis.idUsuario = sighentidades.Usuario
    mo_Urianalisis.idOrden = ml_idOrden
    mo_Urianalisis.idPrueba = ml_IdPruebaSeleccionada
    mo_Urianalisis.nombrePrueba = ml_NombrePruebaSeleccionada
    mo_Urianalisis.idPaciente = ml_idPaciente
    mo_Urianalisis.nombreMedico = ml_nombreMedico
    mo_Urianalisis.nombrePaciente = ml_nombrePaciente
    mo_Urianalisis.AreaTrabajo = ml_areaTrabajo
    mo_Urianalisis.CodigoPruebaSeleccionada = ml_CodigoPruebaSeleccionada
    If Not (grdListaDetalleRS Is Nothing) Then
       mo_Urianalisis.DetalleOrden = grdListaDetalleRS 'grdListaOrdenesDetalle.DataSource
    End If
    mo_Urianalisis.idOrdenLab = ml_idOrdenLab
    mo_Urianalisis.FechaNacimiento = ldFechaNacimiento
    mo_Urianalisis.idTipoSexo = lnIdTipoSexo
    mo_Urianalisis.NoMuestraBotonGrabar = lbNoMuestraBotonGrabar
    mo_Urianalisis.MostrarFormulario
  ElseIf Left(ml_CodigoPruebaSeleccionada, 3) = "CPA" Then
  'Citopatología
    mo_Citopatologia.idUsuario = sighentidades.Usuario
    mo_Citopatologia.idOrden = ml_idOrden
    mo_Citopatologia.idPrueba = ml_IdPruebaSeleccionada
    mo_Citopatologia.nombrePrueba = ml_NombrePruebaSeleccionada
    mo_Citopatologia.idPaciente = ml_idPaciente
    mo_Citopatologia.nombreMedico = ml_nombreMedico
    mo_Citopatologia.nombrePaciente = ml_nombrePaciente
    mo_Citopatologia.AreaTrabajo = ml_areaTrabajo
    mo_Citopatologia.CodigoPruebaSeleccionada = ml_CodigoPruebaSeleccionada
    If Not (grdListaDetalleRS Is Nothing) Then
       mo_Citopatologia.DetalleOrden = grdListaDetalleRS 'grdListaOrdenesDetalle.DataSource
    End If
    mo_Citopatologia.idOrdenLab = ml_idOrdenLab
    mo_Citopatologia.FechaNacimiento = ldFechaNacimiento
    mo_Citopatologia.idTipoSexo = lnIdTipoSexo
    mo_Citopatologia.NoMuestraBotonGrabar = lbNoMuestraBotonGrabar
    mo_Citopatologia.MostrarFormulario
  ElseIf Left(ml_CodigoPruebaSeleccionada, 3) = "PAQ" Then
  'Anatomía Patológica
    mo_Patologia.idUsuario = sighentidades.Usuario
    mo_Patologia.idOrden = ml_idOrden
    mo_Patologia.idPrueba = ml_IdPruebaSeleccionada
    mo_Patologia.nombrePrueba = ml_NombrePruebaSeleccionada
    mo_Patologia.idPaciente = ml_idPaciente
    mo_Patologia.nombreMedico = ml_nombreMedico
    mo_Patologia.nombrePaciente = ml_nombrePaciente
    mo_Patologia.AreaTrabajo = ml_areaTrabajo
    mo_Patologia.CodigoPruebaSeleccionada = ml_CodigoPruebaSeleccionada
    If Not (grdListaDetalleRS Is Nothing) Then
       mo_Patologia.DetalleOrden = grdListaDetalleRS 'grdListaOrdenesDetalle.DataSource
    End If
    mo_Patologia.idOrdenLab = ml_idOrdenLab
    mo_Patologia.FechaNacimiento = ldFechaNacimiento
    mo_Patologia.idTipoSexo = lnIdTipoSexo
    mo_Patologia.NoMuestraBotonGrabar = lbNoMuestraBotonGrabar
    mo_Patologia.MostrarFormulario
  ElseIf Left(ml_CodigoPruebaSeleccionada, 3) = "BSA" Then
  'Banco de Sangre
    mo_BancoSangre.idUsuario = sighentidades.Usuario
    mo_BancoSangre.idOrden = ml_idOrden
    mo_BancoSangre.idPrueba = ml_IdPruebaSeleccionada
    mo_BancoSangre.nombrePrueba = ml_NombrePruebaSeleccionada
    mo_BancoSangre.idPaciente = ml_idPaciente
    mo_BancoSangre.nombreMedico = ml_nombreMedico
    mo_BancoSangre.nombrePaciente = ml_nombrePaciente
    mo_BancoSangre.AreaTrabajo = ml_areaTrabajo
    mo_BancoSangre.CodigoPruebaSeleccionada = ml_CodigoPruebaSeleccionada
    mo_BancoSangre.DetalleOrden = grdListaDetalleRS 'grdListaOrdenesDetalle.DataSource
    mo_BancoSangre.idOrdenLab = ml_idOrdenLab
    mo_BancoSangre.FechaNacimiento = ldFechaNacimiento
    mo_BancoSangre.idTipoSexo = lnIdTipoSexo
    mo_BancoSangre.NoMuestraBotonGrabar = lbNoMuestraBotonGrabar
    mo_BancoSangre.MostrarFormulario
  Else
    MsgBox "La prueba seleccionada no tiene formato para ingreso de resultados, escoja otra prueba.", vbInformation, "SIGH "
  End If

  Set mo_Bioquimica = Nothing
  Set mo_Citopatologia = Nothing
  Set mo_Hematologia = Nothing
  Set mo_Inmunoserologia = Nothing
  Set mo_Microbiologia = Nothing
  Set mo_Parasitologia = Nothing
  Set mo_Patologia = Nothing
  Set mo_Urianalisis = Nothing
  Set mo_BancoSangre = Nothing
  Set mo_ReglasLaboratorio = Nothing

End Sub


