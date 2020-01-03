VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RayosX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Rayos X
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS As Long

Dim mo_IdMovimiento As New ImagRayosX
Dim ml_idMovimiento As Long
Dim ml_IdTipoVentaSeleccionada As Long
Dim ml_SeEligioGridBoleta As Boolean

Property Let SeEligioGridBoleta(lValue As Boolean)
    ml_SeEligioGridBoleta = lValue
End Property
Property Get SeEligioGridBoleta() As Boolean
    SeEligioGridBoleta = ml_SeEligioGridBoleta
End Property
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Sub MostrarFormulario()
    mo_IdMovimiento.Opcion = mi_Opcion
    mo_IdMovimiento.idUsuario = ml_idUsuario
    mo_IdMovimiento.idMovimiento = ml_idMovimiento
    mo_IdMovimiento.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    mo_IdMovimiento.lcNombrePc = mo_lcNombrePc
    mo_IdMovimiento.SeEligioGridBoleta = ml_SeEligioGridBoleta
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
Property Let idMovimiento(lValue As Long)
   ml_idMovimiento = lValue
End Property


Sub ImpresionDelResultado(ml_idMovimiento As Long, lcResponsableImg As String, lcFechaRealizaCpt As String, _
                          lnIdProductoCPT As Long)
    Dim mrs_Tmp As New Recordset
    Dim oRsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oRsResultadoPorItem As New Recordset
    Dim oConexion As New Connection
    Dim mo_Diagnostico As New DODiagnostico
    Dim mo_reglasComunes As New ReglasComunes
    Dim mo_reglasCaja As New ReglasCaja
    Dim mo_ReglasFarmacia As New ReglasFarmacia
    Dim mo_ReglasImagenes As New ReglasImagenes
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim lcDx As String, lnIdComprobantePago As Long, lcBoleta As String, ml_IdDiagnostico As Long
    Dim lcPaciente As String, lcEdad As String, lcFfinanciamiento As String, lnIdCuenta As Long
    Dim lcResultadoFinal As String

    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    With mrs_Tmp
        .Fields.Append "cpt", adVarChar, 250, adFldIsNullable
        .Fields.Append "cptCantidad", adInteger
        .Fields.Append "Insumo", adVarChar, 250, adFldIsNullable
        .Fields.Append "InsumoCantidad", adInteger
        .LockType = adLockOptimistic
        .Open
    End With
    Set oRsTmp = mo_ReglasImagenes.ImagMovimientoCPTSeleccionarPorIdMovimiento(ml_idMovimiento)
    If lnIdProductoCPT > 0 Then
       oRsTmp.Filter = "idProductoCpt=" & lnIdProductoCPT
    End If
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          Set oRsTmp1 = mo_ReglasImagenes.ImagMovimientoDetalleSeleccionarPorIdMovimiento(ml_idMovimiento)
          oRsTmp1.Filter = "idProductoCpt=" & oRsTmp!idProductoCpt
          If oRsTmp1.RecordCount = 0 Then
                mrs_Tmp.AddNew
                mrs_Tmp!cpt = Left(oRsTmp!codigo & " " & oRsTmp!nombre, 250)
                mrs_Tmp!cptCantidad = oRsTmp!Cantidad
                mrs_Tmp.Update
          Else
                oRsTmp1.MoveFirst
                Do While Not oRsTmp1.EOF
                    mrs_Tmp.AddNew
                    mrs_Tmp!cpt = Left(oRsTmp!codigo & " " & oRsTmp!nombre, 250)
                    mrs_Tmp!cptCantidad = oRsTmp!Cantidad
                    mrs_Tmp!insumo = "" 'Left(oRsTmp1!codigo & " " & oRsTmp1!nombre, 250)
                    mrs_Tmp!insumoCantidad = 0 'oRsTmp1!Cantidad
                    mrs_Tmp.Update
                    oRsTmp1.MoveNext
                Loop
          End If
          oRsTmp1.Close
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = mo_ReglasImagenes.ImagMovimientoImagenesSeleccionarPorIdMovimiento(ml_idMovimiento)
    If oRsTmp.RecordCount > 0 Then
            
            '
            lcDx = ""
            ml_IdDiagnostico = IIf(IsNull(oRsTmp!idDiagnostico), 0, oRsTmp!idDiagnostico)
            If ml_IdDiagnostico > 0 Then
                Set mo_Diagnostico = mo_reglasComunes.DiagnosticosSeleccionarPorId(ml_IdDiagnostico)
                lcDx = mo_Diagnostico.CodigoCIE2004 & " " & mo_Diagnostico.descripcion
            End If
            '
            lcPaciente = ""
            lcBoleta = ""
            lcFfinanciamiento = ""
            lnIdComprobantePago = IIf(IsNull(oRsTmp.Fields!idComprobantePago), 0, oRsTmp.Fields!idComprobantePago)
            If lnIdComprobantePago > 0 Then
               Set oRsTmp1 = mo_reglasCaja.CajaComprobantesSeleccionarPorId(lnIdComprobantePago, oConexion)
               If oRsTmp1.RecordCount > 0 Then
                  lcBoleta = oRsTmp1!NroSerie & "-" & oRsTmp1!NroDocumento
                  lcFfinanciamiento = "Particular"
                  lcPaciente = oRsTmp1!razonSocial
               End If
               oRsTmp1.Close
            End If
            '
            lcEdad = ""
            lnIdCuenta = IIf(IsNull(oRsTmp.Fields!idCuentaAtencion), 0, oRsTmp.Fields!idCuentaAtencion)
            If lnIdCuenta > 0 Then
                Set oRsTmp1 = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(lnIdCuenta, oConexion)
                If oRsTmp1.RecordCount > 0 Then
                   wxParametro351 = lcBuscaParametro.SeleccionaFilaParametro(351)
                   lcPaciente = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oRsTmp1!NroHistoriaClinica)), False) & _
                              " " & Trim(oRsTmp1!ApellidoPaterno) & " " & _
                              Trim(oRsTmp1!ApellidoMaterno) & " " & oRsTmp1!PrimerNombre
                   lcFfinanciamiento = oRsTmp1!dFuenteFinanciamiento
                   lcEdad = Trim(Str(oRsTmp1!Edad)) & " " & oRsTmp1!tEdad
'                   iif(orstmp1!idTipoEdad= sghTipoEdades.sghAño," Años", _
'                   iif(orstmp1!idTipoEdad=sghTipoEdades.sghAño," Meses", _
'                   iif(orstmp1!idTipoEdad=sghTipoEdades.sghDias ," Días", "Horas")))
                End If
                oRsTmp1.Close
            Else
            End If
            '
            Set repResultadoImag.DataSource = mrs_Tmp
            repResultadoImag.RightMargin = 10
            repResultadoImag.TopMargin = 10
            repResultadoImag.LeftMargin = 10
            repResultadoImag.BottomMargin = 10
            repResultadoImag.Sections("cabecera").Controls("lblTitulo").Caption = IIf(oRsTmp!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaRayosX, "(RayosX", _
                                                                                 IIf(oRsTmp!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaTomografia, "(Tomografía", _
                                                                                 IIf(oRsTmp!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaEcogGeneral, "(Ecografía Gral", "(Ecografía Obst"))) & _
                                                                                 ") MOVIMIENTO N° " & Trim(Str(ml_idMovimiento))
            repResultadoImag.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
            repResultadoImag.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
            repResultadoImag.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
            repResultadoImag.Sections("cabecera").Controls("lblFecha").Caption = "Fecha: " & lcBuscaParametro.RetornaFechaServidorSQL
            repResultadoImag.Sections("cabecera").Controls("lblHora").Caption = "Hora: " & lcBuscaParametro.RetornaHoraServidorSQL
            Set repResultadoImag.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
            repResultadoImag.Sections("cabecera").Controls("lblFmovimiento").Caption = lcFechaRealizaCpt
            repResultadoImag.Sections("cabecera").Controls("lblPaciente").Caption = lcPaciente
            repResultadoImag.Sections("cabecera").Controls("lblEdad").Caption = lcEdad
            repResultadoImag.Sections("cabecera").Controls("lblDx").Caption = lcDx
            repResultadoImag.Sections("cabecera").Controls("lblFfinanciamiento").Caption = lcFfinanciamiento
            repResultadoImag.Sections("cabecera").Controls("lblCuenta").Caption = lnIdCuenta
            repResultadoImag.Sections("cabecera").Controls("lblBoleta").Caption = lcBoleta
            repResultadoImag.Sections("cabecera").Controls("lblResponsable").Caption = lcResponsableImg
            If lnIdProductoCPT > 0 Then
                lcResultadoFinal = ""
                Set oRsResultadoPorItem = mo_ReglasImagenes.ImagMovimientoResultadosSeleccionarPorId(ml_idMovimiento)
                oRsResultadoPorItem.Filter = "idProductoCpt=" & lnIdProductoCPT
                If oRsResultadoPorItem.RecordCount > 0 Then
                   lcResultadoFinal = oRsResultadoPorItem!Resultado
                End If
                oRsResultadoPorItem.Close
                repResultadoImag.Sections("Pie").Controls("lblResultado").Caption = lcResultadoFinal
            Else
                repResultadoImag.Sections("Pie").Controls("lblResultado").Caption = IIf(IsNull(oRsTmp.Fields!ResultadoFinal), "", oRsTmp.Fields!ResultadoFinal)
            End If
            If Val(lcBuscaParametro.SeleccionaFilaParametro(208)) = 7637 Then   'huaral
                repResultadoImag.Sections("Pie").Controls("lblResultado").Font.Size = 9
            End If
            repResultadoImag.Orientation = rptOrientPortrait
            repResultadoImag.Show 1
    End If
    Set mrs_Tmp = Nothing
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing
    Set oConexion = Nothing
    Set mo_Diagnostico = Nothing
    Set mo_reglasComunes = Nothing
    Set mo_reglasCaja = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_ReglasImagenes = Nothing
    Set lcBuscaParametro = Nothing
    Set oRsResultadoPorItem = Nothing
End Sub
