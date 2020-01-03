VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form rRecetasXservicio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recetas por Servicio"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rRecetasXservicio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   30
      TabIndex        =   2
      Top             =   2565
      Width           =   9180
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rRecetasXservicio.frx":0CCA
         DownPicture     =   "rRecetasXservicio.frx":118E
         Height          =   700
         Left            =   4740
         Picture         =   "rRecetasXservicio.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rRecetasXservicio.frx":1B66
         DownPicture     =   "rRecetasXservicio.frx":1FC6
         Height          =   700
         Left            =   3210
         Picture         =   "rRecetasXservicio.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame fraDatosHistoria 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   30
      TabIndex        =   1
      Top             =   90
      Width           =   9195
      Begin Threed.SSOption optPorServicio 
         Height          =   315
         Left            =   1305
         TabIndex        =   14
         Top             =   1800
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "Recetas por Servicio "
         Value           =   -1
      End
      Begin VB.CheckBox chkExcel 
         Caption         =   "En Excel"
         Height          =   315
         Left            =   1320
         Picture         =   "rRecetasXservicio.frx":28B0
         TabIndex        =   13
         Top             =   1485
         Width           =   1125
      End
      Begin VB.ComboBox cmbAlmacen 
         Height          =   330
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   7590
      End
      Begin VB.CheckBox chkDetalle 
         Caption         =   "Se muestra relación de Documentos"
         Height          =   225
         Left            =   1320
         TabIndex        =   8
         Top             =   1230
         Width           =   3375
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFhasta 
         Height          =   315
         Left            =   6690
         TabIndex        =   5
         Top             =   720
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHrInicio 
         Height          =   315
         Left            =   2730
         TabIndex        =   11
         Top             =   720
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHrFin 
         Height          =   315
         Left            =   8100
         TabIndex        =   12
         Top             =   720
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin Threed.SSOption optRecetasXplan 
         Height          =   315
         Left            =   1305
         TabIndex        =   15
         Top             =   2100
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   556
         _Version        =   262144
         Caption         =   "Recetas por Plan"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Movimiento"
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
         Height          =   210
         Left            =   6180
         TabIndex        =   6
         Top             =   750
         Width           =   435
      End
   End
End
Attribute VB_Name = "rRecetasXservicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte de Recetas por Servicio
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim sMensaje As String
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_TextoDelFiltro As String
Const ml_IdPuntoCarga As Integer = 5
Dim lnIdProducto As Long
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_cmbAlmacen As New sighentidades.ListaDespleglable
Dim ml_idUsuario As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        If Me.optPorServicio.Value = True Then
            Dim oRptClaseCry As New rCrystal
            oRptClaseCry.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
            oRptClaseCry.IdAlmacen = Val(mo_cmbAlmacen.BoundText)
            oRptClaseCry.FechaInicio = CDate(Format(txtFdesde.Text & " " & txtHrInicio & ":00", sighentidades.DevuelveFechaSoloFormato_DMY_HMS))
            oRptClaseCry.FechaFin = CDate(Format(txtFhasta.Text & " " & txtHrFin & ":00", sighentidades.DevuelveFechaSoloFormato_DMY_HMS))
            oRptClaseCry.TextoDelFiltro = ml_TextoDelFiltro
            oRptClaseCry.SeMuestraLotes = IIf(chkDetalle.Value = 1, True, False)
            oRptClaseCry.TipoReporte = Me.Name
            oRptClaseCry.Show vbModal
            Set oRptClaseCry = Nothing
        Else
            RecetasPorPlan CDate(Format(txtFdesde.Text & " " & txtHrInicio & ":00", sighentidades.DevuelveFechaSoloFormato_DMY_HMS)), CDate(Format(txtFhasta.Text & " " & txtHrFin & ":00", sighentidades.DevuelveFechaSoloFormato_DMY_HMS))
        End If
        Me.MousePointer = 1
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    If CDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) > CDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
       Exit Function
    End If
    sMensaje = ""
    ml_TextoDelFiltro = "FILTROS:   Almacén: (" & Trim(cmbAlmacen.Text) & ")      F.Movimiento: (" & txtFdesde.Text & " " & txtHrInicio.Text & "    al " & txtFhasta.Text & " " & txtHrFin.Text & ")"
    If mo_cmbAlmacen.BoundText = "" Then
        sMensaje = sMensaje + "Por favor elija el Almacén" + Chr(13)
        cmbAlmacen.SetFocus
    End If
    If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function


Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub


Sub InicializaFechaHora()

    txtFdesde.Text = sighentidades.PrimerFechaDDMMYYDelMesActual
    txtFhasta.Text = Date
    txtHrInicio.Text = lcBuscaParametro.SeleccionaFilaParametro(267)
    txtHrFin.Text = lcBuscaParametro.SeleccionaFilaParametro(268)
End Sub

Private Sub Form_Load()
    InicializaFechaHora
    '
    mo_cmbAlmacen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacen.ListField = "Descripcion"
    Set mo_cmbAlmacen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' and idtipoSuministro='01'")
    '
    Dim rsIdAlmacen As Recordset
    Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
    Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAlmacenFarmacia, ml_idUsuario)
    Set oBuscaDondeLabora = Nothing
    If rsIdAlmacen.RecordCount > 0 Then
       mo_cmbAlmacen.BoundText = rsIdAlmacen.Fields!idLaboraSubArea
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmacen, False
    End If

End Sub



Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub








Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub optPorServicio_Click(Value As Integer)
    If optPorServicio.Value = True Then
       chkExcel.Enabled = True
    End If
End Sub

Private Sub optRecetasXplan_Click(Value As Integer)
    If optRecetasXplan.Value = True Then
            chkExcel.Enabled = False
            chkExcel.Value = 1
    End If
End Sub

Private Sub txtFdesde_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFdesde

End Sub



Private Sub txtFdesde_LostFocus()
    If txtFdesde <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.esfecha(txtFdesde, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
        End If
    End If

End Sub

Private Sub txtFhasta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFhasta

End Sub

Private Sub txtFhasta_LostFocus()
    If txtFhasta <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.esfecha(txtFhasta, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
        End If
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub


Private Sub Form_Initialize()
    Set mo_cmbAlmacen.MiComboBox = cmbAlmacen
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_ReglasFarmacia = Nothing
    Set mo_Teclado = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_ReglasComunes = Nothing
    Set mo_Formulario = Nothing
    Set mo_cmbAlmacen = Nothing
End Sub

Private Sub txtHrFin_LostFocus()
If Not sighentidades.ValidaHora(txtHrFin.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            InicializaFechaHora
        End If
End Sub

Private Sub txtHrInicio_LostFocus()
If Not sighentidades.ValidaHora(txtHrInicio.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            InicializaFechaHora
        End If
End Sub


Sub RecetasPorPlan(mda_FechaInicio As Date, mda_FechaFin As Date)
        Dim rsRecetas As New Recordset
        Dim rsFuenteFinan As New Recordset
        Dim rsServicios As New Recordset
        Dim rsTmp As New Recordset
        Dim mo_ReglasComunes As New ReglasComunes
        Dim mo_ReglasFacturacion As New ReglasFacturacion
        Dim oConexion As New Connection
        Dim mo_ReglasReportes As New ReglasReportes
        Dim lcFuente As String, lcTServicio As String, lbNuevo As Boolean, lnTotal As Long, lnTotalG As Long
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        With rsRecetas
            .Fields.Append "Tipo", adVarChar, 50, adFldIsNullable
            .Fields.Append "Servicios", adVarChar, 50, adFldIsNullable
            .Fields.Append "cantidad", adInteger
            If chkDetalle.Value = 1 Then
               .Fields.Append "DocumentoNumero", adVarChar, 50, adFldIsNullable
            Else
               .Fields.Append "CantTotal", adInteger
            End If
            .LockType = adLockOptimistic
            .Open
        End With
        Set rsTmp = mo_ReglasFarmacia.farmMovimientoFiltrarXfechas(mda_FechaInicio, mda_FechaFin, oConexion)
        rsTmp.Filter = "MovTipo='S' and idAlmacenOrigen=" & mo_cmbAlmacen.BoundText
        If rsTmp.RecordCount > 0 Then
           Set rsFuenteFinan = mo_ReglasComunes.FuentesFinanciamientoSeleccionarTodos
           Set rsServicios = mo_ReglasFacturacion.ServiciosSeleccionarPorFiltro(" ", sghPorCodigo)
           rsTmp.MoveFirst
           Do While Not rsTmp.EOF
              lcFuente = ""
              rsFuenteFinan.MoveFirst
              rsFuenteFinan.Find "idFuenteFinanciamiento=" & IIf(rsTmp!IdTipoFinanciamiento = sghTipoFinanciamiento.sghPacienteNormal, _
                                                                1, rsTmp!idFuenteFinanciamiento)
              If Not rsFuenteFinan.EOF Then
                 lcFuente = rsFuenteFinan!Descripcion
              End If
              '
              lcTServicio = "PACIENTE EXTERNO"
              If Not IsNull(rsTmp!IdServicioPaciente) Then
                    rsServicios.MoveFirst
                    rsServicios.Find "idServicio=" & rsTmp!IdServicioPaciente
                    If Not rsServicios.EOF Then
                       Select Case rsServicios!IdTipoServicio
                       Case sghTipoServicio.sghConsultaExterna
                            lcTServicio = "Consultorio Externo"
                       Case sghTipoServicio.sghEmergenciaConsultorios
                            lcTServicio = "Emergencia"
                       Case sghTipoServicio.sghHospitalizacion
                            lcTServicio = "Hospitalizaciòn"
                       Case Else
                            lcTServicio = "PACIENTE EXTERNO"
                       End Select
                    End If
              End If
              '
              lbNuevo = True
              If rsRecetas.RecordCount > 0 And chkDetalle.Value = 0 Then
                 rsRecetas.MoveFirst
                 Do While Not rsRecetas.EOF
                    If rsRecetas!tipo = lcFuente And rsRecetas!Servicios = lcTServicio Then
                       lbNuevo = False
                       Exit Do
                    End If
                    rsRecetas.MoveNext
                 Loop
            
              End If
              If lbNuevo = True Then
                 rsRecetas.AddNew
                 rsRecetas.Fields!tipo = lcFuente
                 rsRecetas.Fields!Servicios = lcTServicio
                 rsRecetas.Fields!Cantidad = 1
                 If chkDetalle.Value = 1 Then
                    rsRecetas.Fields!DocumentoNumero = rsTmp!DocumentoNumero & " - " & Format(rsTmp!fechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY_HM)
                 End If
              Else
                 rsRecetas.Fields!Cantidad = rsRecetas.Fields!Cantidad + 1
              End If
              rsRecetas.Update
              
              rsTmp.MoveNext
           Loop
        End If
        rsTmp.Close
        Set rsTmp = mo_ReglasFarmacia.farmMovimientoFiltrarIntervSanitaria(mda_FechaInicio, mda_FechaFin, oConexion)
        rsTmp.Filter = "idAlmacenOrigen=" & mo_cmbAlmacen.BoundText
        If chkDetalle.Value = 1 Then
            If rsTmp.RecordCount > 0 Then
               rsTmp.MoveFirst
               Do While Not rsTmp.EOF
                    rsRecetas.AddNew
                    rsRecetas.Fields!tipo = "INTERVENCIONES SANITARIAS"
                    rsRecetas.Fields!Servicios = "INTERVENCIONES SANITARIAS"
                    rsRecetas.Fields!Cantidad = 1
                    rsRecetas.Fields!DocumentoNumero = rsTmp!DocumentoNumero
                    rsRecetas.Update
                    rsTmp.MoveNext
               Loop
            End If
        Else
            rsRecetas.AddNew
            rsRecetas.Fields!tipo = "INTERVENCIONES SANITARIAS"
            rsRecetas.Fields!Servicios = "INTERVENCIONES SANITARIAS"
            rsRecetas.Fields!Cantidad = rsTmp.RecordCount
            rsRecetas.Update
        End If
        rsTmp.Close
        
        If rsRecetas.RecordCount = 0 Then
           MsgBox "No existe informaciòn con esos datos", vbInformation, Me.Caption
        Else
           If chkDetalle.Value = 0 Then
              rsRecetas.Sort = "tipo,servicios"
              lnTotalG = 0
              rsRecetas.MoveFirst
              Do While Not rsRecetas.EOF
                 lcFuente = rsRecetas!tipo
                 lnTotal = 0
                 Do While Not rsRecetas.EOF And lcFuente = rsRecetas!tipo
                    lnTotal = lnTotal + rsRecetas!Cantidad
                    rsRecetas.MoveNext
                    If rsRecetas.EOF Then
                       Exit Do
                    End If
                 Loop
                 rsRecetas.MovePrevious
                 rsRecetas.Fields!canttotal = lnTotal
                 rsRecetas.Update
                 rsRecetas.MoveNext
                 lnTotalG = lnTotalG + lnTotal
              Loop
              'rsRecetas.Sort = ""
              rsRecetas.AddNew
              rsRecetas.Fields!tipo = ">>"
              rsRecetas.Fields!Servicios = ">>"
              rsRecetas.Fields!Cantidad = lnTotalG
              rsRecetas.Fields!canttotal = lnTotalG
           End If
           If Me.chkExcel.Value = 1 Then
              mo_ReglasReportes.ExportarRecordSetAexcel rsRecetas, optRecetasXplan.Caption, ml_TextoDelFiltro, _
                                                        "", Me.hwnd, True, True
           Else
           End If
        End If

        Set rsRecetas = Nothing
        Set rsFuenteFinan = Nothing
        Set rsServicios = Nothing
        Set rsTmp = Nothing
        Set mo_ReglasComunes = Nothing
        Set mo_ReglasFacturacion = Nothing
        Set oConexion = Nothing
        Set mo_ReglasReportes = Nothing


End Sub

