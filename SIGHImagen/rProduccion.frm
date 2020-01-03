VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form rProduccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Producción por Fechas"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   Icon            =   "rProduccion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2160
      Left            =   45
      TabIndex        =   7
      Top             =   15
      Width           =   11235
      Begin VB.Frame FraCpt 
         Caption         =   "Tipo Reporte"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   7230
         TabIndex        =   17
         Top             =   1470
         Width           =   3840
         Begin Threed.SSOption opttipoServ 
            Height          =   255
            Left            =   135
            TabIndex        =   18
            Top             =   255
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   450
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Tipo Servicio"
            Value           =   -1
         End
         Begin Threed.SSOption optFinanc 
            Height          =   255
            Left            =   1815
            TabIndex        =   19
            Top             =   255
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   450
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Fte.Financiamiento"
         End
      End
      Begin VB.CheckBox chkCPT 
         Alignment       =   1  'Right Justify
         Caption         =   "A nivel de Procedimientos CPT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8160
         Picture         =   "rProduccion.frx":0CCA
         TabIndex        =   15
         Top             =   1140
         Value           =   1  'Checked
         Width           =   2865
      End
      Begin VB.ComboBox cmbResponsable 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8070
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   660
         Width           =   3000
      End
      Begin VB.ComboBox cmbIdPuntoDeCarga 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   660
         Width           =   2235
      End
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "En Excel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         Picture         =   "rProduccion.frx":0FDC
         TabIndex        =   10
         Top             =   1140
         Width           =   1425
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   240
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
         Left            =   8910
         TabIndex        =   2
         Top             =   210
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
         Left            =   2760
         TabIndex        =   1
         Top             =   240
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
         Left            =   10290
         TabIndex        =   3
         Top             =   210
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Responsable"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7020
         TabIndex        =   14
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pto. Carga"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8400
         TabIndex        =   9
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Movimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   8
         Top             =   270
         Width           =   1080
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   45
      TabIndex        =   5
      Top             =   2130
      Width           =   11220
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rProduccion.frx":12EE
         DownPicture     =   "rProduccion.frx":174E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   4193
         Picture         =   "rProduccion.frx":1BC3
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rProduccion.frx":2038
         DownPicture     =   "rProduccion.frx":24FC
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   5723
         Picture         =   "rProduccion.frx":29E8
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdAlertaCantidades 
      Height          =   2760
      Left            =   45
      TabIndex        =   16
      Top             =   3285
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   4868
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lista de MOVIMIENTOS registrados para el chequeo de CANTIDADES"
   End
End
Attribute VB_Name = "rProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte de producción
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_cmbIdPuntoCarga As New SIGHEntidades.ListaDespleglable
Dim mo_cmbResponsable As New SIGHEntidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
Dim sMensaje As String
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim ml_TextoDelFiltro As String
Dim lnIdProducto As Long
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim ml_idUsuario As Long

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRptClaseCry As New rCrystal
        oRptClaseCry.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
        oRptClaseCry.FechaInicio = Format(txtFdesde.Text & " " & txtHrInicio.Text, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
        oRptClaseCry.FechaFin = Format(txtFhasta.Text & " " & txtHrFin.Text, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HM)
        oRptClaseCry.TextoDelFiltro = ml_TextoDelFiltro
        oRptClaseCry.IdPuntoCarga = Val(mo_cmbIdPuntoCarga.BoundText)
        oRptClaseCry.TipoReporte = IIf(opttipoServ.Value = True, Me.Name, "rProduccionTF")
        oRptClaseCry.IdResponsable = Val(mo_cmbResponsable.BoundText)
        oRptClaseCry.EnResumen = IIf(chkCPT.Value = 1, True, False)
        oRptClaseCry.Show vbModal
        Set oRptClaseCry = Nothing
        Me.MousePointer = 1
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    If CDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) > CDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
       Exit Function
    End If
    sMensaje = ""
    ml_TextoDelFiltro = "FILTROS:   " & Trim(cmbIdPuntoDeCarga.Text) & "      F.Movimiento: (" & txtFdesde.Text & " " & txtHrInicio.Text & "   al " & txtFhasta.Text & " " & txtHrFin.Text & ")"
    If cmbResponsable.Text <> "" Then
       ml_TextoDelFiltro = ml_TextoDelFiltro & "     Responsable: " & cmbResponsable.Text
    End If
    If cmbIdPuntoDeCarga.Text = "" Then
        sMensaje = sMensaje + "Por favor elija el Punto de Carga" + Chr(13)
        cmbIdPuntoDeCarga.SetFocus
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






Private Sub chkCPT_Click()
    If chkCPT.Value = 1 Then
       FraCpt.Visible = True
    Else
       FraCpt.Visible = False
    End If
End Sub

Private Sub cmbIdPuntoDeCarga_Click()
    mo_cmbResponsable.BoundColumn = "idEmpleado"
    mo_cmbResponsable.ListField = "ApNom"
    Set mo_cmbResponsable.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =" & mo_ReglasFarmacia.EmpleadosDevuelveIdCargoSegunPuntoCarga(Val(mo_cmbIdPuntoCarga.BoundText)))
End Sub

Private Sub cmbIdPuntoDeCarga_KeyDown(KeyCode As Integer, Shift As Integer)
     mo_Teclado.RealizarNavegacion KeyCode, cmbIdPuntoDeCarga
End Sub


Private Sub Form_Initialize()
    Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPuntoDeCarga
    Set mo_cmbResponsable.MiComboBox = cmbResponsable
End Sub

Sub InicializaFechaHora()
    txtFdesde.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual
    txtFhasta.Text = Date
    txtHrInicio.Text = "00:00"
    txtHrFin.Text = "23:59"

End Sub

Private Sub Form_Load()
    InicializaFechaHora
    '
    mo_cmbIdPuntoCarga.ListField = "Descripcion"
    mo_cmbIdPuntoCarga.BoundColumn = "IdPuntoCarga"
    Set mo_cmbIdPuntoCarga.RowSource = mo_reglasComunes.SeleccionarPuntosDeCargaSegunFiltro("idUPS=1")
    Dim rsIdAlmacen As Recordset
    Set rsIdAlmacen = mo_reglasComunes.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghImageneología, ml_idUsuario)
    If rsIdAlmacen.RecordCount > 0 Then
       mo_cmbIdPuntoCarga.BoundText = rsIdAlmacen.Fields!idLaboraSubArea
       mo_Formulario.HabilitarDeshabilitar cmbIdPuntoDeCarga, False
       cmbIdPuntoDeCarga_Click
    End If
    '
    Set grdAlertaCantidades.DataSource = mo_ReglasImagenes.ImagenesDevuelveAlertaDeCantidades
    mo_Apariencia.ConfigurarFilasBiColores grdAlertaCantidades, SIGHEntidades.GrillaConFilasBicolor
    '
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

Private Sub txtFdesde_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFdesde

End Sub



Private Sub txtFdesde_LostFocus()
    If txtFdesde <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFdesde, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
        End If
    End If

End Sub

Private Sub txtFhasta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFhasta

End Sub

Private Sub txtFhasta_LostFocus()
    If txtFhasta <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFhasta, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
        End If
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_ReglasFarmacia = Nothing
    Set mo_Teclado = Nothing
    Set mo_cmbIdPuntoCarga = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_reglasComunes = Nothing
    Set mo_Formulario = Nothing
End Sub

Private Sub txtHrFin_LostFocus()
If Not SIGHEntidades.ValidaHora(txtHrFin.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            InicializaFechaHora
        End If
End Sub

Private Sub txtHrInicio_LostFocus()
If Not SIGHEntidades.ValidaHora(txtHrInicio.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            InicializaFechaHora
        End If
End Sub
