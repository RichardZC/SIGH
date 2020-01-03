VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form rMontosPorPlan 
   Caption         =   "Montos según Plan"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9300
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rMontosPorPlan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   9300
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
      Left            =   90
      TabIndex        =   5
      Top             =   2220
      Width           =   9180
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rMontosPorPlan.frx":0CCA
         DownPicture     =   "rMontosPorPlan.frx":118E
         Height          =   700
         Left            =   4740
         Picture         =   "rMontosPorPlan.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rMontosPorPlan.frx":1B66
         DownPicture     =   "rMontosPorPlan.frx":1FC6
         Height          =   700
         Left            =   3210
         Picture         =   "rMontosPorPlan.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
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
      Height          =   2175
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   9195
      Begin VB.CheckBox chkConsideraIME 
         Caption         =   "Considerar el formato IME"
         Height          =   255
         Left            =   1350
         TabIndex        =   15
         Top             =   1680
         Value           =   1  'Checked
         Width           =   5835
      End
      Begin VB.CheckBox chkReembolsos 
         Caption         =   "Considerar REEMBOLSOS (registrados el mes y año de F.Movimiento)"
         Height          =   255
         Left            =   1350
         TabIndex        =   14
         Top             =   1380
         Value           =   1  'Checked
         Width           =   5955
      End
      Begin VB.CheckBox chkTodasFarmacias 
         Caption         =   "Todos las Farmacias"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Value           =   1  'Checked
         Width           =   1965
      End
      Begin VB.CheckBox chkRecalculo 
         Caption         =   "Considerar RECALCULOS"
         Height          =   255
         Left            =   1350
         TabIndex        =   12
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2625
      End
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "En Excel"
         Height          =   315
         Left            =   7770
         Picture         =   "rMontosPorPlan.frx":28B0
         TabIndex        =   11
         Top             =   1050
         Width           =   1095
      End
      Begin VB.ComboBox cmbAlmacen 
         Height          =   330
         Left            =   2130
         TabIndex        =   0
         Top             =   240
         Visible         =   0   'False
         Width           =   6720
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   1350
         TabIndex        =   1
         Top             =   660
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
         Left            =   6720
         TabIndex        =   2
         Top             =   660
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
         TabIndex        =   9
         Top             =   660
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
         TabIndex        =   10
         Top             =   660
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
         Height          =   210
         Left            =   6210
         TabIndex        =   8
         Top             =   690
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Movimiento"
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   690
         Width           =   1080
      End
   End
End
Attribute VB_Name = "rMontosPorPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte Montos por Plan
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_MensajeError As String
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_TextoDelFiltro As String
Dim mo_cmbAlmacen As New sighentidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim ml_idUsuario As Long
Dim mo_Formulario As New sighentidades.Formulario
Dim lcBuscaParametro As New SIGHDatos.Parametros

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Private Sub chkTodasFarmacias_Click()
    If chkTodasFarmacias.Value = 1 Then
       cmbAlmacen.Visible = False
    Else
       cmbAlmacen.Visible = True
    End If
End Sub

Private Sub cmbAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmacen

End Sub

Private Sub Form_Initialize()
    Set mo_cmbAlmacen.MiComboBox = cmbAlmacen
End Sub

Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
            Dim oRptClaseCry As New rCrystal
            oRptClaseCry.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
            oRptClaseCry.IdAlmacen = IIf(chkTodasFarmacias.Value = 1, 0, Val(mo_cmbAlmacen.BoundText))
            oRptClaseCry.FechaInicio = CDate(Format(txtFdesde.Text & " " & txtHrInicio & ":00", sighentidades.DevuelveFechaSoloFormato_DMY_HMS))
            oRptClaseCry.FechaFin = CDate(Format(txtFhasta.Text & " " & txtHrFin & ":00", sighentidades.DevuelveFechaSoloFormato_DMY_HMS))
            oRptClaseCry.HoraInicio = txtHrInicio.Text
            oRptClaseCry.HoraFin = txtHrFin.Text
            oRptClaseCry.TextoDelFiltro = ml_TextoDelFiltro
            oRptClaseCry.TipoReporte = Me.Name
            oRptClaseCry.ConsiderarRecalculo = IIf(chkRecalculo.Value = 1, True, False)
            oRptClaseCry.ConsiderarReembolsos = IIf(Me.chkReembolsos.Value = 1, True, False)
            oRptClaseCry.Show vbModal
            Set oRptClaseCry = Nothing
            If Me.chkConsideraIME.Value = 1 Then
               ImprimeFormatoIME
            End If
            
        Me.MousePointer = 1
        
    End If
End Sub


Sub ImprimeFormatoIME()
    Dim oRsIME As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim lcSql As String
    Dim lcBoletaSerie As String, lcBoletaDel As String, lcBoletaAl As String, lnBoletaAnulados As Integer
    Dim lcFacturaSerie As String, lcFacturaDel As String, lcFacturaAl As String, lnFacturaAnulados As Integer
    With oRsIME
          .Fields.Append "Grupo", adVarChar, 30, adFldIsNullable
          .Fields.Append "BoletaSerie", adVarChar, 5, adFldIsNullable
          .Fields.Append "BoletaDel", adVarChar, 20, adFldIsNullable
          .Fields.Append "BoletaAl", adVarChar, 20, adFldIsNullable
          .Fields.Append "BoletaAnulados", adInteger
          .Fields.Append "FacturaSerie", adVarChar, 5, adFldIsNullable
          .Fields.Append "FacturaDel", adVarChar, 20, adFldIsNullable
          .Fields.Append "FacturaAl", adVarChar, 20, adFldIsNullable
          .Fields.Append "FacturaAnulados", adVarChar, 5, adFldIsNullable
          .Fields.Append "NiFecha", adDate, 10, adFldIsNullable
          .Fields.Append "NiGuiaRemision", adVarChar, 20, adFldIsNullable
          .Fields.Append "NiImporte", adDouble
          .LockType = adLockOptimistic
          .Open
    End With
    '
    lcBoletaSerie = "": lcBoletaDel = "": lcBoletaAl = "": lnBoletaAnulados = 0
    lcFacturaSerie = "": lcFacturaDel = "": lcFacturaAl = "": lnFacturaAnulados = 0
    
    Set oRsTmp1 = mo_ReglasCaja.CajaComprobantesPagoSeleccionarPorFechasSoloFarmacia(Me.txtFdesde.Text, Me.txtHrInicio.Text, Me.txtFhasta.Text, Me.txtHrFin.Text, chkTodasFarmacias.Value, Val(mo_cmbAlmacen.BoundText))
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       If oRsTmp1.Fields!IdEstadoComprobante = 4 Then
            If oRsTmp1.Fields!IdTipoComprobante = 3 Then
               lcBoletaSerie = oRsTmp1.Fields!NroSerie
               lcBoletaDel = oRsTmp1.Fields!NroDocumento
            Else
               lcFacturaSerie = oRsTmp1.Fields!NroSerie
               lcFacturaDel = oRsTmp1.Fields!NroDocumento
            End If
       End If
       Do While Not oRsTmp1.EOF
            If oRsTmp1.Fields!IdEstadoComprobante = 4 Then
                 If oRsTmp1.Fields!IdTipoComprobante = 3 Then
                    lcBoletaAl = oRsTmp1.Fields!NroDocumento
                 Else
                    lcFacturaAl = oRsTmp1.Fields!NroDocumento
                 End If
            Else
                 If oRsTmp1.Fields!IdTipoComprobante = 3 Then
                    lnBoletaAnulados = lnBoletaAnulados + 1
                 Else
                    lnFacturaAnulados = lnFacturaAnulados + 1
                 End If
            End If
            oRsTmp1.MoveNext
       Loop
    End If
    oRsTmp1.Close
    oRsIME.AddNew
    oRsIME.Fields!Grupo = "Boleta"
    oRsIME.Fields!BoletaSerie = lcBoletaSerie
    oRsIME.Fields!BoletaDel = lcBoletaDel
    oRsIME.Fields!BoletaAl = lcBoletaAl
    oRsIME.Fields!BoletaAnulados = lnBoletaAnulados
    oRsIME.Update
    oRsIME.AddNew
    oRsIME.Fields!Grupo = "Factura"
    oRsIME.Fields!FacturaSerie = lcFacturaSerie
    oRsIME.Fields!FacturaDel = lcFacturaDel
    oRsIME.Fields!FacturaAl = lcFacturaAl
    oRsIME.Fields!FacturaAnulados = lnFacturaAnulados
    oRsIME.Update
    '
    Set oRsTmp1 = mo_ReglasFarmacia.FarmMovimientoSeleccionarPorFechasCabecera(Me.txtFdesde.Text, Me.txtHrInicio.Text, Me.txtFhasta.Text, Me.txtHrFin.Text, chkTodasFarmacias.Value, Val(mo_cmbAlmacen.BoundText))
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
            oRsIME.AddNew
            oRsIME.Fields!Grupo = "GuiaRemision"
            oRsIME.Fields!NiFecha = oRsTmp1.Fields!FechaCreacion
            oRsIME.Fields!NiGuiaRemision = oRsTmp1.Fields!DocumentoNumero
            oRsIME.Fields!NiImporte = oRsTmp1.Fields!total
            oRsIME.Update
            oRsTmp1.MoveNext
       Loop
    Else
            oRsIME.AddNew
            oRsIME.Fields!Grupo = "GuiaRemision"
            'oRsIME.Fields!NiFecha = ""
            oRsIME.Fields!NiGuiaRemision = ""
            oRsIME.Fields!NiImporte = 0
            oRsIME.Update
    End If
    oRsTmp1.Close
    '
    Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
    mo_ReglasReportes.ExportarRecordSetAexcel oRsIME, "Formato IME", ml_TextoDelFiltro, "", Me.hwnd
    '
    Set oRsTmp1 = Nothing
    Set oRsIME = Nothing
End Sub


Function ValidaDatosObligatorios() As Boolean
    ms_MensajeError = ""
    ml_TextoDelFiltro = "FILTROS:    Farmacia: " & IIf(chkTodasFarmacias.Value = 1, "Todas", Trim(cmbAlmacen.Text)) & "    F.Movimiento: (" & txtFdesde.Text & " " & txtHrInicio.Text & "   al " & txtFhasta.Text & " " & txtHrFin.Text & ") " & _
                        IIf(chkRecalculo.Value = 1, "(Con Recalculo)", "(Sin Recalculo)") & IIf(Me.chkReembolsos.Value = 1, "(Considera REEMBOLSOS)", "(Sin REEMBOLSOS)")
    If chkTodasFarmacias.Value = 0 And cmbAlmacen.Text = "" Then
        ms_MensajeError = ms_MensajeError + "Por favor elija el Almacén" + Chr(13)
        cmbAlmacen.SetFocus
    End If
    If CDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) > CDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
       Exit Function
    End If
    If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError, vbInformation, Me.Caption
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



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
'           ucListaProductos1.RealizarBusqueda
        Case vbKeyEscape
'           btnCancelar_Click
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
If Not sighentidades.esfecha(txtFdesde.Text, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
        End If

End Sub

Private Sub txtFhasta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFhasta

End Sub


Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_ReglasFarmacia = Nothing
    Set mo_Teclado = Nothing
    Set mo_cmbAlmacen = Nothing
    Set mo_Formulario = Nothing
End Sub

Private Sub txtFhasta_LostFocus()
If Not sighentidades.esfecha(txtFhasta.Text, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
        End If
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
