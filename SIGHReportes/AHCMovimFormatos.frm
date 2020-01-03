VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form AHCMovimFormatos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Historias para enviar al ARCHIVO PASIVO"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   Icon            =   "AHCMovimFormatos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   1050
      Left            =   0
      TabIndex        =   12
      Top             =   1530
      Width           =   9300
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "AHCMovimFormatos.frx":0CCA
         DownPicture     =   "AHCMovimFormatos.frx":118E
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
         Left            =   4860
         Picture         =   "AHCMovimFormatos.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AHCMovimFormatos.frx":1B66
         DownPicture     =   "AHCMovimFormatos.frx":1FC6
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
         Left            =   3330
         Picture         =   "AHCMovimFormatos.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1485
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9300
      Begin VB.CheckBox chkConActivas 
         Caption         =   "Incluir historias ACTIVAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   165
         TabIndex        =   19
         Top             =   660
         Width           =   3705
      End
      Begin VB.ComboBox cmbIdEstadoHistoria 
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
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   180
         Width           =   4770
      End
      Begin VB.TextBox txtAnios 
         Enabled         =   0   'False
         Height          =   360
         Left            =   4395
         TabIndex        =   11
         Top             =   930
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "años sin movimientos"
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
         Left            =   4860
         TabIndex        =   18
         Top             =   1005
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Historias clínicas que pasan de"
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
         Left            =   1860
         TabIndex        =   17
         Top             =   1005
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Label lblIdEstadoHistoria 
         AutoSize        =   -1  'True
         Caption         =   "Estado de historia"
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
         Left            =   180
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblParametro559 
         AutoSize        =   -1  'True
         Caption         =   "Años que la HISTORIA pasa a PASIVO según RS MINSA XXXXX"
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
         Left            =   165
         TabIndex        =   10
         Top             =   1005
         Visible         =   0   'False
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   870
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   285
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
         Left            =   120
         Picture         =   "AHCMovimFormatos.frx":28B0
         TabIndex        =   8
         Top             =   900
         Width           =   1755
      End
      Begin VB.ComboBox cmbConsiderar 
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
         ItemData        =   "AHCMovimFormatos.frx":2BC2
         Left            =   1680
         List            =   "AHCMovimFormatos.frx":2BCF
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   225
         Width           =   3570
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   570
         Width           =   1395
         _ExtentX        =   2461
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F. Movimiento"
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
         Left            =   150
         TabIndex        =   4
         Top             =   615
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Especialidad"
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
         Left            =   105
         TabIndex        =   3
         Top             =   285
         Width           =   1380
      End
   End
   Begin VB.Frame Frame3 
      Height          =   420
      Left            =   1695
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   240
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AHCMovimFormatos.frx":2C05
         DownPicture     =   "AHCMovimFormatos.frx":3065
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
         Left            =   1320
         Picture         =   "AHCMovimFormatos.frx":34DA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "AHCMovimFormatos.frx":394F
         DownPicture     =   "AHCMovimFormatos.frx":3E13
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
         Left            =   2850
         Picture         =   "AHCMovimFormatos.frx":42FF
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   210
         Width           =   1365
      End
   End
End
Attribute VB_Name = "AHCMovimFormatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Movimiento de Formatos
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_cmbIdEstadoHistoria As New sighentidades.ListaDespleglable
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim oRsHistoriasNormas As New Recordset
Dim lnSoloPasivas As Long
Dim lnIdEstado As Integer

       
Private Sub btnAceptar_Click()


        If Me.txtFechaInicio = sighentidades.FECHA_VACIA_DMY Then
            MsgBox "Ingrese la fecha de movimiento", vbInformation, Me.Caption
            Exit Sub
        Else
            If Not sighentidades.EsFecha(Me.txtFechaInicio, "DD/MM/AAAA") Then
                MsgBox "La fecha de movimiento no tiene el formato correcto", vbInformation, Me.Caption
                Exit Sub
            End If
        End If

        Me.MousePointer = 11
        Dim oRptClaseCry As New rCrystal
        oRptClaseCry.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
        oRptClaseCry.FechaInicio = Format(txtFechaInicio.Text & " 00:00:01", sighentidades.DevuelveFechaSoloFormato_DMY_HMS)
        oRptClaseCry.FechaFin = Format(txtFechaInicio.Text & " 23:59:59", sighentidades.DevuelveFechaSoloFormato_DMY_HMS)
        oRptClaseCry.TipoServicioHosp = IIf(cmbConsiderar.ListIndex = 0, "(3)", IIf(cmbConsiderar.ListIndex = 1, "(2,4)", "(1)"))
        oRptClaseCry.TextoDelFiltro = "Fecha Movimiento: " & txtFechaInicio.Text & "     " & IIf(cmbConsiderar.ListIndex = 0, "(Hospitalización)", IIf(cmbConsiderar.ListIndex = 1, "(Emergencia)", "(Consultorios Externos)"))
        oRptClaseCry.TipoReporte = Me.Name
        oRptClaseCry.Show vbModal
        Set oRptClaseCry = Nothing
        Me.MousePointer = 1
    
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Sub ActivaDesactivaControles()
    lblParametro559.Visible = IIf(chkConActivas.Value = 1, True, False)
    Label.Visible = IIf(chkConActivas.Value = 1, True, False)
    txtAnios.Visible = IIf(chkConActivas.Value = 1, True, False)
    Label1.Visible = IIf(chkConActivas.Value = 1, True, False)

End Sub

Private Sub chkConActivas_Click()
    ActivaDesactivaControles
    If chkConActivas.Value = 1 Then
       CambiaDeDiasEnNormaMINSA
    End If
End Sub

Private Sub cmbIdEstadoHistoria_Change()
    CambiaDeDiasEnNormaMINSA
End Sub

Private Sub cmbIdEstadoHistoria_Click()
    CambiaDeDiasEnNormaMINSA
End Sub

Private Sub cmdAceptar_Click()
   On Error GoTo errAcp
   Me.MousePointer = 11
   cmdAceptar.Enabled = False
   Dim oRsTmp1 As New Recordset
   Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
   Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
   Set oRsTmp1 = mo_ReglasArchivoClinico.ListaHistoriasParaPasivo(lnIdEstado)
   If oRsTmp1.RecordCount > 0 Then
      mo_ReglasReportes.ExportarRecordSetAexcel oRsTmp1, "HISTORIAS CLINICAS CON ESTADO: " & _
                        Mid(cmbIdEstadoHistoria.Text, 4), _
                        Me.lblParametro559.Caption & " : " & Label.Caption & Me.txtAnios.Text & Label1.Caption, _
                        "N° Historias: " & Trim(Str(oRsTmp1.RecordCount)), Me.hwnd, False, True
   Else
      MsgBox "No hay datos", vbInformation, ""
   End If
errAcp:
   Set oRsTmp1 = Nothing
   Set mo_ReglasReportes = Nothing
   Set mo_ReglasArchivoClinico = Nothing
   Me.MousePointer = 1
   cmdAceptar.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
    Me.Visible = False
End Sub

Private Sub Form_Load()
    Me.txtFechaInicio.Text = Date
    cmbConsiderar.ListIndex = 0
    
    'lblParametro559.Caption = lcBuscaParametro.SeleccionaFilaParametro(559)
    'txtAnios.Text = lcBuscaParametro.SeleccionaFilaParametroValorInt(559)
    
    Set oRsHistoriasNormas = mo_AdminArchivoClinico.HistoriasnormasSeleccionarTodos
    
    Set mo_cmbIdEstadoHistoria.MiComboBox = cmbIdEstadoHistoria
    mo_cmbIdEstadoHistoria.BoundColumn = "IdEstadoHistoria"
    mo_cmbIdEstadoHistoria.ListField = "DescripcionLarga"
    Set mo_cmbIdEstadoHistoria.RowSource = mo_AdminArchivoClinico.EstadosHistoriaClinicaSeleccionarTodos()
    mo_cmbIdEstadoHistoria.BoundText = sghEstadosHistoria.sghDepurada
    
    CambiaDeDiasEnNormaMINSA
    
    
End Sub

Sub CambiaDeDiasEnNormaMINSA()
    lblParametro559.Caption = oRsHistoriasNormas!NoResolucion
    Label.Caption = "Historias clínicas que pasan de "
    Label1.Caption = " años de su último movimiento"
    chkConActivas.Enabled = True
    lnIdEstado = Val(mo_cmbIdEstadoHistoria.BoundText)
    Select Case Val(mo_cmbIdEstadoHistoria.BoundText)
    Case sghEstadosHistoria.sghActiva
         chkConActivas.Value = 0
         chkConActivas.Enabled = False
         ActivaDesactivaControles
         Label.Caption = "Historias clínicas menores a "
         Label1.Caption = " años de su último movimiento"
         txtAnios.Text = lnSoloPasivas
    Case sghEstadosHistoria.sghDepurada
         txtAnios.Text = oRsHistoriasNormas!AnioPasivo_N1
         lnSoloPasivas = Val(txtAnios.Text)
         If chkConActivas.Value = 0 Then
            lblParametro559.Caption = ""
            Label.Caption = ""
            Label1.Caption = ""
            txtAnios.Text = ""
         Else
            lnIdEstado = 4              'solo pasivas - sin considerar ACTIVAS
         End If
    Case sghEstadosHistoria.sghDepuradaXeliminar
         txtAnios.Text = oRsHistoriasNormas!AnioElimin_N2
         If chkConActivas.Value = 0 Then
            lblParametro559.Caption = ""
            Label.Caption = ""
            Label1.Caption = ""
            txtAnios.Text = ""
         Else
            lnIdEstado = 5              'solo pasivas x eliminar- sin considerar ACTIVAS
         End If
    End Select
End Sub

Private Sub txtFechaInicio_LostFocus()
    If txtFechaInicio <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFechaInicio, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaInicio = sighentidades.FECHA_VACIA_DMY
        End If
    End If
End Sub
