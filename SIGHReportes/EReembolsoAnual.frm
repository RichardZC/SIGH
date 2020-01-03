VERSION 5.00
Begin VB.Form EReembolsoAnual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reembolso Anual"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "EReembolsoAnual.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   1
      Top             =   1890
      Width           =   5370
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "EReembolsoAnual.frx":0CCA
         DownPicture     =   "EReembolsoAnual.frx":112A
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
         Picture         =   "EReembolsoAnual.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "EReembolsoAnual.frx":1A14
         DownPicture     =   "EReembolsoAnual.frx":1ED8
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
         Picture         =   "EReembolsoAnual.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1845
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   5370
      Begin VB.ComboBox cmbAreaTramitaR 
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
         Left            =   1740
         TabIndex        =   7
         Top             =   150
         Width           =   3450
      End
      Begin VB.ComboBox cmbFuenteFinanciamiento 
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
         Left            =   1740
         TabIndex        =   6
         Top             =   510
         Width           =   3450
      End
      Begin VB.ComboBox cmbAnio 
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
         ItemData        =   "EReembolsoAnual.frx":28B0
         Left            =   1740
         List            =   "EReembolsoAnual.frx":28B2
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   870
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Por Fte.Financ/IAFA"
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
         TabIndex        =   9
         Top             =   570
         Width           =   1635
      End
      Begin VB.Label Label17 
         Caption         =   "Area que Tramita"
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
         TabIndex        =   8
         Top             =   210
         Width           =   1515
      End
      Begin VB.Label Departamento 
         AutoSize        =   -1  'True
         Caption         =   "Año"
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
         Left            =   60
         TabIndex        =   5
         Top             =   930
         Width           =   330
      End
   End
End
Attribute VB_Name = "EReembolsoAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reembolso Anual
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Dim mo_cmbAreaTramitaR As New SIGHEntidades.ListaDespleglable
Dim mo_cmbFuenteFinanciamiento As New SIGHEntidades.ListaDespleglable
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim ml_idUsuario As Long
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Private Sub Form_Initialize()
    Set mo_cmbFuenteFinanciamiento.MiComboBox = cmbFuenteFinanciamiento
    Set mo_cmbAreaTramitaR.MiComboBox = cmbAreaTramitaR
End Sub

Private Sub btnAceptar_Click()
        Me.MousePointer = 11
        Dim oRpt As New RptERembolsoAnual
        Dim lcTextoFiltro As String
        lcTextoFiltro = "Año: " & cmbAnio.Text
        If cmbAreaTramitaR.Text <> "" Then
           lcTextoFiltro = lcTextoFiltro & "     Area Tramita Seguros: " & Trim(cmbAreaTramitaR.Text)
        End If
        If cmbFuenteFinanciamiento.Text <> "" Then
           lcTextoFiltro = lcTextoFiltro & "     IAFA: " & cmbFuenteFinanciamiento.Text
        End If
        oRpt.CrearReporte_excel lcTextoFiltro, Val(cmbAnio.Text), Val(mo_cmbAreaTramitaR.BoundText), Val(mo_cmbFuenteFinanciamiento.BoundText), Me.hwnd
        Set oRpt = Nothing
        Me.MousePointer = 1
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Private Sub Form_Load()
       mo_Formulario.LlenaComboConAnios cmbAnio
       CargarComboBoxes
End Sub

Sub CargarComboBoxes()
       mo_cmbAreaTramitaR.ListField = "Descripcion"
       mo_cmbAreaTramitaR.BoundColumn = "idAreaTramitaSeguros"
       Set mo_cmbAreaTramitaR.RowSource = mo_ReglasFacturacion.AreaTramitaSegurosDevuelveTodosSegunFiltro("")
       Dim rsPuntoCargaDondeLabora As Recordset
       Set rsPuntoCargaDondeLabora = mo_ReglasComunes.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAreaTramitaSeguros, ml_idUsuario)
       If rsPuntoCargaDondeLabora.RecordCount > 0 Then
           mo_cmbAreaTramitaR.BoundText = rsPuntoCargaDondeLabora.Fields!idLaboraSubArea
           mo_Formulario.HabilitarDeshabilitar cmbAreaTramitaR, False
           cmbAreaTramitaR_Click
       End If
End Sub
Private Sub cmbAreaTramitaR_Click()
       mo_cmbFuenteFinanciamiento.ListField = "Descripcion"
       mo_cmbFuenteFinanciamiento.BoundColumn = "idFuenteFinanciamiento"
       If Val(mo_cmbAreaTramitaR.BoundText) = 4 Then
          '**************Referencia************
          Set mo_cmbFuenteFinanciamiento.RowSource = mo_ReglasFacturacion.FuentesFinanciamientoDevuelveTodosSegunFiltro(" utilizadoEn=3 ")
       Else
          Set mo_cmbFuenteFinanciamiento.RowSource = mo_ReglasFacturacion.FuentesFinanciamientoDevuelveTodosSegunFiltro(" idAreaTramitaSeguros= " & mo_cmbAreaTramitaR.BoundText)
       End If

End Sub
