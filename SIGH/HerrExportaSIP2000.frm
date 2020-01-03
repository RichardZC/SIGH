VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form HerrExportaSIP2000 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Norma de Historias Clínicas"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   Icon            =   "HerrExportaSIP2000.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Height          =   4590
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   15
      Width           =   8430
      Begin VB.Frame Frame 
         Height          =   3135
         Index           =   1
         Left            =   30
         TabIndex        =   17
         Top             =   120
         Width           =   8355
         Begin VB.ComboBox cmbInstitucion 
            Enabled         =   0   'False
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
            ItemData        =   "HerrExportaSIP2000.frx":0CCA
            Left            =   5985
            List            =   "HerrExportaSIP2000.frx":0CD1
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   345
            Width           =   2310
         End
         Begin VB.TextBox txtEliminacion 
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
            Left            =   7890
            MaxLength       =   2
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   2190
            Width           =   375
         End
         Begin VB.TextBox txtPasivo 
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
            Left            =   7905
            MaxLength       =   2
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   1785
            Width           =   375
         End
         Begin VB.TextBox txtNresolucion 
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
            Left            =   1665
            MaxLength       =   100
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   375
            Width           =   2520
         End
         Begin VB.TextBox txtDirectiva 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   1665
            MaxLength       =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Top             =   795
            Width           =   6615
         End
         Begin VB.ComboBox cmbEstado 
            Enabled         =   0   'False
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
            ItemData        =   "HerrExportaSIP2000.frx":0CDC
            Left            =   1665
            List            =   "HerrExportaSIP2000.frx":0CE6
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   2625
            Width           =   1605
         End
         Begin MSMask.MaskEdBox txtFnorma 
            Height          =   315
            Left            =   1665
            TabIndex        =   28
            Top             =   1830
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
         Begin MSMask.MaskEdBox txtFvigencia 
            Height          =   315
            Left            =   1665
            TabIndex        =   29
            Top             =   2220
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
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
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
            Left            =   135
            TabIndex        =   25
            Top             =   2670
            Width           =   555
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de vigencia"
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
            Left            =   135
            TabIndex        =   24
            Top             =   2235
            Width           =   1455
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Años de antiguedad para considerar eliminación"
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
            Left            =   3990
            TabIndex        =   23
            Top             =   2235
            Width           =   3870
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Años de antiguedad para entrar a pasivo"
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
            Left            =   4545
            TabIndex        =   22
            Top             =   1845
            Width           =   3315
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de la Norma"
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
            Left            =   135
            TabIndex        =   21
            Top             =   1845
            Width           =   1515
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Institución"
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
            Left            =   5085
            TabIndex        =   20
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Directiva o Norma Técnica"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   135
            TabIndex        =   19
            Top             =   810
            Width           =   1500
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "N° de Resolución"
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
            Left            =   135
            TabIndex        =   18
            Top             =   390
            Width           =   1395
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   30
         TabIndex        =   14
         Top             =   3270
         Width           =   8355
         Begin VB.CommandButton cmdSalir 
            Cancel          =   -1  'True
            Caption         =   "Cancelar (ESC)"
            DisabledPicture =   "HerrExportaSIP2000.frx":0CFF
            DownPicture     =   "HerrExportaSIP2000.frx":11C3
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   4020
            Picture         =   "HerrExportaSIP2000.frx":16AF
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   210
            Width           =   1335
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "Aceptar"
            DisabledPicture =   "HerrExportaSIP2000.frx":1B9B
            DownPicture     =   "HerrExportaSIP2000.frx":1FFB
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2576
            Picture         =   "HerrExportaSIP2000.frx":2470
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   210
            Width           =   1365
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consideraciones:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   975
      TabIndex        =   4
      Top             =   270
      Width           =   7860
      Begin VB.ListBox cmbConsideraciones 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   3420
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   7665
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   60
      TabIndex        =   5
      Top             =   3930
      Width           =   7845
      Begin MSMask.MaskEdBox txtHoraDesde 
         Height          =   315
         Left            =   2250
         TabIndex        =   6
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
      Begin MSMask.MaskEdBox txtFechaDesde 
         Height          =   315
         Left            =   810
         TabIndex        =   7
         Top             =   210
         Width           =   1380
         _ExtentX        =   2434
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
      Begin MSMask.MaskEdBox txtFechaHasta 
         Height          =   315
         Left            =   3960
         TabIndex        =   8
         Top             =   195
         Width           =   1380
         _ExtentX        =   2434
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
      Begin MSMask.MaskEdBox txtHoraHasta 
         Height          =   315
         Left            =   5385
         TabIndex        =   9
         Top             =   195
         Width           =   765
         _ExtentX        =   1349
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
      Begin SISGalenPlus.XP_ProgressBar progressRpt 
         Height          =   300
         Left            =   180
         TabIndex        =   12
         Top             =   990
         Width           =   7470
         _ExtentX        =   13176
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   6956042
      End
      Begin VB.Label lblFechaRequerida 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Left            =   3150
         TabIndex        =   11
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lblFechaSolicitud 
         Caption         =   "Desde"
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
         Left            =   210
         TabIndex        =   10
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   60
      TabIndex        =   3
      Top             =   5970
      Width           =   7845
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Exporta al SIP2000"
         DisabledPicture =   "HerrExportaSIP2000.frx":28E5
         DownPicture     =   "HerrExportaSIP2000.frx":2D45
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2576
         Picture         =   "HerrExportaSIP2000.frx":31BA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HerrExportaSIP2000.frx":362F
         DownPicture     =   "HerrExportaSIP2000.frx":3AF3
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4039
         Picture         =   "HerrExportaSIP2000.frx":3FDF
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   1335
      End
   End
End
Attribute VB_Name = "HerrExportaSIP2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Exporta información para el Sistema SIP2000
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim ml_IdUsuario As Long
Dim mo_lcNombrePc  As String

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Let IdUsuario(lIdValue As Long)
    ml_IdUsuario = lIdValue
End Property


Private Sub btnAceptar_Click()
    If MsgBox("Esta seguro", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        Me.MousePointer = 11
        Dim oProcesos As New Procesos
        Set oProcesos.progressRpt1 = Me.progressRpt
        oProcesos.IdUsuario = ml_IdUsuario
        oProcesos.lcNombrePc = mo_lcNombrePc
        oProcesos.ExportaSIP2000 Me.txtFechaDesde.Text, Me.txtFechaHasta.Text, Me.txtHoraDesde.Text, Me.txtHoraHasta.Text
        Set oProcesos = Nothing
        Unload Me
        Exit Sub
    End If
        
        
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub




Private Sub cmdAceptar_Click()
    If mo_ReglasArchivoClinico.HistoriasNormasModificar(txtNresolucion.Text, txtDirectiva.Text, cmbInstitucion.ListIndex, _
                               CDate(txtFnorma.Text), Val(txtPasivo.Text), Val(txtEliminacion.Text), _
                               CDate(txtFvigencia.Text), cmbEstado.ListIndex) = False Then
       MsgBox mo_ReglasArchivoClinico.MensajeError
    Else
       Me.Visible = False
    End If
    
End Sub

Private Sub cmdSalir_Click()
    Me.Visible = False
End Sub

Private Sub Form_Load()
    txtFechaDesde.Text = Date
    txtHoraDesde.Text = "00:01"
    txtFechaHasta.Text = Date
    txtHoraHasta.Text = Format(Now, "hh:mm")
    mo_reglasComunes.LlenaListBoxConTablaMensajesEnVentana cmbConsideraciones, "HerrExportaSIP2000"
    
    CargaHistoriasNormas
End Sub

Sub CargaHistoriasNormas()
    Dim oRsTmp1 As New Recordset
    Set oRsTmp1 = mo_ReglasArchivoClinico.HistoriasnormasSeleccionarTodos
    If oRsTmp1.RecordCount > 0 Then
        txtNresolucion.Text = oRsTmp1!NoResolucion
        cmbInstitucion.ListIndex = oRsTmp1!Institucion
        txtDirectiva.Text = oRsTmp1!NoDirectiva
        txtFnorma.Text = oRsTmp1!FechaNorma
        txtPasivo.Text = oRsTmp1!AnioPasivo_N1
        txtFvigencia.Text = oRsTmp1!FechaVigencia
        txtEliminacion.Text = oRsTmp1!AnioElimin_N2
        cmbEstado.ListIndex = oRsTmp1!estado
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
        Case vbKeyEscape
'           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub



Private Sub txtFechaDesde_LostFocus()
If Not EsFecha(txtFechaDesde.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaDesde.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtfechaHasta_LostFocus()
If Not EsFecha(txtFechaHasta.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFechaHasta.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtHoraDesde_LostFocus()
 If Not sighentidades.ValidaHora(txtHoraDesde) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
             txtHoraDesde = sighentidades.HORA_VACIA_HM
        End If
End Sub

Private Sub txtHoraHasta_LostFocus()
 If Not sighentidades.ValidaHora(txtHoraHasta) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
             txtHoraHasta = sighentidades.HORA_VACIA_HM
        End If
End Sub
