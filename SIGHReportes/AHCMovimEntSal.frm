VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form AHCMovimEntSal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimiento de Historias"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "AHCMovimEntSal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   5
      Top             =   1530
      Width           =   5370
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AHCMovimEntSal.frx":0CCA
         DownPicture     =   "AHCMovimEntSal.frx":112A
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
         Picture         =   "AHCMovimEntSal.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "AHCMovimEntSal.frx":1A14
         DownPicture     =   "AHCMovimEntSal.frx":1ED8
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
         Picture         =   "AHCMovimEntSal.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1485
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   5370
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
         Picture         =   "AHCMovimEntSal.frx":28B0
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
         ItemData        =   "AHCMovimEntSal.frx":2BC2
         Left            =   1680
         List            =   "AHCMovimEntSal.frx":2BCF
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
End
Attribute VB_Name = "AHCMovimEntSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Movimiento de Historias
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Private Sub btnAceptar_Click()
        If Me.txtFechaInicio = SIGHEntidades.FECHA_VACIA_DMY Then
            MsgBox "Ingrese la fecha de movimiento", vbInformation, Me.Caption
            Exit Sub
        Else
            If Not SIGHEntidades.EsFecha(Me.txtFechaInicio, "DD/MM/AAAA") Then
                MsgBox "La fecha de movimiento no tiene el formato correcto", vbInformation, Me.Caption
                Exit Sub
            End If
        End If

        Me.MousePointer = 11
        Dim oRptClaseCry As New rCrystal
        oRptClaseCry.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
        oRptClaseCry.FechaInicio = Format(txtFechaInicio.Text & " 00:00:01", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS)
        oRptClaseCry.FechaFin = Format(txtFechaInicio.Text & " 23:59:59", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS)
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

Private Sub Form_Load()
    Me.txtFechaInicio.Text = Date
    cmbConsiderar.ListIndex = 2
End Sub



Private Sub txtFechaInicio_LostFocus()
    If txtFechaInicio <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFechaInicio, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaInicio = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub
