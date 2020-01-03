VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form EExoGeneral 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exoneraciones en General"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   Icon            =   "rExoGeneral.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   9225
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
      Height          =   2445
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Width           =   9195
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
         Left            =   7920
         Picture         =   "rExoGeneral.frx":0CCA
         TabIndex        =   10
         Top             =   660
         Width           =   1125
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
         Left            =   6900
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
         Left            =   8280
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
         Left            =   6390
         TabIndex        =   9
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Documento"
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
         Width           =   1110
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   5
      Top             =   2520
      Width           =   9180
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rExoGeneral.frx":0FDC
         DownPicture     =   "rExoGeneral.frx":143C
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
         Left            =   3210
         Picture         =   "rExoGeneral.frx":18B1
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rExoGeneral.frx":1D26
         DownPicture     =   "rExoGeneral.frx":21EA
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
         Left            =   4740
         Picture         =   "rExoGeneral.frx":26D6
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "EExoGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Exoneraciones en General
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim sMensaje As String
Dim mo_Teclado As New SIGHEntidades.Teclado
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
        oRptClaseCry.IdPuntoCarga = 20
        oRptClaseCry.TipoReporte = Me.Name
        oRptClaseCry.Show vbModal
        Set oRptClaseCry = Nothing
        Me.MousePointer = 1
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    ml_TextoDelFiltro = "FILTROS:   F.Documento: (" & txtFdesde.Text & " " & txtHrInicio.Text & "   al " & txtFhasta.Text & " " & txtHrFin.Text & ")     "
    
    If Me.txtFdesde = SIGHEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha desde del movimiento"
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFdesde, "DD/MM/AAAA") Then
            sMensaje = "La fecha desde del movimiento no tiene el formato correcto"
        End If
    End If
    
    If Me.txtFhasta = SIGHEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha hasta del movimiento"
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFhasta, "DD/MM/AAAA") Then
            sMensaje = "La fecha hasta del movimiento no tiene el formato correcto"
        End If
    End If
    
    If Me.txtHrInicio = SIGHEntidades.HORA_VACIA_HM Then
        sMensaje = "Ingrese la hora desde del movimiento"
    Else
        If Not SIGHEntidades.EsHora(txtHrInicio) Then
            sMensaje = "La hora desde del movimiento no tiene el formato correcto"
        End If
    End If
    
    If Me.txtHrFin = SIGHEntidades.HORA_VACIA_HM Then
        sMensaje = "Ingrese la hora hasta del movimiento"
    Else
        If Not SIGHEntidades.EsHora(txtHrFin) Then
            sMensaje = "La hora hasta del movimiento no tiene el formato correcto"
        End If
    End If
    If CDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) > CDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
       Exit Function
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











Private Sub Form_Load()
    txtFdesde.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual
    txtFhasta.Text = Date
    txtHrInicio.Text = "00:00"
    txtHrFin.Text = "23:59"
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
            txtFdesde = SIGHEntidades.FECHA_VACIA_DMY
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
            txtFhasta = SIGHEntidades.FECHA_VACIA_DMY
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
    Set mo_ReglasFacturacion = Nothing
    Set mo_ReglasComunes = Nothing
    Set mo_Formulario = Nothing
End Sub

Private Sub txtHrFin_LostFocus()
    If txtHrFin <> SIGHEntidades.HORA_VACIA_HM Then
        If Not SIGHEntidades.EsHora(txtHrFin) Then
            MsgBox "La hora ingresada no es válida", vbInformation, Me.Caption
            txtHrFin = SIGHEntidades.HORA_VACIA_HM
        End If
    End If
End Sub

Private Sub txtHrInicio_LostFocus()
    If txtHrInicio <> SIGHEntidades.HORA_VACIA_HM Then
        If Not SIGHEntidades.EsHora(txtHrInicio) Then
            MsgBox "La hora ingresada no es válida", vbInformation, Me.Caption
            txtHrInicio = SIGHEntidades.HORA_VACIA_HM
        End If
    End If
End Sub
