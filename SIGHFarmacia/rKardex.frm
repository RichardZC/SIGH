VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form rKardex 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kardex"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rKardex.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   15435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   8880
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   15663
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Kardex por Item"
      TabPicture(0)   =   "rKardex.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ProgressBar1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "grdKardex"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraDatosHistoria"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Kardex Psicotrópicos"
      TabPicture(1)   =   "rKardex.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "ProgressBar21"
      Tab(1).Control(3)=   "ProgressBar22"
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame2 
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
         Left            =   -74910
         TabIndex        =   31
         Top             =   1635
         Width           =   15090
         Begin VB.CommandButton btnImprimePsicotropicos 
            Caption         =   "Imprime"
            Height          =   700
            Left            =   6135
            Picture         =   "rKardex.frx":0D02
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   210
            Width           =   1365
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Cancelar (ESC)"
            DisabledPicture =   "rKardex.frx":11DB
            DownPicture     =   "rKardex.frx":169F
            Height          =   700
            Left            =   7650
            Picture         =   "rKardex.frx":1B8B
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   225
            Width           =   1365
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   -74910
         TabIndex        =   24
         Top             =   450
         Width           =   15090
         Begin VB.ComboBox cmbTipo 
            Height          =   330
            Left            =   1275
            TabIndex        =   36
            Top             =   570
            Width           =   4980
         End
         Begin MSMask.MaskEdBox txtfinicio11 
            Height          =   315
            Left            =   1275
            TabIndex        =   25
            Top             =   180
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
         Begin MSMask.MaskEdBox txtFechafinal11 
            Height          =   315
            Left            =   4140
            TabIndex        =   26
            Top             =   195
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
         Begin MSMask.MaskEdBox txtHoraInicio11 
            Height          =   315
            Left            =   2655
            TabIndex        =   27
            Top             =   180
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   9
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
         Begin MSMask.MaskEdBox txtHoraFinal11 
            Height          =   315
            Left            =   5490
            TabIndex        =   28
            Top             =   195
            Width           =   750
            _ExtentX        =   1323
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
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   210
            Left            =   135
            TabIndex        =   37
            Top             =   630
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "F.Movimiento"
            Height          =   210
            Left            =   135
            TabIndex        =   30
            Top             =   255
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   210
            Left            =   3930
            TabIndex        =   29
            Top             =   225
            Width           =   120
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
         Height          =   1125
         Left            =   105
         TabIndex        =   4
         Top             =   375
         Width           =   15090
         Begin VB.CommandButton btnBuscarPaciente 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2520
            TabIndex        =   11
            Top             =   660
            Width           =   315
         End
         Begin VB.TextBox txtNproducto 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2850
            MaxLength       =   30
            TabIndex        =   10
            Top             =   660
            Width           =   5805
         End
         Begin VB.TextBox txtCodigo 
            Height          =   315
            Left            =   1350
            MaxLength       =   30
            TabIndex        =   9
            ToolTipText     =   "Ingrese el Código SISMED"
            Top             =   660
            Width           =   1155
         End
         Begin VB.ComboBox cmbAlmacen 
            Height          =   330
            Left            =   1350
            TabIndex        =   8
            Top             =   240
            Width           =   7320
         End
         Begin VB.CheckBox chkExcel 
            Alignment       =   1  'Right Justify
            Caption         =   "En Excel"
            Height          =   315
            Left            =   12345
            Picture         =   "rKardex.frx":2077
            TabIndex        =   7
            Top             =   690
            Width           =   1065
         End
         Begin VB.ComboBox cmbTipoSalida 
            Height          =   330
            ItemData        =   "rKardex.frx":2389
            Left            =   10050
            List            =   "rKardex.frx":238B
            TabIndex        =   6
            Top             =   690
            Width           =   2160
         End
         Begin VB.CommandButton btnBuscar 
            Height          =   315
            Left            =   13695
            Picture         =   "rKardex.frx":238D
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   690
            Width           =   1305
         End
         Begin MSMask.MaskEdBox txtFdesde 
            Height          =   315
            Left            =   10050
            TabIndex        =   12
            Top             =   270
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
            Left            =   12915
            TabIndex        =   13
            Top             =   285
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
            Left            =   11430
            TabIndex        =   14
            Top             =   270
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
            Left            =   14265
            TabIndex        =   15
            Top             =   285
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
            Caption         =   "al"
            Height          =   210
            Left            =   12705
            TabIndex        =   20
            Top             =   315
            Width           =   120
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "F.Movimiento"
            Height          =   210
            Left            =   8910
            TabIndex        =   19
            Top             =   345
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Producto"
            Height          =   210
            Left            =   120
            TabIndex        =   18
            Top             =   675
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
            Height          =   210
            Left            =   120
            TabIndex        =   17
            Top             =   270
            Width           =   690
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Salida"
            Height          =   210
            Left            =   9150
            TabIndex        =   16
            Top             =   750
            Width           =   870
         End
      End
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
         Left            =   135
         TabIndex        =   1
         Top             =   7620
         Width           =   15090
         Begin VB.CommandButton btnCancelar 
            Caption         =   "Cancelar (ESC)"
            DisabledPicture =   "rKardex.frx":4FD6
            DownPicture     =   "rKardex.frx":549A
            Height          =   700
            Left            =   7650
            Picture         =   "rKardex.frx":5986
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   225
            Width           =   1365
         End
         Begin VB.CommandButton btnImprimir 
            Caption         =   "Imprime"
            Height          =   700
            Left            =   6135
            Picture         =   "rKardex.frx":5E72
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   210
            Width           =   1365
         End
      End
      Begin UltraGrid.SSUltraGrid grdKardex 
         Height          =   5730
         Left            =   105
         TabIndex        =   21
         Top             =   1545
         Width           =   15090
         _ExtentX        =   26617
         _ExtentY        =   10107
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   71303188
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   "rKardex.frx":634B
         Caption         =   "Kardex"
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   5115
         TabIndex        =   22
         Top             =   7380
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar21 
         Height          =   225
         Left            =   -74865
         TabIndex        =   34
         Top             =   2955
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar22 
         Height          =   225
         Left            =   -74865
         TabIndex        =   35
         Top             =   3195
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Pulse DOBLE CLIC PARA ver detalle del Documento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   135
         TabIndex        =   23
         Top             =   7365
         Width           =   4680
      End
   End
End
Attribute VB_Name = "rKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 '------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte de Kardex
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_cmbAlmacen As New SIGHEntidades.ListaDespleglable
Dim mo_cmbTipo As New SIGHEntidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
Dim mrs_Tmp As New Recordset
Dim oRsFarm_PsicotropTipos As New Recordset
Dim ms_MensajeError As String
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim ml_TextoDelFiltro As String
Const ml_IdPuntoCarga As Integer = 5
Dim lnIdProducto As Long
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim ml_idUsuario As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property



Private Sub btnImprimePsicotropicos_Click()
    Me.MousePointer = 11
    ReportePsicoTropicos
    Me.MousePointer = 1
End Sub

Private Sub btnImprimir_Click()
    On Error GoTo ErrImp
    If mrs_Tmp.RecordCount > 0 Then
       If Me.chkExcel.Value = 1 Then
           mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, "Kardex", ml_TextoDelFiltro, "", Me.hwnd
       Else
            Dim oRptClaseCry As New rCrystal
            oRptClaseCry.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
            oRptClaseCry.IdAlmacen = Val(mo_cmbAlmacen.BoundText)
            oRptClaseCry.idProducto = lnIdProducto
            oRptClaseCry.FechaInicio = CDate(Format(txtFdesde.Text & " " & txtHrInicio & ":00", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
            oRptClaseCry.FechaFin = CDate(Format(txtFhasta.Text & " " & txtHrFin & ":00", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
            oRptClaseCry.TextoDelFiltro = ml_TextoDelFiltro
            oRptClaseCry.TipoReporte = Me.Name
            oRptClaseCry.idTipoSalidaBienInsumo = cmbTipoSalida.ListIndex
            Set oRptClaseCry.oRsRecord = mrs_Tmp
            oRptClaseCry.Show vbModal
            Set oRptClaseCry = Nothing
      End If
    End If
ErrImp:
End Sub
Private Sub btnBuscar_Click()
  If ValidaDatosObligatorios Then
   Me.MousePointer = 11
   Call ProcesaKardex(Val(mo_cmbAlmacen.BoundText), lnIdProducto, _
                  CDate(Format(txtFhasta.Text & " " & txtHrFin & ":00", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS)), _
                  cmbTipoSalida.ListIndex, _
                  CDate(Format(txtFdesde.Text & " " & txtHrInicio & ":00", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS)))
   Set grdKardex.DataSource = mrs_Tmp
   mo_Apariencia.ConfigurarFilasBiColores Me.grdKardex, SIGHEntidades.GrillaConFilasBicolor
   Me.MousePointer = 1
  End If
End Sub

Sub ProcesaKardex(lnIdAlmacen As Long, ml_idProducto As Long, mda_FechaFin As Date, _
                       ml_idTipoSalidaBienInsumo As Long, mda_FechaInicio As Date)
        Dim oConexion As New Connection
        Dim mrs_Tmp99 As New Recordset
        Dim rsReporte As New Recordset
        Dim mrs_Tmp3 As New Recordset
        Dim oBuscaMovimientos As New farmMovimientoDetalle
        Dim lnTotalRegistros As Long, lnSaldoInicial As Long, lnSaldoFinal As Long, lnSalidas As Long, lcTexto3 As String
        Dim LnDic As Long, lcSerieB As String, lcDocumentoB As String, lnIngresos As Long
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open SIGHEntidades.CadenaConexion
        Set rsReporte = oBuscaMovimientos.FarmDevuelveMovimientosDeProducto(ml_idProducto, mda_FechaFin)
        If ml_idTipoSalidaBienInsumo > 0 Then
           rsReporte.Filter = "IdTipoSalidaBienInsumo=" & ml_idTipoSalidaBienInsumo
        End If
        lnTotalRegistros = rsReporte.RecordCount
        
        If lnTotalRegistros = 0 Then
            MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
        Else
            Me.ProgressBar1.Min = 0: Me.ProgressBar1.Max = lnTotalRegistros: Me.ProgressBar1.Value = 0
            If mrs_Tmp.State = 1 Then
               Set mrs_Tmp = Nothing
            End If
            With mrs_Tmp
                  .Fields.Append "FechaCreacion", adDate, 10, adFldIsNullable
                  .Fields.Append "HoraCreacion", adVarChar, 5, adFldIsNullable
                  .Fields.Append "MovTipo", adVarChar, 1, adFldIsNullable
                  .Fields.Append "MovNumero", adVarChar, 10, adFldIsNullable
                  .Fields.Append "Ingresos", adInteger, 4, adFldIsNullable
                  .Fields.Append "salidas", adInteger, 4, adFldIsNullable
                  .Fields.Append "saldo", adInteger, 4, adFldIsNullable
                  .Fields.Append "Abreviatura", adVarChar, 10, adFldIsNullable
                  .Fields.Append "DocumentoNumero", adVarChar, 20, adFldIsNullable
                  .Fields.Append "Concepto", adVarChar, 100, adFldIsNullable
                  .Fields.Append "fOrigen", adVarChar, 100, adFldIsNullable
                  .Fields.Append "Lote", adVarChar, 20, adFldIsNullable
                  .Fields.Append "FechaVencimiento", adDate, 10, adFldIsNullable
                 ' .Fields.Append "fDestino", adVarChar, 100, adFldIsNullable
                 ' .Fields.Append "Estado", adVarChar, 30, adFldIsNullable
                 ' .Fields.Append "Total", adDouble
                  
                  .LockType = adLockOptimistic
                  .Open
            End With
        
            With mrs_Tmp99
                  .Fields.Append "Ingresos", adInteger, 4, adFldIsNullable
                  .Fields.Append "salidas", adInteger, 4, adFldIsNullable
                  .Fields.Append "saldo", adInteger, 4, adFldIsNullable
                  .Fields.Append "Concepto", adVarChar, 100, adFldIsNullable
                  .LockType = adLockOptimistic
                  .Open
            End With
            
            
            
            
            lnSaldoInicial = 0
            'Saldo Inicial
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF And rsReporte.Fields!fechaCreacion < mda_FechaInicio
               If rsReporte.Fields!MovTipo = "S" Then
                  If rsReporte.Fields!IdAlmacenOrigen = lnIdAlmacen Then
                    lnSaldoInicial = lnSaldoInicial - rsReporte.Fields!Cantidad
                  End If
               Else
                  If rsReporte.Fields!IdAlmacenDestino = lnIdAlmacen Then
                     lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!Cantidad
                  End If
               End If
               DoEvents: Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1: Me.Refresh
               rsReporte.MoveNext
               If rsReporte.EOF Then
                  Exit Do
               End If
            Loop
            lnSaldoFinal = lnSaldoInicial
            mrs_Tmp.AddNew
            mrs_Tmp.Fields!movNumero = "<<Saldo>>"
            mrs_Tmp.Fields!ingresos = lnSaldoInicial
            mrs_Tmp.Fields!saldo = lnSaldoInicial
            mrs_Tmp.Update
            '
            mrs_Tmp99.AddNew
            mrs_Tmp99.Fields!Concepto = "Saldo Inicial"
            mrs_Tmp99.Fields!saldo = lnSaldoInicial
            mrs_Tmp99.Update
            '
            Do While Not rsReporte.EOF
               If rsReporte.Fields!MovTipo = "S" Then
                  If rsReporte.Fields!IdAlmacenOrigen = lnIdAlmacen Then
                    lnSaldoFinal = lnSaldoFinal - rsReporte.Fields!Cantidad
                    lnSalidas = lnSalidas + rsReporte.Fields!Cantidad
                    '
                    mrs_Tmp99.MoveFirst
                    mrs_Tmp99.Find "concepto='" & rsReporte.Fields!Concepto & "'"
                    If mrs_Tmp99.EOF Then
                       mrs_Tmp99.AddNew
                       mrs_Tmp99.Fields!Concepto = rsReporte.Fields!Concepto
                       mrs_Tmp99.Fields!salidas = rsReporte.Fields!Cantidad
                    Else
                       mrs_Tmp99.Fields!salidas = mrs_Tmp99.Fields!salidas + rsReporte.Fields!Cantidad
                    End If
                    mrs_Tmp99.Update
                    
                    'debb-03/11/2015 (inicio)
                    lcTexto3 = Trim(rsReporte.Fields!Concepto)
                    Set mrs_Tmp3 = mo_ReglasFarmacia.farmMovimientoVentasSeleccionarXMovimiento("S", rsReporte!movNumero, oConexion)
                    If mrs_Tmp3.RecordCount > 0 Then
                       If Not IsNull(mrs_Tmp3!idCuentaAtencion) Then
                       LnDic = mrs_Tmp3!idCuentaAtencion
                       mrs_Tmp3.Close
                       Set mrs_Tmp3 = mo_ReglasAdmision.AtencionesFiltraDatosCabecera(LnDic, oConexion)
                       If mrs_Tmp3.RecordCount > 0 Then
                          lcTexto3 = Left(lcTexto3, 18) & " <Actual: " & Trim(mrs_Tmp3!dFuenteFinanciamiento) & ">"
                       End If
                       End If
                    End If
                    mrs_Tmp3.Close
                    'debb-03/11/2015 (fin)
                    
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!fechaCreacion = Format(rsReporte.Fields!fechaCreacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
                    mrs_Tmp.Fields!HoraCreacion = Format(rsReporte.Fields!fechaCreacion, SIGHEntidades.DevuelveHoraSoloFormato_HM)
                    mrs_Tmp.Fields!MovTipo = rsReporte.Fields!MovTipo
                    mrs_Tmp.Fields!movNumero = rsReporte.Fields!movNumero
                    mrs_Tmp.Fields!salidas = rsReporte.Fields!Cantidad
                    mrs_Tmp.Fields!saldo = lnSaldoFinal
                    mrs_Tmp.Fields!Abreviatura = rsReporte.Fields!Abreviatura
                    mrs_Tmp.Fields!DocumentoNumero = rsReporte.Fields!DocumentoNumero
                    mrs_Tmp.Fields!Concepto = Left(lcTexto3, 100)                                     'debb-03/11/2015
                   
                    'debb-17/08/2015 (inicio)
                    If Left(rsReporte!Concepto, 5) = "VENTA" And InStr(rsReporte!DocumentoNumero, "-") > 0 Then
                       lcSerieB = Trim(Left(rsReporte!DocumentoNumero, InStr(rsReporte!DocumentoNumero, "-") - 1))
                       lcDocumentoB = Trim(Mid(rsReporte!DocumentoNumero, InStr(rsReporte!DocumentoNumero, "-") + 1, 100))
                       Set mrs_Tmp3 = mo_ReglasCaja.CajaComprobantesPagoSeleccionarPorNroSerieNroDocumento(lcSerieB, lcDocumentoB)
                       If mrs_Tmp3.RecordCount > 0 Then
                          If Not IsNull(mrs_Tmp3!razonSocial) Then
                             mrs_Tmp.Fields!fOrigen = Left(rsReporte.Fields!fDestino & " " & Trim(mrs_Tmp3!razonSocial), 100)
                          Else
                             mrs_Tmp.Fields!fOrigen = Left(rsReporte.Fields!fDestino, 100)
                          End If
                       Else
                          mrs_Tmp.Fields!fOrigen = Left(rsReporte.Fields!fDestino, 100)
                       End If
                       mrs_Tmp3.Close
                    Else
                       mrs_Tmp.Fields!fOrigen = Left(rsReporte.Fields!fDestino & " " & rsReporte.Fields!Datpaciente, 100)
                    End If
                    'debb-17/08/2015 (fin)
                    mrs_Tmp.Fields!Lote = rsReporte.Fields!Lote
                    mrs_Tmp.Fields!FechaVencimiento = rsReporte.Fields!FechaVencimiento
                    mrs_Tmp.Update
                  End If
               Else
                  If rsReporte.Fields!IdAlmacenDestino = lnIdAlmacen Then
                        lnSaldoFinal = lnSaldoFinal + rsReporte.Fields!Cantidad
                        lnIngresos = lnIngresos + rsReporte.Fields!Cantidad
                        '
                        
                        lcTexto3 = ""
                        Set mrs_Tmp3 = mo_ReglasFarmacia.farmMovimientoNotaIngresoSeleccionarXmovimiento(rsReporte!movNumero, rsReporte!MovTipo, oConexion)
                        If mrs_Tmp3.RecordCount > 0 Then
                           If Not IsNull(mrs_Tmp3!Abreviatura) Then
                              lcTexto3 = Trim(mrs_Tmp3!Abreviatura)
                           End If
                           If Not IsNull(mrs_Tmp3!oRigenNumero) Then
                              lcTexto3 = lcTexto3 & " " & Trim(mrs_Tmp3!oRigenNumero)
                           End If
                           If lcTexto3 <> "" Then
                              lcTexto3 = " (" & lcTexto3 & ")"
                           End If
                        End If
                        '
                        mrs_Tmp99.MoveFirst
                        mrs_Tmp99.Find "concepto='" & rsReporte.Fields!Concepto & "'"
                        If mrs_Tmp99.EOF Then
                           mrs_Tmp99.AddNew
                           mrs_Tmp99.Fields!Concepto = rsReporte.Fields!Concepto
                           mrs_Tmp99.Fields!ingresos = rsReporte.Fields!Cantidad
                        Else
                           mrs_Tmp99.Fields!ingresos = mrs_Tmp99.Fields!ingresos + rsReporte.Fields!Cantidad
                        End If
                        mrs_Tmp99.Update
                        '
                        mrs_Tmp3.Close
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!fechaCreacion = Format(rsReporte.Fields!fechaCreacion, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
                        mrs_Tmp.Fields!HoraCreacion = Format(rsReporte.Fields!fechaCreacion, SIGHEntidades.DevuelveHoraSoloFormato_HM)
                        mrs_Tmp.Fields!MovTipo = rsReporte.Fields!MovTipo
                        mrs_Tmp.Fields!movNumero = rsReporte.Fields!movNumero
                        mrs_Tmp.Fields!ingresos = rsReporte.Fields!Cantidad
                        mrs_Tmp.Fields!saldo = lnSaldoFinal
                        mrs_Tmp.Fields!Abreviatura = rsReporte.Fields!Abreviatura
                        mrs_Tmp.Fields!DocumentoNumero = rsReporte.Fields!DocumentoNumero
                        mrs_Tmp.Fields!Concepto = Trim(rsReporte.Fields!Concepto)
                        mrs_Tmp.Fields!fOrigen = Left(Trim(rsReporte.Fields!fOrigen) & lcTexto3, 100)
                        mrs_Tmp.Fields!Lote = rsReporte.Fields!Lote
                        mrs_Tmp.Fields!FechaVencimiento = rsReporte.Fields!FechaVencimiento
                        
                        mrs_Tmp.Update
                  End If
               End If
              ' DoEvents: Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1: Me.Refresh
               rsReporte.MoveNext
               If rsReporte.EOF Then
                  Exit Do
               End If
            Loop
            
            
            'actualiza tablas: farmSaldos, farmSaldosdetallado   'debb-08/09/2019aa-inicio
            If ml_idTipoSalidaBienInsumo > 0 And mda_FechaFin >= Now And mrs_Tmp.RecordCount > 0 Then
                mrs_Tmp.MoveLast
                Dim lnDiferencia As Long, lbActualizoDetalle As Boolean, lnTotalSaldoDetallado As Long, lcSql111 As String
                If rsReporte.State = 1 Then rsReporte.Close
                lcSql111 = "select * from farmSaldo where idAlmacen=" & lnIdAlmacen & " and idProducto=" & ml_idProducto & " and idTipoSalidaBienInsumo=" & ml_idTipoSalidaBienInsumo
                rsReporte.Open lcSql111, oConexion, adOpenKeyset, adLockOptimistic
                If rsReporte.RecordCount > 0 Then
                    lnDiferencia = mrs_Tmp!saldo - rsReporte!Cantidad
                    If lnDiferencia <> 0 Then
                        rsReporte!Cantidad = mrs_Tmp!saldo
                        rsReporte.Update
                    End If
                End If
                '
                If rsReporte.State = 1 Then rsReporte.Close
                lcSql111 = "select * from farmSaldoDetallado where idAlmacen=" & lnIdAlmacen & " and idProducto=" & ml_idProducto & " and idTipoSalidaBienInsumo=" & ml_idTipoSalidaBienInsumo
                rsReporte.Open lcSql111, oConexion, adOpenKeyset, adLockOptimistic
                If rsReporte.RecordCount > 0 Then
                  lnTotalSaldoDetallado = 0
                  rsReporte.MoveFirst
                  Do While Not rsReporte.EOF
                       lnTotalSaldoDetallado = lnTotalSaldoDetallado + rsReporte!Cantidad
                       rsReporte.MoveNext
                  Loop
                  lnDiferencia = mrs_Tmp!saldo - lnTotalSaldoDetallado
                  If lnDiferencia <> 0 Then
                        lbActualizoDetalle = False
                        rsReporte.MoveFirst
                        Do While Not rsReporte.EOF
                            If rsReporte!FechaVencimiento = mrs_Tmp!FechaVencimiento And Trim(rsReporte!Lote) = Trim(mrs_Tmp!Lote) Then
                                 lbActualizoDetalle = True
                                 rsReporte!Cantidad = rsReporte!Cantidad + lnDiferencia
                                 rsReporte.Update
                                 Exit Do
                            End If
                            rsReporte.MoveNext
                        Loop
                        If lbActualizoDetalle = False Then
                          rsReporte.MoveFirst
                          rsReporte!Cantidad = rsReporte!Cantidad + lnDiferencia
                          rsReporte.Update
                        End If
                   End If
                End If
                '
                mrs_Tmp.MoveFirst
            End If
        End If
        Set oConexion = Nothing
        Set mrs_Tmp99 = Nothing
        Set rsReporte = Nothing
        Set mrs_Tmp3 = Nothing
        Set oBuscaMovimientos = Nothing
End Sub

Function ValidaDatosObligatorios() As Boolean
    ms_MensajeError = ""
    ml_TextoDelFiltro = "FILTROS:   Almacén: (" & Trim(cmbAlmacen.Text) & ")     Producto: (" & Trim(txtCodigo.Text) & " - " & Trim(txtNproducto.Text) & ")     F.Movimiento: (" & txtFdesde.Text & " al " & txtFhasta.Text & ")  " & IIf(Me.cmbTipoSalida.Text <> "", " (Tipo Salida: " & Trim(Me.cmbTipoSalida.Text) & ")", "")
    If mo_cmbAlmacen.BoundText = "" Then
        ms_MensajeError = ms_MensajeError + "Por favor elija el Almacén" + Chr(13)
        cmbAlmacen.SetFocus
    ElseIf txtNproducto.Text = "" Then
        ms_MensajeError = ms_MensajeError + "Por favor elija el Producto" + Chr(13)
        txtCodigo.SetFocus
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



Private Sub btnBuscarPaciente_Click()
    Dim oBusqueda As New ListaProductos
    oBusqueda.MuestraTodosItems = False
    oBusqueda.Show 1
    If oBusqueda.BotonPresionado = sghAceptar Then
        lnIdProducto = oBusqueda.IdRegistroSeleccionado
        txtNproducto.Text = oBusqueda.NombreSeleccionado
        txtCodigo.Text = oBusqueda.CodigoSeleccionado
    End If
    Set oBusqueda = Nothing
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub









Private Sub cmbAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmacen
End Sub

Private Sub cmbAlmacen_LostFocus()
    If Val(mo_cmbAlmacen.BoundText) > 0 Then
       Dim lcIdTipoSuministro As String
       mo_ReglasComunes.LlenaDataComboTipoSalidaBienSegunAlmacen Me.cmbTipoSalida, Val(mo_cmbAlmacen.BoundText), lcIdTipoSuministro
    End If
End Sub





Private Sub Command1_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria

End Sub

Private Sub Form_Initialize()
    Set mo_cmbAlmacen.MiComboBox = cmbAlmacen
    Set mo_cmbTipo.MiComboBox = cmbTipo
End Sub

Sub InicializaFechaHora()
    txtFdesde.Text = Date
    txtFhasta.Text = Date
    txtHrInicio.Text = lcBuscaParametro.SeleccionaFilaParametro(267)
    txtHrFin.Text = lcBuscaParametro.SeleccionaFilaParametro(268)
    
    Me.txtfinicio11.Text = "01/" & Format(Date, "mm/yyyy")
    Me.txtFechafinal11.Text = Date
    Me.txtHoraInicio11.Text = "00:00"
    Me.txtHoraFinal11.Text = "23:59"

End Sub

Private Sub Form_Load()
  
    InicializaFechaHora
    '
    Set oRsFarm_PsicotropTipos = mo_ReglasFarmacia.farm_psicotropTipoSeleccionarTodos
    mo_cmbTipo.BoundColumn = "tipo"
    mo_cmbTipo.ListField = "Descripcion"
    Set mo_cmbTipo.RowSource = oRsFarm_PsicotropTipos
    '
    mo_Formulario.HabilitarDeshabilitar Me.txtNproducto, False
    mo_cmbAlmacen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacen.ListField = "Descripcion"
    Set mo_cmbAlmacen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarTodosMenosExternos
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
       'Case vbKeyF2
       '    btnAceptar_Click
       Case vbKeyF6
           btnBuscar_Click
       End Select
End Sub


Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub grdKardex_DblClick()
    On Error GoTo ErrDbl
    Dim oRsTmp1 As New Recordset
    Set oRsTmp1 = grdKardex.DataSource
    If oRsTmp1!MovTipo = "S" Or oRsTmp1!MovTipo = "E" Then
       If oRsTmp1!MovTipo = "S" Then
            Dim oFarmNotaSalida As New FarmNotaSalida
            oFarmNotaSalida.Opcion = sghConsultar
            oFarmNotaSalida.movNumero = oRsTmp1!movNumero
            oFarmNotaSalida.Show 1
            If InStr(ml_TextoDelFiltro, "Especializado") > 0 Then     'debb1212
               oFarmNotaSalida.lnIdTablaLISTBARITEMS = 1305
            Else
               oFarmNotaSalida.lnIdTablaLISTBARITEMS = 1358
            End If
            Set oFarmNotaSalida = Nothing
       Else
            Dim oFarmNotaIngreso As New FarmNotaIngreso
            oFarmNotaIngreso.Opcion = sghConsultar
            oFarmNotaIngreso.movNumero = oRsTmp1!movNumero
            If InStr(ml_TextoDelFiltro, "Especializado") > 0 Then
               oFarmNotaIngreso.lnIdTablaLISTBARITEMS = 1304              'debb1212
            Else
               oFarmNotaIngreso.lnIdTablaLISTBARITEMS = 1357
            End If
            oFarmNotaIngreso.Show 1
            Set oFarmNotaIngreso = Nothing
       End If
    End If
ErrDbl:
End Sub

Private Sub grdKardex_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    'grdKardex.Bands(0).Columns("movTipo").Hidden = True
    'grdKardex.Bands(0).Columns("fDestino").Hidden = True
    'grdKardex.Bands(0).Columns("Estado").Hidden = True
    'grdKardex.Bands(0).Columns("Total").Hidden = True
    grdKardex.Bands(0).Columns("FechaCreacion").Header.Caption = "Fecha"
    grdKardex.Bands(0).Columns("FechaCreacion").Width = 800
    grdKardex.Bands(0).Columns("FechaCreacion").Activation = ssActivationActivateNoEdit
    grdKardex.Bands(0).Columns("HoraCreacion").Header.Caption = "Hora"
    grdKardex.Bands(0).Columns("HoraCreacion").Width = 500
    grdKardex.Bands(0).Columns("HoraCreacion").Activation = ssActivationActivateNoEdit
    grdKardex.Bands(0).Columns("movTipo").Header.Caption = "Tipo"
    grdKardex.Bands(0).Columns("movTipo").Width = 200
    grdKardex.Bands(0).Columns("movTipo").Activation = ssActivationActivateNoEdit
    
    grdKardex.Bands(0).Columns("MovNumero").Header.Caption = "N° Registro"
    grdKardex.Bands(0).Columns("MovNumero").Width = 1000
    grdKardex.Bands(0).Columns("MovNumero").Activation = ssActivationActivateNoEdit
    grdKardex.Bands(0).Columns("Ingresos").Header.Caption = "Ingresos"
    grdKardex.Bands(0).Columns("Ingresos").Width = 700
    grdKardex.Bands(0).Columns("Ingresos").Activation = ssActivationActivateNoEdit
    grdKardex.Bands(0).Columns("salidas").Header.Caption = "Salidas"
    grdKardex.Bands(0).Columns("salidas").Width = 700
    grdKardex.Bands(0).Columns("salidas").Activation = ssActivationActivateNoEdit
    grdKardex.Bands(0).Columns("saldo").Header.Caption = "Saldo"
    grdKardex.Bands(0).Columns("saldo").Width = 700
    grdKardex.Bands(0).Columns("saldo").Activation = ssActivationActivateNoEdit
    grdKardex.Bands(0).Columns("Abreviatura").Header.Caption = "Doc.Tipo"
    grdKardex.Bands(0).Columns("Abreviatura").Width = 500
    grdKardex.Bands(0).Columns("Abreviatura").Activation = ssActivationActivateNoEdit
    grdKardex.Bands(0).Columns("DocumentoNumero").Header.Caption = "Doc.N°"
    grdKardex.Bands(0).Columns("DocumentoNumero").Width = 1200
    grdKardex.Bands(0).Columns("DocumentoNumero").Activation = ssActivationActivateNoEdit
    grdKardex.Bands(0).Columns("Concepto").Header.Caption = "Concepto"
    grdKardex.Bands(0).Columns("Concepto").Width = 3000
    grdKardex.Bands(0).Columns("Concepto").Activation = ssActivationActivateNoEdit
    grdKardex.Bands(0).Columns("fOrigen").Header.Caption = "Origen/Destino"
    grdKardex.Bands(0).Columns("fOrigen").Width = 3700
    grdKardex.Bands(0).Columns("fOrigen").Activation = ssActivationActivateNoEdit
    grdKardex.Bands(0).Columns("Lote").Header.Caption = "Lote"
    grdKardex.Bands(0).Columns("Lote").Width = 700
    grdKardex.Bands(0).Columns("Lote").Activation = ssActivationActivateNoEdit
    grdKardex.Bands(0).Columns("FechaVencimiento").Header.Caption = "F.Vencimiento"
    grdKardex.Bands(0).Columns("FechaVencimiento").Width = 800
    grdKardex.Bands(0).Columns("FechaVencimiento").Activation = ssActivationActivateNoEdit
    
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigo

End Sub




Private Sub txtCodigo_LostFocus()
    If txtCodigo.Text <> "" Then
        Dim rs As New ADODB.Recordset
        txtCodigo.Text = Trim(txtCodigo.Text)
        Set rs = mo_ReglasFarmacia.FactCatalogoBienesInsumosSeleccionarXDescripYcodigo(txtCodigo.Text, "")
        If rs.RecordCount > 0 Then
           lnIdProducto = rs.Fields("idproducto").Value
           txtNproducto.Text = rs.Fields("NombreProducto").Value
        Else
            txtNproducto.Text = ""
            txtCodigo.Text = ""
            lnIdProducto = 0
        End If
        rs.Close
        Set rs = Nothing
    End If

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
    Set mo_cmbAlmacen = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_ReglasComunes = Nothing
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
Sub ReportePsicoTropicosAnterior()
    Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
    Dim oBuscaMovimientos As New farmMovimientoDetalle
    Dim rs As New ADODB.Recordset
    Dim rsReporte As New Recordset
    Dim mrs_Tmp As New Recordset
    Dim mrs_TmpCab As New Recordset
    Dim lnFor As Integer, ml_idProducto As Long, lcProducto As String, mda_FechaFin As Date, mda_FechaInicio As Date
    Dim lnSaldoInicial As Long, lcFiltro As String, lnTotalRegistros As Long, lcPaciente As String
    Dim ml_TextoDelFiltro As String, lcMedico As String, lcMedicoCMP As String, lcNreceta As String
    Dim lnIdCuenta As Long, lnIdMedico As Long, lnIdPaciente As Long
    ml_TextoDelFiltro = "Desde: " & Me.txtfinicio11.Text & " " & Me.txtHoraInicio11.Text & "      hasta: " & Me.txtFechafinal11.Text & " " & Me.txtHoraFinal11.Text & "  (Movimientos solo en FARMACIAS)"
    mda_FechaFin = CDate(Format(Me.txtFechafinal11.Text & " " & Me.txtHoraFinal11.Text & ":00", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
    mda_FechaInicio = CDate(Format(Me.txtfinicio11.Text & " " & Me.txtHoraInicio11.Text & ":00", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
    With mrs_TmpCab
          .Fields.Append "item", adVarChar, 250
          .LockType = adLockOptimistic
          .Open
    End With
    With mrs_Tmp
          .Fields.Append "diaMesMovimiento", adVarChar, 15
          .Fields.Append "mes", adInteger
          .Fields.Append "dia", adInteger
          .Fields.Append "medico", adVarChar, 150
          .Fields.Append "cmp", adVarChar, 20
          .Fields.Append "Paciente", adVarChar, 150
          .Fields.Append "nReceta", adVarChar, 20
          .Fields.Append "Debe1", adInteger
          .Fields.Append "Haber1", adInteger
          .Fields.Append "Debe2", adInteger
          .Fields.Append "Haber2", adInteger
          .Fields.Append "Debe3", adInteger
          .Fields.Append "Haber3", adInteger
          .Fields.Append "Debe4", adInteger
          .Fields.Append "Haber4", adInteger
          .Fields.Append "Debe5", adInteger
          .Fields.Append "Haber5", adInteger
          .LockType = adLockOptimistic
          .Open
          .AddNew
          .Fields!diaMesMovimiento = "0000"
          .Fields!dia = 0
          .Fields!Mes = 0
          .Fields!medico = " "
          .Fields!cmp = " "
          .Fields!nReceta = " "
          .Update
    End With
    
    For lnFor = 1 To 5
        ml_idProducto = 0
        Select Case lnFor
        Case 1
           BuscarDatosCabecera "03452", ml_IdPuntoCarga, mrs_TmpCab, ml_idProducto, lcProducto
'           Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigodebb("03452", 1, ml_IdPuntoCarga)
'           If rs.RecordCount > 0 Then
'                ml_idProducto = rs.Fields("idproducto").Value
'                lcProducto = rs.Fields("NombreProducto").Value
'                mrs_TmpCab.AddNew
'                mrs_TmpCab!Item = Left(rs!nombreProducto, 240) & " (" & Trim(rs!codigo) & ")"
'                mrs_TmpCab.Update
'           End If
'           rs.Close
        Case 2
           BuscarDatosCabecera "03454", ml_IdPuntoCarga, mrs_TmpCab, ml_idProducto, lcProducto
'           Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigodebb("03454", 1, ml_IdPuntoCarga)
'           If rs.RecordCount > 0 Then
'                ml_idProducto = rs.Fields("idproducto").Value
'                lcProducto = rs.Fields("NombreProducto").Value
'           End If
'           rs.Close
        Case 3
          BuscarDatosCabecera "06188", ml_IdPuntoCarga, mrs_TmpCab, ml_idProducto, lcProducto
'           Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigodebb("06188", 1, ml_IdPuntoCarga)
'           If rs.RecordCount > 0 Then
'                ml_idProducto = rs.Fields("idproducto").Value
'                lcProducto = rs.Fields("NombreProducto").Value
'           End If
'           rs.Close
        Case 4
           BuscarDatosCabecera "00393", ml_IdPuntoCarga, mrs_TmpCab, ml_idProducto, lcProducto
        Case 5
           BuscarDatosCabecera "00670", ml_IdPuntoCarga, mrs_TmpCab, ml_idProducto, lcProducto
        Case 444
            ml_idProducto = 0
            lcProducto = ""
        Case 445
            ml_idProducto = 0
            lcProducto = ""
        End Select
        Set rsReporte = oBuscaMovimientos.FarmDevuelveMovimientosDeProducto(ml_idProducto, mda_FechaFin)
        lnTotalRegistros = rsReporte.RecordCount
        If lnTotalRegistros > 0 Then
            lnSaldoInicial = 0
            'Saldo Inicial
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF And rsReporte.Fields!fechaCreacion < mda_FechaInicio
               If rsReporte.Fields!MovTipo = "S" Then
                  If rsReporte!idTipoLocalesOrig = "F" And rsReporte!idEstadoOrig = 1 Then
                    lnSaldoInicial = lnSaldoInicial - rsReporte.Fields!Cantidad
                  End If
               Else
                  If rsReporte.Fields!idTipoLocalesDest = "F" And rsReporte!idEstadoDest = 1 Then
                     lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!Cantidad
                  End If
               End If
               rsReporte.MoveNext
               If rsReporte.EOF Then
                  Exit Do
               End If
            Loop
            mrs_Tmp.MoveFirst
            Select Case lnFor
            Case 1
                mrs_Tmp!debe1 = lnSaldoInicial
            Case 2
                mrs_Tmp!debe2 = lnSaldoInicial
            Case 3
                mrs_Tmp!debe3 = lnSaldoInicial
            Case 4
                mrs_Tmp!debe4 = lnSaldoInicial
            Case 5
                mrs_Tmp!debe5 = lnSaldoInicial
            End Select
            mrs_Tmp.Update
            Do While Not rsReporte.EOF
               lcFiltro = Right("0" & Trim(Str(Day(rsReporte!fechaCreacion))), 2) & Right("0" & Trim(Str(Month(rsReporte!fechaCreacion))), 2) & rsReporte!MovTipo & rsReporte!movNumero
               If rsReporte.Fields!MovTipo = "S" And rsReporte.Fields!idTipoLocalesOrig = "F" And rsReporte!idEstadoOrig = 1 Then
                    lcMedico = "": lcMedicoCMP = "": lcNreceta = "": lcPaciente = ""
                    Set rs = mo_ReglasFarmacia.farmMovimientoVentasSeleccionarXId(rsReporte!movNumero, rsReporte!MovTipo)
                    If rs.RecordCount > 0 Then
                       lnIdCuenta = IIf(IsNull(rs!idCuentaAtencion), 0, rs!idCuentaAtencion)
                       If lnIdCuenta > 0 Then
                          'toma los datos del Medico y Paciente
                          lnIdMedico = rs!idPrescriptor
                          lnIdPaciente = rs!IdPaciente
                          rs.Close
                          Set rs = mo_ReglasDeProgMedica.MedicosSeleccionarXIdMedico(lnIdMedico)
                          If rs.RecordCount > 0 Then
                             lcMedicoCMP = rs!colegiatura
                             lcMedico = rs!ApellidoPaterno & " " & rs!ApellidoMaterno & " " & rs!Nombres
                          End If
                          rs.Close
                          Set rs = mo_ReglasAdmision.PacientesSeleccionarPorIdentificador(lnIdPaciente)
                          If rs.RecordCount > 0 Then
                             lcPaciente = Trim(rs!ApellidoPaterno) & " " & Trim(rs!ApellidoMaterno) & _
                                          " " & Trim(rs!PrimerNombre) & "(" & _
                                          SIGHEntidades.HCigualDNI_DevuelveHistoriaConCerosIzquierda(rs!NroHistoriaClinica, False) & _
                                          ") "
                                          
                          End If
                       Else
                          'pagante, Preventa
                          rs.Close
                          Set rs = mo_ReglasCaja.CajaComprobantePagoXmovnumero(rsReporte!movNumero, rsReporte!MovTipo)
                          If rs.RecordCount > 0 Then
                             lcMedicoCMP = "Boleta"
                             lcMedico = rs!nroSerie + "-" & rs!nroDocumento & " " + Format(rs!fechaCobranza, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
                             lcPaciente = rs!razonSocial
                          End If
                       End If
                    End If
                    rs.Close
                    If Trim(lcPaciente) = "" Then
                       lcPaciente = rsReporte!Abreviatura & "   Dcto: " & rsReporte!DocumentoNumero & "    Mov: " & rsReporte!MovTipo & "-" & rsReporte!movNumero
                    End If
                    If Trim(lcMedico) = "" Then
                       lcMedico = rsReporte!Abreviatura & "    Dcto: " & rsReporte!DocumentoNumero & "    Mov: " & rsReporte!MovTipo & "-" & rsReporte!movNumero
                    End If
                    
                    mrs_Tmp.MoveFirst
                    mrs_Tmp.Find "diaMesMovimiento='" & lcFiltro & "'"
                    If mrs_Tmp.EOF Then
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!diaMesMovimiento = lcFiltro
                        mrs_Tmp.Fields!dia = Val(Left(lcFiltro, 2))
                        mrs_Tmp.Fields!Mes = Val(Mid(lcFiltro, 3, 2))
                    End If
                    mrs_Tmp!paciente = lcPaciente
                    mrs_Tmp!medico = lcMedico
                    mrs_Tmp!cmp = lcMedicoCMP
                    mrs_Tmp!nReceta = lcNreceta
                    Select Case lnFor
                    Case 1
                        mrs_Tmp!haber1 = mrs_Tmp!haber1 + rsReporte!Cantidad
                    Case 2
                        mrs_Tmp!haber2 = mrs_Tmp!haber2 + rsReporte!Cantidad
                    Case 3
                        mrs_Tmp!haber3 = mrs_Tmp!haber3 + rsReporte!Cantidad
                    Case 4
                        mrs_Tmp!haber4 = mrs_Tmp!haber4 + rsReporte!Cantidad
                    Case 5
                        mrs_Tmp!haber5 = mrs_Tmp!haber5 + rsReporte!Cantidad
                    End Select
                    mrs_Tmp.Update
               ElseIf rsReporte.Fields!MovTipo <> "S" And rsReporte!idTipoLocalesDest = "F" And rsReporte!idEstadoDest = 1 Then
                    mrs_Tmp.MoveFirst
                    mrs_Tmp.Find "diamesMovimiento='" & lcFiltro & "'"
                    If mrs_Tmp.EOF Then
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!diaMesMovimiento = lcFiltro
                        mrs_Tmp.Fields!dia = Val(Left(lcFiltro, 2))
                        mrs_Tmp.Fields!Mes = Val(Mid(lcFiltro, 3, 2))
                        mrs_Tmp!paciente = rsReporte!Abreviatura & "   Dcto: " & rsReporte!DocumentoNumero & "    Mov: " & rsReporte!MovTipo & "-" & rsReporte!movNumero
                    End If
                    mrs_Tmp!medico = " "
                    mrs_Tmp!cmp = " "
                    mrs_Tmp!nReceta = " "
                    Select Case lnFor
                    Case 1
                        mrs_Tmp!debe1 = mrs_Tmp!debe1 + rsReporte!Cantidad
                    Case 2
                        mrs_Tmp!debe2 = mrs_Tmp!debe2 + rsReporte!Cantidad
                    Case 3
                        mrs_Tmp!debe3 = mrs_Tmp!debe3 + rsReporte!Cantidad
                    Case 4
                        mrs_Tmp!debe4 = mrs_Tmp!debe4 + rsReporte!Cantidad
                    Case 5
                        mrs_Tmp!debe5 = mrs_Tmp!debe5 + rsReporte!Cantidad
                    End Select
                    mrs_Tmp.Update
               End If
               rsReporte.MoveNext
            Loop
        End If
        rsReporte.Close
    Next
    If mrs_Tmp.RecordCount = 0 Then
       MsgBox "No hay información con esos datos", vbInformation, ""
    Else
        mrs_Tmp.Sort = "mes,dia"
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
        Dim iFila As Long
        Dim lnTotal As Long
        Dim mo_ReporteUtil As New ReporteUtil
        Dim lnDebe1 As Long, lnDebe2   As Long, lnDebe3  As Long, lnDebe4  As Long, lnDebe5  As Long, lnDebe6 As Long
        Dim lnDebeT1 As Long, lnDebeT2   As Long, lnDebeT3  As Long, lnDebeT4  As Long, lnDebeT5  As Long, lnDebeT6 As Long
        Dim lnHaberT1 As Long, lnHaberT2   As Long, lnHaberT3  As Long, lnHaberT4  As Long, lnHaberT5  As Long, lnHaberT6 As Long
        'Crea nueva hoja
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\Fpsicotropicos.xls")
        oWorkBookPlantilla.Worksheets("Psicotropicos").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        '
        mrs_Tmp.MoveFirst
        oWorkSheet.Cells(1, 3).Value = " " & ml_TextoDelFiltro
        oWorkSheet.Cells(4, 7).Value = mrs_Tmp!debe1
        oWorkSheet.Cells(4, 9).Value = mrs_Tmp!debe2
        oWorkSheet.Cells(4, 11).Value = mrs_Tmp!debe3
        oWorkSheet.Cells(4, 13).Value = mrs_Tmp!debe4
        oWorkSheet.Cells(4, 15).Value = mrs_Tmp!debe5
        iFila = 6
        lnDebe1 = 0: lnDebe2 = 0: lnDebe3 = 0: lnDebe4 = 0: lnDebe5 = 0
        lnDebeT1 = 0: lnDebeT2 = 0: lnDebeT3 = 0: lnDebeT4 = 0: lnDebeT5 = 0
        lnHaberT1 = 0: lnHaberT2 = 0: lnHaberT3 = 0: lnHaberT4 = 0: lnHaberT5 = 0: lnHaberT6 = 0
        mrs_Tmp.MoveNext
        Do While Not mrs_Tmp.EOF
           oWorkSheet.Cells(iFila, 1).Value = mrs_Tmp!Mes
           oWorkSheet.Cells(iFila, 2).Value = mrs_Tmp!dia
           oWorkSheet.Cells(iFila, 3).Value = mrs_Tmp!medico & " (" & Trim(mrs_Tmp!cmp) & ")"
           oWorkSheet.Cells(iFila, 4).Value = mrs_Tmp!cmp
           oWorkSheet.Cells(iFila, 5).Value = mrs_Tmp!paciente
           oWorkSheet.Cells(iFila, 6).Value = mrs_Tmp!nReceta
           oWorkSheet.Cells(iFila, 7).Value = mrs_Tmp!debe1
           oWorkSheet.Cells(iFila, 8).Value = mrs_Tmp!haber1
           oWorkSheet.Cells(iFila, 9).Value = mrs_Tmp!debe2
           oWorkSheet.Cells(iFila, 10).Value = mrs_Tmp!haber2
           oWorkSheet.Cells(iFila, 11).Value = mrs_Tmp!debe3
           oWorkSheet.Cells(iFila, 12).Value = mrs_Tmp!haber3
           oWorkSheet.Cells(iFila, 13).Value = mrs_Tmp!debe4
           oWorkSheet.Cells(iFila, 14).Value = mrs_Tmp!haber4
           oWorkSheet.Cells(iFila, 15).Value = mrs_Tmp!debe5
           oWorkSheet.Cells(iFila, 16).Value = mrs_Tmp!haber5
           iFila = iFila + 1
           lnDebe1 = lnDebe1 + (mrs_Tmp!debe1 - mrs_Tmp!haber1)
           lnDebe2 = lnDebe2 + (mrs_Tmp!debe2 - mrs_Tmp!haber2)
           lnDebe3 = lnDebe3 + (mrs_Tmp!debe3 - mrs_Tmp!haber3)
           lnDebe4 = lnDebe4 + (mrs_Tmp!debe4 - mrs_Tmp!haber4)
           lnDebe5 = lnDebe5 + (mrs_Tmp!debe5 - mrs_Tmp!haber5)
           lnDebeT1 = lnDebeT1 + mrs_Tmp!debe1
           lnDebeT2 = lnDebeT2 + mrs_Tmp!debe2
           lnDebeT3 = lnDebeT3 + mrs_Tmp!debe3
           lnDebeT4 = lnDebeT4 + mrs_Tmp!debe4
           lnDebeT5 = lnDebeT5 + mrs_Tmp!debe5
           lnHaberT1 = lnHaberT1 + mrs_Tmp!haber1
           lnHaberT2 = lnHaberT2 + mrs_Tmp!haber2
           lnHaberT3 = lnHaberT3 + mrs_Tmp!haber3
           lnHaberT4 = lnHaberT4 + mrs_Tmp!haber4
           lnHaberT5 = lnHaberT5 + mrs_Tmp!haber5
           mrs_Tmp.MoveNext
        Loop
        iFila = iFila + 1
        oWorkSheet.Cells(iFila, 5).Value = "SALDO"
        oWorkSheet.Cells(iFila, 8).Value = lnHaberT1
        oWorkSheet.Cells(iFila, 10).Value = lnHaberT2
        oWorkSheet.Cells(iFila, 12).Value = lnHaberT3
        oWorkSheet.Cells(iFila, 14).Value = lnHaberT4
        oWorkSheet.Cells(iFila, 16).Value = lnHaberT5
        oWorkSheet.Cells(iFila, 7).Value = lnDebeT1
        oWorkSheet.Cells(iFila, 9).Value = lnDebeT2
        oWorkSheet.Cells(iFila, 11).Value = lnDebeT3
        oWorkSheet.Cells(iFila, 13).Value = lnDebeT4
        oWorkSheet.Cells(iFila, 15).Value = lnDebeT5
        iFila = iFila + 1
        oWorkSheet.Cells(iFila, 5).Value = "BALANCE " & txtFhasta.Text & " " & txtHrFin
        oWorkSheet.Cells(iFila, 7).Value = lnDebe1
        oWorkSheet.Cells(iFila, 9).Value = lnDebe2
        oWorkSheet.Cells(iFila, 11).Value = lnDebe3
        oWorkSheet.Cells(iFila, 13).Value = lnDebe4
        oWorkSheet.Cells(iFila, 15).Value = lnDebe5
        iFila = iFila + 1
        oWorkSheet.PageSetup.PrintTitleRows = "$1:$5"
        If oWorkSheet.PageSetup.PrintArea <> "" Then
            oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
        End If
        oExcel.Visible = True
        oWorkSheet.PrintPreview
        'oWorkSheet.PrintOut
    End If
    Set rs = Nothing
    Set rsReporte = Nothing
    Set mrs_Tmp = Nothing
    Set oBuscaMovimientos = Nothing
    Set mrs_TmpCab = Nothing
End Sub

Sub BuscarDatosCabecera(lcCodigo As String, ml_IdPuntoCarga As Long, ByRef mrs_TmpCab As Recordset, _
                        ByRef ml_idProducto As Long, ByRef lcProducto As String)
           ml_idProducto = 0
           lcProducto = ""
           Dim rs As New Recordset
           Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigodebb(lcCodigo, 1, ml_IdPuntoCarga)
           If rs.RecordCount > 0 Then
                ml_idProducto = rs.Fields("idproducto").Value
                lcProducto = rs.Fields("NombreProducto").Value
                mrs_TmpCab.AddNew
                mrs_TmpCab!Item = Left(rs!nombreProducto, 240) & " (" & Trim(rs!codigo) & ")"
                mrs_TmpCab.Update
           End If
           rs.Close
           Set rs = Nothing
End Sub

Sub ReportePsicoTropicos()
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
    Dim oBuscaMovimientos As New farmMovimientoDetalle
    Dim rs As New ADODB.Recordset
    Dim rsReporte As New Recordset
    Dim mrs_Tmp As New Recordset
    Dim mrs_TmpCab As New Recordset
    Dim oRsPsicotropicos As New Recordset
    Dim orsTmp1122 As New Recordset
    Dim lnFor As Integer, ml_idProducto As Long, lcProducto As String, mda_FechaFin As Date, mda_FechaInicio As Date
    Dim lnSaldoInicial As Long, lcFiltro As String, lnTotalRegistros As Long, lcPaciente As String
    Dim ml_TextoDelFiltro As String, lcMedico As String, lcMedicoCMP As String, lcNreceta As String
    Dim lnIdCuenta As Long, lnIdMedico As Long, lnIdPaciente As Long, lnNumeroCodigos As Long
    Dim lnColInicio As Long, lnPosColDebe As Long, lnPosColHaber As Long, lcMensaje As String
    Const lxColInicio As Long = 6
    Const lxPosColDebe As Long = -1
    ml_TextoDelFiltro = "Desde: " & Me.txtfinicio11.Text & " " & Me.txtHoraInicio11.Text & "      hasta: " & Me.txtFechafinal11.Text & " " & Me.txtHoraFinal11.Text & "  (Movimientos solo en FARMACIAS)"
    mda_FechaFin = CDate(Format(Me.txtFechafinal11.Text & " " & Me.txtHoraFinal11.Text & ":00", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
    mda_FechaInicio = CDate(Format(Me.txtfinicio11.Text & " " & Me.txtHoraInicio11.Text & ":00", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
    Set orsTmp1122 = mo_ReglasFarmacia.farm_psicotropicosSeleccionarTodos
    If cmbTipo.Text <> "" Then
       oRsFarm_PsicotropTipos.MoveFirst
       Do While Not oRsFarm_PsicotropTipos.EOF
          If oRsFarm_PsicotropTipos!Descripcion = cmbTipo.Text Then
             orsTmp1122.Filter = "tipo='" & oRsFarm_PsicotropTipos!tipo & "'"
             ml_TextoDelFiltro = UCase(Trim(cmbTipo.Text)) & ": " & ml_TextoDelFiltro
             Exit Do
          End If
          oRsFarm_PsicotropTipos.MoveNext
       Loop
    End If
    If orsTmp1122.RecordCount = 0 Then
       MsgBox "Debe llenar tabla FARM_PSICOTROPICOS"
       Exit Sub
    End If
    With oRsPsicotropicos
          .Fields.Append "numero", adInteger
          .Fields.Append "codigo", adVarChar, 7
          .Fields.Append "item", adVarChar, 250, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
    lnFor = 1
    lcMensaje = ""
    orsTmp1122.MoveFirst
    Do While Not orsTmp1122.EOF
        Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigodebb(orsTmp1122!codigo, 1, ml_IdPuntoCarga)
        If rs.RecordCount = 0 Then
           lcMensaje = lcMensaje & "El Código: " & orsTmp1122!codigo & " no tiene PRECIO" & Chr(10) & Chr(13)
        Else
            oRsPsicotropicos.AddNew
            oRsPsicotropicos.Fields!numero = lnFor
            oRsPsicotropicos.Fields!codigo = orsTmp1122!codigo
            'oRsPsicotropicos.Fields!Item = rs.Fields("NombreProducto").Value
            oRsPsicotropicos.Update
            lnFor = lnFor + 1
        End If
        rs.Close
        orsTmp1122.MoveNext
    Loop
    If lcMensaje <> "" Then
       MsgBox lcMensaje, vbInformation, "Tienen que tener PRECIOS todos, sino eliminelo de la tabla FARM_PSICOTROPICOS"
       Exit Sub
    End If
    lnNumeroCodigos = oRsPsicotropicos.RecordCount
    '
    With mrs_Tmp
          .Fields.Append "diaMesMovimiento", adVarChar, 15
          .Fields.Append "mes", adInteger
          .Fields.Append "dia", adInteger
          .Fields.Append "medico", adVarChar, 150
          .Fields.Append "cmp", adVarChar, 20
          .Fields.Append "Paciente", adVarChar, 150
          .Fields.Append "nReceta", adVarChar, 20
          For lnFor = 1 To lnNumeroCodigos
                .Fields.Append "Debe" & Trim(Str(lnFor)), adInteger
                .Fields.Append "Haber" & Trim(Str(lnFor)), adInteger
          Next
          .LockType = adLockOptimistic
          .Open
          .AddNew
          .Fields!diaMesMovimiento = "0000"
          .Fields!dia = 0
          .Fields!Mes = 0
          .Fields!medico = " "
          .Fields!cmp = " "
          .Fields!nReceta = " "
          .Update
    End With
    lnColInicio = lxColInicio
    lnPosColDebe = lxPosColDebe
    Me.ProgressBar21.Max = lnNumeroCodigos + 1
    Me.ProgressBar21.Min = 0
    Me.ProgressBar21.Value = 0
    Me.ProgressBar22.Value = 0
    For lnFor = 1 To lnNumeroCodigos
        DoEvents
        Me.ProgressBar21.Value = Me.ProgressBar21.Value + 1
        Me.Refresh
        
        lnColInicio = lnColInicio + 1
        lnPosColDebe = lnColInicio + (lnFor - 1)
        lnPosColHaber = lnPosColDebe + 1
        oRsPsicotropicos.MoveFirst
        oRsPsicotropicos.Find "numero=" & Trim(Str(lnFor))
        ml_idProducto = 0
        Set rs = mo_ReglasFacturacion.FacturacionBienesPorCodigodebb(oRsPsicotropicos!codigo, 1, ml_IdPuntoCarga)
        If rs.RecordCount > 0 Then
             ml_idProducto = rs.Fields("idproducto").Value
             lcProducto = rs.Fields("NombreProducto").Value
             oRsPsicotropicos!Item = Left(rs!nombreProducto, 240) & " (" & Trim(rs!codigo) & ")"
             oRsPsicotropicos.Update
        Else
             oRsPsicotropicos!Item = "No existe (" & Trim(rs!codigo) & ")"
             oRsPsicotropicos.Update
        End If
        rs.Close
        '
        Set rsReporte = oBuscaMovimientos.FarmDevuelveMovimientosDeProducto(ml_idProducto, mda_FechaFin)
        lnTotalRegistros = rsReporte.RecordCount
        If lnTotalRegistros > 0 Then
            lnSaldoInicial = 0
            'Saldo Inicial
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF And rsReporte.Fields!fechaCreacion < mda_FechaInicio
               If rsReporte.Fields!MovTipo = "S" Then
                  If rsReporte!idTipoLocalesOrig = "F" And rsReporte!idEstadoOrig = 1 Then
                    lnSaldoInicial = lnSaldoInicial - rsReporte.Fields!Cantidad
                  End If
               Else
                  If rsReporte.Fields!idTipoLocalesDest = "F" And rsReporte!idEstadoDest = 1 Then
                     lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!Cantidad
                  End If
               End If
               rsReporte.MoveNext
               If rsReporte.EOF Then
                  Exit Do
               End If
            Loop
            mrs_Tmp.MoveFirst
            
            mrs_Tmp.Fields(lnPosColDebe).Value = lnSaldoInicial
'            Select Case lnFor
'            Case 1
'                mrs_Tmp!debe1 = lnSaldoInicial
'            Case 2
'                mrs_Tmp!debe2 = lnSaldoInicial
'            Case 3
'                mrs_Tmp!debe3 = lnSaldoInicial
'            Case 4
'                mrs_Tmp!debe4 = lnSaldoInicial
'            Case 5
'                mrs_Tmp!debe5 = lnSaldoInicial
'            End Select

            mrs_Tmp.Update
            Do While Not rsReporte.EOF
               lcFiltro = Right("0" & Trim(Str(Day(rsReporte!fechaCreacion))), 2) & Right("0" & Trim(Str(Month(rsReporte!fechaCreacion))), 2) & rsReporte!MovTipo & rsReporte!movNumero
               If rsReporte.Fields!MovTipo = "S" And rsReporte.Fields!idTipoLocalesOrig = "F" And rsReporte!idEstadoOrig = 1 Then
                    lcMedico = "": lcMedicoCMP = "": lcNreceta = "": lcPaciente = ""
                    Set rs = mo_ReglasFarmacia.farmMovimientoVentasSeleccionarXId(rsReporte!movNumero, rsReporte!MovTipo)
                    If rs.RecordCount > 0 Then
                       lnIdCuenta = IIf(IsNull(rs!idCuentaAtencion), 0, rs!idCuentaAtencion)
                       If lnIdCuenta > 0 Then
                          'toma los datos del Medico y Paciente
                          lnIdMedico = IIf(IsNull(rs!idPrescriptor), 0, rs!idPrescriptor)
                          lnIdPaciente = rs!IdPaciente
                          rs.Close
                          Set rs = mo_ReglasDeProgMedica.MedicosSeleccionarXIdMedico(lnIdMedico)
                          If rs.RecordCount > 0 Then
                             lcMedicoCMP = rs!colegiatura
                             lcMedico = rs!ApellidoPaterno & " " & rs!ApellidoMaterno & " " & rs!Nombres
                          End If
                          rs.Close
                          Set rs = mo_ReglasAdmision.PacientesSeleccionarPorIdentificador(lnIdPaciente)
                          If rs.RecordCount > 0 Then
                             lcPaciente = Trim(rs!ApellidoPaterno) & " " & Trim(rs!ApellidoMaterno) & _
                                          " " & Trim(rs!PrimerNombre) & "(" & _
                                          SIGHEntidades.HCigualDNI_DevuelveHistoriaConCerosIzquierda(rs!NroHistoriaClinica, False) & _
                                          ") "
                                          
                          End If
                       Else
                          'pagante, Preventa
                          rs.Close
                          Set rs = mo_ReglasCaja.CajaComprobantePagoXmovnumero(rsReporte!movNumero, rsReporte!MovTipo)
                          If rs.RecordCount > 0 Then
                             lcMedicoCMP = "Boleta"
                             lcMedico = rs!nroSerie + "-" & rs!nroDocumento & " " + Format(rs!fechaCobranza, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
                             lcPaciente = rs!razonSocial
                          End If
                       End If
                    End If
                    rs.Close
                    If Trim(lcPaciente) = "" Then
                       lcPaciente = rsReporte!Abreviatura & "   Dcto: " & rsReporte!DocumentoNumero & "    Mov: " & rsReporte!MovTipo & "-" & rsReporte!movNumero
                    End If
                    If Trim(lcMedico) = "" Then
                       lcMedico = rsReporte!Abreviatura & "    Dcto: " & rsReporte!DocumentoNumero & "    Mov: " & rsReporte!MovTipo & "-" & rsReporte!movNumero
                    End If
                    
                    mrs_Tmp.MoveFirst
                    mrs_Tmp.Find "diaMesMovimiento='" & lcFiltro & "'"
                    If mrs_Tmp.EOF Then
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!diaMesMovimiento = lcFiltro
                        mrs_Tmp.Fields!dia = Val(Left(lcFiltro, 2))
                        mrs_Tmp.Fields!Mes = Val(Mid(lcFiltro, 3, 2))
                    End If
                    mrs_Tmp!paciente = lcPaciente
                    mrs_Tmp!medico = lcMedico
                    mrs_Tmp!cmp = lcMedicoCMP
                    mrs_Tmp!nReceta = lcNreceta
                    
                    mrs_Tmp.Fields(lnPosColHaber).Value = mrs_Tmp.Fields(lnPosColHaber).Value + rsReporte!Cantidad
'                    Select Case lnFor
'                    Case 1
'                        mrs_Tmp!haber1 = mrs_Tmp!haber1 + rsReporte!Cantidad
'                    Case 2
'                        mrs_Tmp!haber2 = mrs_Tmp!haber2 + rsReporte!Cantidad
'                    Case 3
'                        mrs_Tmp!haber3 = mrs_Tmp!haber3 + rsReporte!Cantidad
'                    Case 4
'                        mrs_Tmp!haber4 = mrs_Tmp!haber4 + rsReporte!Cantidad
'                    Case 5
'                        mrs_Tmp!haber5 = mrs_Tmp!haber5 + rsReporte!Cantidad
'                    End Select
                    
                    mrs_Tmp.Update
               ElseIf rsReporte.Fields!MovTipo <> "S" And rsReporte!idTipoLocalesDest = "F" And rsReporte!idEstadoDest = 1 Then
                    mrs_Tmp.MoveFirst
                    mrs_Tmp.Find "diamesMovimiento='" & lcFiltro & "'"
                    If mrs_Tmp.EOF Then
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!diaMesMovimiento = lcFiltro
                        mrs_Tmp.Fields!dia = Val(Left(lcFiltro, 2))
                        mrs_Tmp.Fields!Mes = Val(Mid(lcFiltro, 3, 2))
                        mrs_Tmp!paciente = rsReporte!Abreviatura & "   Dcto: " & rsReporte!DocumentoNumero & "    Mov: " & rsReporte!MovTipo & "-" & rsReporte!movNumero
                    End If
                    mrs_Tmp!medico = " "
                    mrs_Tmp!cmp = " "
                    mrs_Tmp!nReceta = " "
                    
                    mrs_Tmp.Fields(lnPosColDebe).Value = mrs_Tmp.Fields(lnPosColDebe).Value + rsReporte!Cantidad
'                    Select Case lnFor
'                    Case 1
'                        mrs_Tmp!debe1 = mrs_Tmp!debe1 + rsReporte!Cantidad
'                    Case 2
'                        mrs_Tmp!debe2 = mrs_Tmp!debe2 + rsReporte!Cantidad
'                    Case 3
'                        mrs_Tmp!debe3 = mrs_Tmp!debe3 + rsReporte!Cantidad
'                    Case 4
'                        mrs_Tmp!debe4 = mrs_Tmp!debe4 + rsReporte!Cantidad
'                    Case 5
'                        mrs_Tmp!debe5 = mrs_Tmp!debe5 + rsReporte!Cantidad
'                    End Select
                    mrs_Tmp.Update
               End If
               rsReporte.MoveNext
            Loop
        End If
        rsReporte.Close
    Next
    If mrs_Tmp.RecordCount = 0 Then
       MsgBox "No hay información con esos datos", vbInformation, ""
    Else
        mrs_Tmp.Sort = "mes,dia"
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
        Dim iFila As Long
        Dim lnTotal As Long
        Dim mo_ReporteUtil As New ReporteUtil
        Dim lnDebe1 As Long, lnDebe2   As Long, lnDebe3  As Long, lnDebe4  As Long, lnDebe5  As Long, lnDebe6 As Long
        Dim lnDebeT1 As Long, lnDebeT2   As Long, lnDebeT3  As Long, lnDebeT4  As Long, lnDebeT5  As Long, lnDebeT6 As Long
        Dim lnHaberT1 As Long, lnHaberT2   As Long, lnHaberT3  As Long, lnHaberT4  As Long, lnHaberT5  As Long, lnHaberT6 As Long
        'Crea nueva hoja
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\Fpsicotropicos.xls")
        oWorkBookPlantilla.Worksheets("Psicotropicos").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        '
        mrs_Tmp.MoveFirst
        oWorkSheet.Cells(1, 3).Value = " " & ml_TextoDelFiltro
        
        Dim lnColExcel  As Long
        lnColExcel = 0
        lnColInicio = lxColInicio
        lnPosColDebe = lxPosColDebe
        For lnFor = 1 To lnNumeroCodigos
            lnColInicio = lnColInicio + 1
            lnPosColDebe = lnColInicio + (lnFor - 1)
            lnPosColHaber = lnPosColDebe + 1
            oRsPsicotropicos.MoveFirst
            oRsPsicotropicos.Find "numero=" & Trim(Str(lnFor))
            oWorkSheet.Cells(3, 7 + lnColExcel).Value = oRsPsicotropicos!Item
            oWorkSheet.Cells(4, 7 + lnColExcel).Value = mrs_Tmp.Fields(lnPosColDebe).Value
            lnColExcel = lnColExcel + 2
        Next
'        oWorkSheet.Cells(4, 7).Value = mrs_Tmp!debe1
'        oWorkSheet.Cells(4, 9).Value = mrs_Tmp!debe2
'        oWorkSheet.Cells(4, 11).Value = mrs_Tmp!debe3
'        oWorkSheet.Cells(4, 13).Value = mrs_Tmp!debe4
'        oWorkSheet.Cells(4, 15).Value = mrs_Tmp!debe5
        
        
        Me.ProgressBar22.Max = mrs_Tmp.RecordCount + 1
        Me.ProgressBar22.Min = 0
        Me.ProgressBar22.Value = 0
        
        iFila = 6
        lnDebe1 = 0: lnDebe2 = 0: lnDebe3 = 0: lnDebe4 = 0: lnDebe5 = 0
        lnDebeT1 = 0: lnDebeT2 = 0: lnDebeT3 = 0: lnDebeT4 = 0: lnDebeT5 = 0
        lnHaberT1 = 0: lnHaberT2 = 0: lnHaberT3 = 0: lnHaberT4 = 0: lnHaberT5 = 0: lnHaberT6 = 0
        
        Dim lnArrayDebe(200) As Long
        Dim lnArrayDebeT(200) As Long
        Dim lnArrayHaberT(200) As Long
        
        lnColInicio = lxColInicio
        lnPosColDebe = lxPosColDebe
        For lnFor = 1 To lnNumeroCodigos
             lnColInicio = lnColInicio + 1
             lnPosColDebe = lnColInicio + (lnFor - 1)
             lnPosColHaber = lnPosColDebe + 1
             lnArrayDebe(lnFor) = lnArrayDebe(lnFor) + (mrs_Tmp.Fields(lnPosColDebe).Value - mrs_Tmp.Fields(lnPosColHaber).Value)
             lnArrayDebeT(lnFor) = lnArrayDebeT(lnFor) + mrs_Tmp.Fields(lnPosColDebe).Value
        Next
        
        mrs_Tmp.MoveNext
        Do While Not mrs_Tmp.EOF
           
           DoEvents
           Me.ProgressBar22.Value = Me.ProgressBar22.Value + 1
           Me.Refresh
           
           oWorkSheet.Cells(iFila, 1).Value = mrs_Tmp!Mes
           oWorkSheet.Cells(iFila, 2).Value = mrs_Tmp!dia
           oWorkSheet.Cells(iFila, 3).Value = mrs_Tmp!medico & " (" & Trim(mrs_Tmp!cmp) & ")"
           oWorkSheet.Cells(iFila, 4).Value = mrs_Tmp!cmp
           oWorkSheet.Cells(iFila, 5).Value = mrs_Tmp!paciente
           oWorkSheet.Cells(iFila, 6).Value = mrs_Tmp!nReceta
           
           lnColExcel = 0
           lnColInicio = lxColInicio
           lnPosColDebe = lxPosColDebe
           For lnFor = 1 To lnNumeroCodigos
                lnColInicio = lnColInicio + 1
                lnPosColDebe = lnColInicio + (lnFor - 1)
                lnPosColHaber = lnPosColDebe + 1
                If mrs_Tmp.Fields(lnPosColDebe).Value > 0 Then
                    oWorkSheet.Cells(iFila, 7 + lnColExcel).Value = mrs_Tmp.Fields(lnPosColDebe).Value
                End If
                lnColExcel = lnColExcel + 1
                If mrs_Tmp.Fields(lnPosColHaber).Value <> 0 Then
                    oWorkSheet.Cells(iFila, 7 + lnColExcel).Value = mrs_Tmp.Fields(lnPosColHaber).Value
                End If
                lnColExcel = lnColExcel + 1
           Next
'           oWorkSheet.Cells(iFila, 7).Value = mrs_Tmp!debe1
'           oWorkSheet.Cells(iFila, 8).Value = mrs_Tmp!haber1
'           oWorkSheet.Cells(iFila, 9).Value = mrs_Tmp!debe2
'           oWorkSheet.Cells(iFila, 10).Value = mrs_Tmp!haber2
'           oWorkSheet.Cells(iFila, 11).Value = mrs_Tmp!debe3
'           oWorkSheet.Cells(iFila, 12).Value = mrs_Tmp!haber3
'           oWorkSheet.Cells(iFila, 13).Value = mrs_Tmp!debe4
'           oWorkSheet.Cells(iFila, 14).Value = mrs_Tmp!haber4
'           oWorkSheet.Cells(iFila, 15).Value = mrs_Tmp!debe5
'           oWorkSheet.Cells(iFila, 16).Value = mrs_Tmp!haber5
           
           iFila = iFila + 1
           
           lnColInicio = lxColInicio
           lnPosColDebe = lxPosColDebe
           For lnFor = 1 To lnNumeroCodigos
                lnColInicio = lnColInicio + 1
                lnPosColDebe = lnColInicio + (lnFor - 1)
                lnPosColHaber = lnPosColDebe + 1
                lnArrayDebe(lnFor) = lnArrayDebe(lnFor) + (mrs_Tmp.Fields(lnPosColDebe).Value - mrs_Tmp.Fields(lnPosColHaber).Value)
                lnArrayDebeT(lnFor) = lnArrayDebeT(lnFor) + mrs_Tmp.Fields(lnPosColDebe).Value
                lnArrayHaberT(lnFor) = lnArrayHaberT(lnFor) + mrs_Tmp.Fields(lnPosColHaber).Value
           Next
'           lnDebe1 = lnDebe1 + (mrs_Tmp!debe1 - mrs_Tmp!haber1)
'           lnDebe2 = lnDebe2 + (mrs_Tmp!debe2 - mrs_Tmp!haber2)
'           lnDebe3 = lnDebe3 + (mrs_Tmp!debe3 - mrs_Tmp!haber3)
'           lnDebe4 = lnDebe4 + (mrs_Tmp!debe4 - mrs_Tmp!haber4)
'           lnDebe5 = lnDebe5 + (mrs_Tmp!debe5 - mrs_Tmp!haber5)
           
           
'           lnColInicio = lxColInicio
'           lnPosColDebe = lxPosColDebe
'           For lnFor = 1 To lnNumeroCodigos
'                lnColInicio = lnColInicio + 1
'                lnPosColDebe = lnColInicio + (lnFor - 1)
'                lnPosColHaber = lnPosColDebe + 1
'                lnArrayDebeT(lnFor) = lnArrayDebeT(lnFor) + mrs_Tmp.Fields(lnPosColDebe).Value
'           Next
'           lnDebeT1 = lnDebeT1 + mrs_Tmp!debe1
'           lnDebeT2 = lnDebeT2 + mrs_Tmp!debe2
'           lnDebeT3 = lnDebeT3 + mrs_Tmp!debe3
'           lnDebeT4 = lnDebeT4 + mrs_Tmp!debe4
'           lnDebeT5 = lnDebeT5 + mrs_Tmp!debe5
           
           
           
'           lnColInicio = lxColInicio
'           lnPosColDebe = lxPosColDebe
'           For lnFor = 1 To lnNumeroCodigos
'                lnColInicio = lnColInicio + 1
'                lnPosColDebe = lnColInicio + (lnFor - 1)
'                lnPosColHaber = lnPosColDebe + 1
'                lnArrayHaberT(lnFor) = lnArrayHaberT(lnFor) + mrs_Tmp.Fields(lnPosColHaber).Value
'           Next
'           lnHaberT1 = lnHaberT1 + mrs_Tmp!haber1
'           lnHaberT2 = lnHaberT2 + mrs_Tmp!haber2
'           lnHaberT3 = lnHaberT3 + mrs_Tmp!haber3
'           lnHaberT4 = lnHaberT4 + mrs_Tmp!haber4
'           lnHaberT5 = lnHaberT5 + mrs_Tmp!haber5
           
           mrs_Tmp.MoveNext
        Loop
        iFila = iFila + 1
        oWorkSheet.Cells(iFila, 5).Value = "SALDO"
        
        lnColExcel = 0
        For lnFor = 1 To lnNumeroCodigos
             If lnArrayHaberT(lnFor) <> 0 Then
                oWorkSheet.Cells(iFila, 8 + lnColExcel).Value = lnArrayHaberT(lnFor)
             End If
             lnColExcel = lnColExcel + 2
        Next
'        oWorkSheet.Cells(iFila, 8).Value = lnHaberT1
'        oWorkSheet.Cells(iFila, 10).Value = lnHaberT2
'        oWorkSheet.Cells(iFila, 12).Value = lnHaberT3
'        oWorkSheet.Cells(iFila, 14).Value = lnHaberT4
'        oWorkSheet.Cells(iFila, 16).Value = lnHaberT5

        lnColExcel = 0
        For lnFor = 1 To lnNumeroCodigos
             If lnArrayDebeT(lnFor) <> 0 Then
                oWorkSheet.Cells(iFila, 7 + lnColExcel).Value = lnArrayDebeT(lnFor)
             End If
             lnColExcel = lnColExcel + 2
        Next
'        oWorkSheet.Cells(iFila, 7).Value = lnDebeT1
'        oWorkSheet.Cells(iFila, 9).Value = lnDebeT2
'        oWorkSheet.Cells(iFila, 11).Value = lnDebeT3
'        oWorkSheet.Cells(iFila, 13).Value = lnDebeT4
'        oWorkSheet.Cells(iFila, 15).Value = lnDebeT5

        iFila = iFila + 1
        oWorkSheet.Cells(iFila, 5).Value = "BALANCE " & Me.txtFechafinal11.Text & " " & Me.txtHoraFinal11.Text
        
        lnColExcel = 0
        For lnFor = 1 To lnNumeroCodigos
             If lnArrayDebe(lnFor) <> 0 Then
                oWorkSheet.Cells(iFila, 7 + lnColExcel).Value = lnArrayDebe(lnFor)
             End If
             lnColExcel = lnColExcel + 2
        Next
'        oWorkSheet.Cells(iFila, 7).Value = lnDebe1
'        oWorkSheet.Cells(iFila, 9).Value = lnDebe2
'        oWorkSheet.Cells(iFila, 11).Value = lnDebe3
'        oWorkSheet.Cells(iFila, 13).Value = lnDebe4
'        oWorkSheet.Cells(iFila, 15).Value = lnDebe5
        
        iFila = iFila + 1
        oWorkSheet.PageSetup.PrintTitleRows = "$1:$5"
        If oWorkSheet.PageSetup.PrintArea <> "" Then
            oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
        End If
        oExcel.Visible = True
        oWorkSheet.PrintPreview
        'oWorkSheet.PrintOut
        
    End If
    Set rs = Nothing
    Set rsReporte = Nothing
    Set mrs_Tmp = Nothing
    Set oBuscaMovimientos = Nothing
    Set mrs_TmpCab = Nothing



End Sub
