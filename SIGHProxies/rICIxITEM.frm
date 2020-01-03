VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form rICIxITEM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ICI por ITEM"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rICIxITEM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   15
      TabIndex        =   54
      Top             =   6345
      Width           =   9180
      Begin VB.CheckBox chkExportaDBF 
         Caption         =   "Exporta datos a tablas DBF (tformdet.dbf, tformdetl.dbf....) ?"
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
         Height          =   270
         Left            =   210
         TabIndex        =   55
         Top             =   330
         Width           =   7035
      End
   End
   Begin VB.Frame Frame 
      Height          =   4950
      Index           =   0
      Left            =   30
      TabIndex        =   26
      Top             =   1185
      Width           =   9195
      Begin VB.Frame Frame 
         Height          =   615
         Index           =   3
         Left            =   3795
         TabIndex        =   49
         Top             =   4050
         Width           =   5295
         Begin VB.TextBox txtPrecio 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   750
            MaxLength       =   30
            TabIndex        =   52
            Top             =   240
            Width           =   1155
         End
         Begin MSMask.MaskEdBox txtFvencimiento 
            Height          =   315
            Left            =   3870
            TabIndex        =   50
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
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Precio"
            Height          =   210
            Left            =   120
            TabIndex        =   53
            Top             =   270
            Width           =   495
         End
         Begin VB.Label lblFvencimiento 
            AutoSize        =   -1  'True
            Caption         =   "F.Vencimiento"
            Height          =   210
            Left            =   2640
            TabIndex        =   51
            Top             =   270
            Width           =   1170
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Ingresos"
         Height          =   735
         Index           =   2
         Left            =   0
         TabIndex        =   41
         Top             =   600
         Width           =   9135
         Begin VB.TextBox txtTotalIngresos 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   7830
            MaxLength       =   30
            TabIndex        =   44
            Top             =   240
            Width           =   1155
         End
         Begin VB.TextBox txtDevoluciones 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4590
            MaxLength       =   30
            TabIndex        =   2
            Top             =   240
            Width           =   1155
         End
         Begin VB.TextBox txtIngresos 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1470
            MaxLength       =   30
            TabIndex        =   1
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Total Ingresos"
            Height          =   210
            Left            =   6600
            TabIndex        =   45
            Top             =   270
            Width           =   1170
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Devoluciones"
            Height          =   210
            Left            =   3480
            TabIndex        =   43
            Top             =   270
            Width           =   1065
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ingresos"
            Height          =   210
            Left            =   240
            TabIndex        =   42
            Top             =   270
            Width           =   690
         End
      End
      Begin VB.TextBox txtSaldoFinal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         MaxLength       =   30
         TabIndex        =   39
         Top             =   4170
         Width           =   1155
      End
      Begin VB.Frame Frame 
         Caption         =   "Salidas"
         Height          =   2490
         Index           =   1
         Left            =   15
         TabIndex        =   28
         Top             =   1500
         Width           =   9135
         Begin VB.TextBox txtDevMerma 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1425
            MaxLength       =   30
            TabIndex        =   11
            Top             =   1692
            Width           =   1155
         End
         Begin VB.TextBox txtOtrDevol 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1425
            MaxLength       =   30
            TabIndex        =   9
            Top             =   1329
            Width           =   1155
         End
         Begin VB.TextBox txtDevVencim 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4590
            MaxLength       =   30
            TabIndex        =   10
            Top             =   1329
            Width           =   1155
         End
         Begin VB.TextBox txtDefNacional 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4590
            MaxLength       =   30
            TabIndex        =   8
            Top             =   966
            Width           =   1155
         End
         Begin VB.TextBox txtTotalSalidas 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   7860
            MaxLength       =   30
            TabIndex        =   37
            Top             =   2070
            Width           =   1155
         End
         Begin VB.TextBox txtOtrasSal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4590
            MaxLength       =   30
            TabIndex        =   14
            Top             =   2055
            Width           =   1155
         End
         Begin VB.TextBox txtIntSanit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1425
            MaxLength       =   30
            TabIndex        =   13
            Top             =   2055
            Width           =   1155
         End
         Begin VB.TextBox txtExoneracion 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4590
            MaxLength       =   30
            TabIndex        =   12
            Top             =   1692
            Width           =   1155
         End
         Begin VB.TextBox txtCreditoHosp 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1425
            MaxLength       =   30
            TabIndex        =   7
            Top             =   966
            Width           =   1155
         End
         Begin VB.TextBox txtConvenios 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4590
            MaxLength       =   30
            TabIndex        =   6
            Top             =   600
            Width           =   1155
         End
         Begin VB.TextBox txtSoat 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1425
            MaxLength       =   30
            TabIndex        =   5
            Top             =   603
            Width           =   1155
         End
         Begin VB.TextBox txtSis 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4590
            MaxLength       =   30
            TabIndex        =   4
            Top             =   240
            Width           =   1155
         End
         Begin VB.TextBox txtVentas 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1425
            MaxLength       =   30
            TabIndex        =   3
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Dev.Merma"
            Height          =   210
            Left            =   240
            TabIndex        =   59
            Top             =   1722
            Width           =   915
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Otras Devoluc"
            Height          =   210
            Left            =   240
            TabIndex        =   58
            Top             =   1359
            Width           =   1140
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Dev.Vencim"
            Height          =   210
            Left            =   3570
            TabIndex        =   57
            Top             =   1359
            Width           =   975
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Def.Nacional"
            Height          =   210
            Left            =   3585
            TabIndex        =   56
            Top             =   975
            Width           =   1005
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Total Salidas"
            Height          =   210
            Left            =   6750
            TabIndex        =   38
            Top             =   2100
            Width           =   1005
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Otras salidas"
            Height          =   210
            Left            =   3555
            TabIndex        =   36
            Top             =   2085
            Width           =   990
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Interv.Sanitar."
            Height          =   210
            Left            =   240
            TabIndex        =   35
            Top             =   2085
            Width           =   1170
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Exoneración"
            Height          =   210
            Left            =   3555
            TabIndex        =   34
            Top             =   1722
            Width           =   990
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Credito Hosp."
            Height          =   210
            Left            =   240
            TabIndex        =   33
            Top             =   996
            Width           =   1110
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Convenios"
            Height          =   210
            Left            =   3720
            TabIndex        =   32
            Top             =   630
            Width           =   825
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Soat"
            Height          =   210
            Left            =   240
            TabIndex        =   31
            Top             =   633
            Width           =   375
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Sis"
            Height          =   210
            Left            =   4335
            TabIndex        =   30
            Top             =   240
            Width           =   210
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Ventas"
            Height          =   210
            Left            =   240
            TabIndex        =   29
            Top             =   270
            Width           =   570
         End
      End
      Begin VB.TextBox txtSaldoInicial 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1470
         MaxLength       =   30
         TabIndex        =   0
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Final"
         Height          =   210
         Left            =   240
         TabIndex        =   40
         Top             =   4245
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Inicial"
         Height          =   210
         Left            =   240
         TabIndex        =   27
         Top             =   270
         Width           =   930
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
      Left            =   45
      TabIndex        =   22
      Top             =   7020
      Width           =   9180
      Begin VB.CommandButton btnElimina 
         DisabledPicture =   "rICIxITEM.frx":0CCA
         DownPicture     =   "rICIxITEM.frx":1055
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
         Left            =   7320
         Picture         =   "rICIxITEM.frx":13E8
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Elimina el Medicamento/Insumo del ICI"
         Top             =   600
         Width           =   825
      End
      Begin VB.CommandButton btnNuevo 
         DisabledPicture =   "rICIxITEM.frx":1779
         DownPicture     =   "rICIxITEM.frx":1B62
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
         Left            =   7320
         Picture         =   "rICIxITEM.frx":1F6E
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Nuevo Medicamento/Insumo que no està en el ICI"
         Top             =   240
         Width           =   825
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rICIxITEM.frx":237A
         DownPicture     =   "rICIxITEM.frx":283E
         Height          =   700
         Left            =   4740
         Picture         =   "rICIxITEM.frx":2D2A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rICIxITEM.frx":3216
         DownPicture     =   "rICIxITEM.frx":3676
         Height          =   700
         Left            =   3210
         Picture         =   "rICIxITEM.frx":3AEB
         Style           =   1  'Graphical
         TabIndex        =   15
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
      Height          =   1155
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   9195
      Begin VB.TextBox txtAnio 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1470
         MaxLength       =   30
         TabIndex        =   24
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox txtMes 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   25
         Top             =   240
         Width           =   1185
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1470
         MaxLength       =   30
         TabIndex        =   16
         ToolTipText     =   "Ingrese el Código SISMED"
         Top             =   660
         Width           =   1035
      End
      Begin VB.TextBox txtNproducto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2850
         MaxLength       =   30
         TabIndex        =   19
         Top             =   660
         Width           =   6225
      End
      Begin VB.CommandButton btnBuscarPaciente 
         Caption         =   "..."
         Enabled         =   0   'False
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
         TabIndex        =   18
         Top             =   660
         Width           =   315
      End
      Begin VB.Label lblOpcion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "MODIFICAR"
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
         Left            =   8010
         TabIndex        =   46
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año y Mes"
         Height          =   210
         Left            =   240
         TabIndex        =   21
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         Height          =   210
         Left            =   240
         TabIndex        =   20
         Top             =   675
         Width           =   750
      End
   End
End
Attribute VB_Name = "rICIxITEM"
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
Dim lnOpcion As sghOpciones
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim oDofarm_formDet As New Dofarm_formDet
Dim oFarm_FormDet As New farm_formdet
Dim ms_MensajeError As String
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim ml_FarmaciaICI As String
Dim ml_AnioMes As String
Dim ml_Codigo As String
Dim lbYaCargoActivate As Boolean
Dim lcCodigoEje As String, lcTipSum As String
Property Let FarmaciaICI(lValue As String)
   ml_FarmaciaICI = lValue
End Property
Property Let AnioMes(lValue As String)
   ml_AnioMes = lValue
   txtAnio.Text = Left(lValue, 4)
   txtMes.Text = Right(lValue, 2)
End Property
Property Let codigo(lValue As String)
   txtCodigo.Text = lValue
   txtCodigo_LostFocus
End Property


Private Sub btnAceptar_Click()
        On Error GoTo erroAcpt
        If lnOpcion = sghAgregar Then
            If txtCodigo.Text = "" Then
               MsgBox "No hay CODIGO", vbInformation, ""
               Exit Sub
            ElseIf IsDate(Me.txtFvencimiento.Text) = False Then
               MsgBox "La FECHA DE VENCIMIENTO no es válida", vbInformation, ""
               Exit Sub
            ElseIf Val(Me.txtPrecio.Text) <= 0 Then
               MsgBox "El PRECIO no es válido", vbInformation, ""
               Exit Sub
            
            End If
        End If
        Dim lbEmpiezaExportarTablas As Boolean
        Dim oConexion As New Connection
        Dim oDoFarm_FormDetL As New DoFarm_FormDetL
        Dim oFarm_FormDetL As New Farm_FormDetL
        sighentidades.AbreConexionSIGH oConexion
        oConexion.BeginTrans
        Set oFarm_FormDet.Conexion = oConexion
        Set oFarm_FormDetL.Conexion = oConexion
        If lnOpcion = sghAgregar Then
'            oDofarm_formDet.CODIGO_EJE = "..."
'            oDofarm_formDet.CODIGO_PRE = "PARTEDIARIO"
'            oDofarm_formDet.ANNOMES = txtAnio.Text & txtMes.Text
            oDofarm_formDet.CODIGO_MED = txtCodigo.Text
            oDofarm_formDet.credHosp = Val(txtCreditoHosp.Text)
            oDofarm_formDet.distri = 0
            oDofarm_formDet.do_con = 0
            oDofarm_formDet.do_fecExp = 0
            oDofarm_formDet.do_ingre = 0
            oDofarm_formDet.do_otr = 0
            oDofarm_formDet.do_saldo = 0
            oDofarm_formDet.do_stk = 0
            oDofarm_formDet.do_tot = 0
            oDofarm_formDet.dstkCero = 0
            oDofarm_formDet.exo = Val(txtExoneracion.Text)
            oDofarm_formDet.fac_perd = 0
            oDofarm_formDet.FEC_EXP = CDate(Me.txtFvencimiento.Text)
            oDofarm_formDet.fecha = Date
            oDofarm_formDet.indiProc = " "
            oDofarm_formDet.indiSiga = " "
            oDofarm_formDet.ingre = Val(Me.txtIngresos.Text)
            oDofarm_formDet.intersan = Val(txtIntSanit.Text)
            oDofarm_formDet.merma = 0
            oDofarm_formDet.mptoRepo = 0
            oDofarm_formDet.otr_conv = Val(txtConvenios.Text)
            oDofarm_formDet.otras_sal = Val(txtOtrasSal.Text)
            oDofarm_formDet.precio = CCur(txtPrecio.Text)
            oDofarm_formDet.reingre = Val(txtDevoluciones.Text)
            oDofarm_formDet.REQ = Val(Me.txtTotalSalidas.Text)
            oDofarm_formDet.saldo = Val(txtSaldoInicial.Text)
            oDofarm_formDet.sis = Val(txtSis.Text)
            oDofarm_formDet.soat = Val(txtSoat.Text)
            oDofarm_formDet.STOCK_FIN = Val(txtSaldoFinal.Text)
            oDofarm_formDet.stock_fin1 = Val(txtSaldoFinal.Text)
            oDofarm_formDet.total = Val(Me.txtTotalSalidas.Text)
            oDofarm_formDet.transf = 0
            oDofarm_formDet.Usuario = " "
            oDofarm_formDet.vencido = 0
            oDofarm_formDet.VENTA = Val(txtVentas.Text)
            oDofarm_formDet.ventaInst = 0
            
            oDofarm_formDet.DEFNAC = Val(Me.txtDefNacional.Text)
            oDofarm_formDet.DEVOL = Val(Me.txtOtrDevol.Text)
            oDofarm_formDet.DEV_VEN = Val(Me.txtDevVencim.Text)
            oDofarm_formDet.DEV_MERMA = Val(txtDevMerma.Text)
            oDofarm_formDet.SIT = "1"
            
            If oFarm_FormDet.Insertar(oDofarm_formDet) = False Then
                MsgBox "Error: " & oFarm_FormDet.MensajeError: GoTo erroAcpt
            Else
                oDoFarm_FormDetL.CODIGO_EJE = oDofarm_formDet.CODIGO_EJE
                oDoFarm_FormDetL.CODIGO_PRE = oDofarm_formDet.CODIGO_PRE
                oDoFarm_FormDetL.TIPSUM = oDofarm_formDet.TIPSUM
                oDoFarm_FormDetL.ANNOMES = oDofarm_formDet.ANNOMES
                oDoFarm_FormDetL.IdUsuarioAuditoria = sighentidades.Usuario
                oDoFarm_FormDetL.CODIGO_MED = oDofarm_formDet.CODIGO_MED
                oDoFarm_FormDetL.Lote = "LOTE" & oDofarm_formDet.ANNOMES
                oDoFarm_FormDetL.FECHVTO = oDofarm_formDet.FEC_EXP
                oDoFarm_FormDetL.saldo = oDofarm_formDet.STOCK_FIN
                oDoFarm_FormDetL.SIT = oDofarm_formDet.SIT
                If oFarm_FormDetL.Insertar(oDoFarm_FormDetL) = False Then
                   MsgBox "Error: " & oFarm_FormDetL.MensajeError: GoTo erroAcpt
                End If
            
            End If
        ElseIf lnOpcion = sghModificar Then
            oDofarm_formDet.credHosp = Val(txtCreditoHosp.Text)
            oDofarm_formDet.exo = Val(txtExoneracion.Text)
            oDofarm_formDet.ingre = Val(Me.txtIngresos.Text)
            oDofarm_formDet.intersan = Val(txtIntSanit.Text)
            oDofarm_formDet.otr_conv = Val(txtConvenios.Text)
            oDofarm_formDet.otras_sal = Val(txtOtrasSal.Text)
            oDofarm_formDet.reingre = Val(txtDevoluciones.Text)
            oDofarm_formDet.saldo = Val(txtSaldoInicial.Text)
            oDofarm_formDet.sis = Val(txtSis.Text)
            oDofarm_formDet.soat = Val(txtSoat.Text)
            oDofarm_formDet.STOCK_FIN = Val(txtSaldoFinal.Text)
            oDofarm_formDet.stock_fin1 = Val(txtSaldoFinal.Text)
            oDofarm_formDet.VENTA = Val(txtVentas.Text)
            oDofarm_formDet.FEC_EXP = CDate(Me.txtFvencimiento.Text)
            oDofarm_formDet.precio = CCur(txtPrecio.Text)
            oDofarm_formDet.total = Val(Me.txtTotalSalidas.Text)
            oDofarm_formDet.DEFNAC = Val(Me.txtDefNacional.Text)
            oDofarm_formDet.DEVOL = Val(Me.txtOtrDevol.Text)
            oDofarm_formDet.DEV_VEN = Val(Me.txtDevVencim.Text)
            oDofarm_formDet.DEV_MERMA = Val(txtDevMerma.Text)
            oDofarm_formDet.REQ = Val(Me.txtTotalSalidas.Text)
            If oFarm_FormDet.Modificar(oDofarm_formDet) = False Then
               MsgBox "Error: " & oFarm_FormDet.MensajeError: GoTo erroAcpt
            Else
               oDoFarm_FormDetL.CODIGO_EJE = oDofarm_formDet.CODIGO_EJE
               oDoFarm_FormDetL.CODIGO_PRE = oDofarm_formDet.CODIGO_PRE
               oDoFarm_FormDetL.TIPSUM = oDofarm_formDet.TIPSUM
               oDoFarm_FormDetL.ANNOMES = oDofarm_formDet.ANNOMES
               oDoFarm_FormDetL.CODIGO_MED = oDofarm_formDet.CODIGO_MED
               If oFarm_FormDetL.SeleccionarPorId(oDoFarm_FormDetL) Then
                  oDoFarm_FormDetL.saldo = Val(Me.txtSaldoFinal.Text)
                  oDoFarm_FormDetL.FECHVTO = CDate(Me.txtFvencimiento.Text)
                  If oFarm_FormDetL.Modificar(oDoFarm_FormDetL) = False Then
                     MsgBox "Error: " & oFarm_FormDetL.MensajeError: GoTo erroAcpt
                  End If
               End If
            End If
            
        End If
        oConexion.CommitTrans
        lbEmpiezaExportarTablas = True
        '****************************** exporta DBF ******************************
        If chkExportaDBF.Value = 1 Then
           Dim orsFoxCab As New Recordset
           Dim oRsFox As New Recordset
           Dim oRsFox1 As New Recordset
           Dim oRsFox2 As New Recordset
           Dim oRsICI As New Recordset
           Dim oConexionFox As New Connection
           Dim lcAnioMes  As String, lcSql As String, ldFecha As Date
           
           oConexionFox.CommandTimeout = 300
           oConexionFox.Open "DSN=his"
           oConexionFox.CursorLocation = adUseClient
           '
           lcAnioMes = txtAnio.Text & txtMes.Text
           mo_ReglasFarmacia.PreparaTablasDBF oRsFox, oRsFox1, oRsFox2, ml_FarmaciaICI, lcAnioMes, oConexionFox, lcAnioMes, False
           'ICI-Cabecera
           lcSql = "select * from formato where codigo_pre='" & ml_FarmaciaICI & "' and (annomes>='" & lcAnioMes & "' and annomes<='" & lcAnioMes & "')"
           orsFoxCab.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
           If orsFoxCab.RecordCount = 0 Then
              MsgBox "No tiene datos en FORMATO.DBF"
           Else
                ldFecha = orsFoxCab!fecha
                'ICI-Detalle
                Dim ldFecha2 As Date
                ldFecha2 = CDate("01/" & Me.txtMes.Text & "/" & Me.txtAnio.Text)
                Set oRsICI = mo_ReglasFarmacia.Farm_formDetSeleccionarUltimoSaldoPorIdproductoXmes("", ml_FarmaciaICI, ldFecha2, oConexion)
                oRsICI.MoveFirst
                Do While Not oRsICI.EOF
                     oRsFox.AddNew
                     oRsFox.Fields!CODIGO_EJE = oRsICI!CODIGO_EJE
                     oRsFox.Fields!CODIGO_PRE = oRsICI!CODIGO_PRE
                     oRsFox.Fields!TIPSUM = oRsICI!TIPSUM
                     oRsFox.Fields!ANNOMES = oRsICI!ANNOMES
                     oRsFox.Fields!CODIGO_MED = oRsICI!CODIGO_MED
                     oRsFox.Fields!saldo = oRsICI!saldo
                     oRsFox.Fields!precio = oRsICI!precio
                     oRsFox.Fields!ingre = oRsICI!ingre
                     oRsFox.Fields!reingre = oRsICI!reingre
                     oRsFox.Fields!VENTA = oRsICI!VENTA
                     oRsFox.Fields!sis = oRsICI!sis
                     oRsFox.Fields!intersan = oRsICI!intersan
                     oRsFox.Fields!fac_perd = 0                        'falta
                     oRsFox.Fields!DEFNAC = oRsICI!DEFNAC
                     oRsFox.Fields!exo = oRsICI!exo
                     oRsFox.Fields!soat = oRsICI!soat
                     oRsFox.Fields!credHosp = oRsICI!credHosp
                     oRsFox.Fields!otr_conv = oRsICI!otr_conv
                     oRsFox.Fields!DEVOL = oRsICI!DEVOL
                     oRsFox.Fields!vencido = 0
                     oRsFox.Fields!merma = 0
                     oRsFox.Fields!distri = 0
                     oRsFox.Fields!transf = 0
                     oRsFox.Fields!ventaInst = 0
                     oRsFox.Fields!DEV_VEN = oRsICI!DEV_VEN
                     oRsFox.Fields!DEV_MERMA = oRsICI!DEV_MERMA
                     oRsFox.Fields!otras_sal = oRsICI!otras_sal
                     oRsFox.Fields!STOCK_FIN = oRsICI!STOCK_FIN
                     oRsFox.Fields!stock_fin1 = oRsICI!stock_fin1
                     oRsFox.Fields!REQ = oRsICI!total
                     oRsFox.Fields!total = oRsICI!total
                     oRsFox.Fields!FEC_EXP = oRsICI!FEC_EXP
                     oRsFox.Fields!do_saldo = oRsICI!do_saldo
                     oRsFox.Fields!do_ingre = oRsICI!do_ingre
                     oRsFox.Fields!do_con = oRsICI!do_con
                     oRsFox.Fields!do_otr = oRsICI!do_otr
                     oRsFox.Fields!do_tot = oRsICI!do_tot
                     oRsFox.Fields!do_stk = oRsICI!do_stk
                     oRsFox.Fields!fecha = ldFecha
                     oRsFox.Fields!Usuario = " "
                     oRsFox.Fields!indiProc = " "
                     oRsFox.Fields!SIT = "1"
                     oRsFox.Fields!indiSiga = " "
                     oRsFox.Fields!dstkCero = 0
                     oRsFox.Fields!mptoRepo = 0
                     oRsFox.Update
                     oRsICI.MoveNext
                Loop
                oRsICI.Close
                'ICI-Lotes
                Set oRsICI = mo_ReglasFarmacia.farm_formDetLXmes(lcAnioMes, lcCodigoEje, lcTipSum, ml_FarmaciaICI)
                oRsICI.MoveFirst
                Do While Not oRsICI.EOF
                    'FormDetL
                    oRsFox1.AddNew
                    oRsFox1.Fields!CODIGO_EJE = oRsICI!CODIGO_EJE
                    oRsFox1.Fields!CODIGO_PRE = oRsICI!CODIGO_PRE
                    oRsFox1.Fields!TIPSUM = oRsICI!TIPSUM
                    oRsFox1.Fields!ANNOMES = oRsICI!ANNOMES
                    oRsFox1.Fields!CODIGO_MED = oRsICI!CODIGO_MED
                    oRsFox1.Fields!Lote = oRsICI!Lote
                    oRsFox1.Fields!FECHVTO = oRsICI!FECHVTO
                    oRsFox1.Fields!saldo = oRsICI!saldo
                    oRsFox1.Fields!SIT = "1"
                    oRsFox1.Update
                    'FormDetM
                    oRsFox2.AddNew
                    oRsFox2.Fields!CODIGO_EJE = oRsICI!CODIGO_EJE
                    oRsFox2.Fields!CODIGO_PRE = oRsICI!CODIGO_PRE
                    oRsFox2.Fields!TIPSUM = oRsICI!TIPSUM
                    oRsFox2.Fields!ANNOMES = oRsICI!ANNOMES
                    oRsFox2.Fields!CODIGO_MED = oRsICI!CODIGO_MED
                    oRsFox2.Fields!Lote = oRsICI!Lote
                    oRsFox2.Fields!FECHVTO = oRsICI!FECHVTO
                    oRsFox2.Fields!saldo = oRsICI!saldo
                    oRsFox2.Fields!SIT = "1"
                    oRsFox2.Update
                    oRsICI.MoveNext
               Loop
               MsgBox "Se EXPORTO CORRECTAMENTE las tablas DBF", vbInformation, ""
           End If
           oConexionFox.Close
           Set orsFoxCab = Nothing
           Set oRsFox = Nothing
           Set oRsFox1 = Nothing
           Set oRsFox2 = Nothing
           Set oRsICI = Nothing
           Set oConexionFox = Nothing
        End If
        oConexion.Close
        Set oConexion = Nothing
        Set oDoFarm_FormDetL = Nothing
        Set oFarm_FormDetL = Nothing
        '
        btnCancelar_Click
        Exit Sub
erroAcpt:
        If lbEmpiezaExportarTablas = False Then
           oConexion.RollbackTrans
        End If
        oConexion.Close
        oConexionFox.Close
        Set orsFoxCab = Nothing
        Set oRsFox = Nothing
        Set oRsFox1 = Nothing
        Set oRsFox2 = Nothing
        Set oRsICI = Nothing
        Set oConexionFox = Nothing
        Set oConexion = Nothing
        Set oDoFarm_FormDetL = Nothing
        Set oFarm_FormDetL = Nothing
        Exit Sub
        Resume
End Sub

Private Sub btnBuscarPaciente_Click()
    Dim oBusqueda As New ListaProductos
    oBusqueda.MuestraTodosItems = False
    oBusqueda.Show 1
    If oBusqueda.BotonPresionado = sghAceptar Then
        txtNproducto.Text = oBusqueda.NombreSeleccionado
        txtCodigo.Text = oBusqueda.CodigoSeleccionado
        ml_Codigo = txtCodigo.Text
    End If
    Set oBusqueda = Nothing

End Sub

Private Sub btnElimina_Click()
    On Error GoTo ErrDel
    If MsgBox("Està seguro de ELIMINAR DEL ICI el MEDICAMENTO/INSUMO? ", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
       lnOpcion = sghEliminar
       CambiaDeOpcion
       Dim lcClave As String
       lcClave = InputBox("Ingrese la Clave: ", "")
       If Month(Date) = Val(Left(lcClave, 2)) And Day(Date) = Val(Right(lcClave, 2)) Then
            Dim oConexion As New Connection
            Dim oDoFarm_FormDetL As New DoFarm_FormDetL
            Dim oFarm_FormDetL As New Farm_FormDetL
            sighentidades.AbreConexionSIGH oConexion
            oConexion.BeginTrans
            Set oFarm_FormDet.Conexion = oConexion
            Set oFarm_FormDetL.Conexion = oConexion
            oDofarm_formDet.CODIGO_MED = Me.txtCodigo.Text
            oDofarm_formDet.CODIGO_PRE = ml_FarmaciaICI
            oDofarm_formDet.ANNOMES = Me.txtAnio.Text & Me.txtMes.Text
            If oFarm_FormDet.EliminarPorCodigo(oDofarm_formDet) = False Then
               MsgBox "Error: " & oFarm_FormDet.MensajeError: GoTo ErrDel
            End If
            oDoFarm_FormDetL.CODIGO_EJE = oDofarm_formDet.CODIGO_EJE
            oDoFarm_FormDetL.CODIGO_PRE = oDofarm_formDet.CODIGO_PRE
            oDoFarm_FormDetL.TIPSUM = oDofarm_formDet.TIPSUM
            oDoFarm_FormDetL.ANNOMES = oDofarm_formDet.ANNOMES
            oDoFarm_FormDetL.CODIGO_MED = oDofarm_formDet.CODIGO_MED
            oDoFarm_FormDetL.IdUsuarioAuditoria = sighentidades.Usuario
            If oFarm_FormDetL.EliminarXcodigo(oDoFarm_FormDetL) = False Then
               MsgBox "Error: " & oFarm_FormDetL.MensajeError: GoTo ErrDel
            End If
            oConexion.CommitTrans
            oConexion.Close
            Set oConexion = Nothing
            btnCancelar_Click
            
       End If
    End If
    Exit Sub
ErrDel:
    oConexion.RollbackTrans
    oConexion.Close
    Set oConexion = Nothing
    Set oDoFarm_FormDetL = Nothing
    Set oFarm_FormDetL = Nothing
    Exit Sub
    Resume
End Sub

Private Sub btnNuevo_Click()
    If MsgBox("Està seguro de AGREGAR un NUEVO MEDICAMENTO/INSUMO AL ICI ? ", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
        lnOpcion = sghAgregar
        CambiaDeOpcion
        btnBuscarPaciente.Enabled = True
        Frame(3).Visible = True
        Me.txtFvencimiento.Text = "01/01/2022"
        LimpiaDatos
        btnBuscarPaciente_Click
    End If
End Sub

Sub LimpiaDatos()
    txtCodigo.Text = ""
    txtNproducto = ""
    txtSaldoInicial.Text = "0"
    Me.txtIngresos.Text = "0"
    Me.txtDevoluciones.Text = "0"
    Me.txtVentas.Text = "0"
    Me.txtSis.Text = "0"
    Me.txtSoat.Text = "0"
    Me.txtConvenios.Text = "0"
    Me.txtCreditoHosp.Text = "0"
    Me.txtExoneracion.Text = "0"
    Me.txtIntSanit.Text = "0"
    Me.txtOtrasSal.Text = "0"
    Totalizar
    
End Sub

Private Sub Form_Activate()
    If lbYaCargoActivate = False Then
       lbYaCargoActivate = True
       Dim oConexion As New Connection
       oConexion.CommandTimeout = 900
       oConexion.CursorLocation = adUseClient
       oConexion.Open sighentidades.CadenaConexion
       Set oFarm_FormDet.Conexion = oConexion
       oDofarm_formDet.CODIGO_MED = Me.txtCodigo.Text
       oDofarm_formDet.CODIGO_PRE = ml_FarmaciaICI
       oDofarm_formDet.ANNOMES = Me.txtAnio.Text & Me.txtMes.Text
       If oFarm_FormDet.SeleccionarPorCodigo(oDofarm_formDet) Then
            txtSaldoInicial.Text = oDofarm_formDet.saldo
            Me.txtIngresos.Text = oDofarm_formDet.ingre
            Me.txtDevoluciones.Text = oDofarm_formDet.reingre
            Me.txtVentas.Text = oDofarm_formDet.VENTA
            Me.txtSis.Text = oDofarm_formDet.sis
            Me.txtSoat.Text = oDofarm_formDet.soat
            Me.txtConvenios.Text = oDofarm_formDet.otr_conv
            Me.txtCreditoHosp.Text = oDofarm_formDet.credHosp
            Me.txtExoneracion.Text = oDofarm_formDet.exo
            Me.txtIntSanit.Text = oDofarm_formDet.intersan
            Me.txtOtrasSal.Text = oDofarm_formDet.otras_sal
            Me.txtPrecio.Text = oDofarm_formDet.precio
            Me.txtFvencimiento.Text = Format(oDofarm_formDet.FEC_EXP, sighentidades.DevuelveFechaSoloFormato_DMY)
            Me.txtDefNacional.Text = oDofarm_formDet.DEFNAC
            Me.txtOtrDevol.Text = oDofarm_formDet.DEVOL
            Me.txtDevVencim.Text = oDofarm_formDet.DEV_VEN
            Me.txtDevMerma.Text = oDofarm_formDet.DEV_MERMA
            lcCodigoEje = oDofarm_formDet.CODIGO_EJE
            lcTipSum = oDofarm_formDet.TIPSUM
       End If
       Totalizar
       oConexion.Close
       Set oConexion = Nothing
    End If
End Sub


Private Sub txtCodigo_LostFocus()
    If txtCodigo.Text <> "" Then
        Dim rs As New ADODB.Recordset
        txtCodigo.Text = Trim(txtCodigo.Text)
        Set rs = mo_ReglasFarmacia.FactCatalogoBienesInsumosSeleccionarXDescripYcodigo(txtCodigo.Text, "")
        If rs.RecordCount > 0 Then
           txtNproducto.Text = rs.Fields("NombreProducto").Value
        Else
            txtNproducto.Text = ""
            txtCodigo.Text = ""
        End If
        rs.Close
        Set rs = Nothing
    End If

End Sub


Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Private Sub Form_Load()
    mo_Formulario.HabilitarDeshabilitar txtAnio, False
    mo_Formulario.HabilitarDeshabilitar txtMes, False
    mo_Formulario.HabilitarDeshabilitar txtCodigo, False
    mo_Formulario.HabilitarDeshabilitar txtNproducto, False
    mo_Formulario.HabilitarDeshabilitar txtTotalIngresos, False
    mo_Formulario.HabilitarDeshabilitar txtTotalSalidas, False
    mo_Formulario.HabilitarDeshabilitar txtSaldoFinal, False
    lnOpcion = sghModificar
End Sub

Sub CambiaDeOpcion()
    Select Case lnOpcion
    Case sghEliminar
         Me.lblOpcion.Caption = "ELIMINAR"
    Case sghModificar
         Me.lblOpcion.Caption = "MODIFICAR"
    Case sghAgregar
         Me.lblOpcion.Caption = "AGREGAR"
    End Select
End Sub










Private Sub txtConvenios_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtConvenios
End Sub

Private Sub txtConvenios_KeyPress(KeyAscii As Integer)
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If

End Sub



Private Sub txtConvenios_LostFocus()
Totalizar
End Sub

Private Sub txtCreditoHosp_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtCreditoHosp
End Sub

Private Sub txtCreditoHosp_KeyPress(KeyAscii As Integer)
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If

End Sub

Private Sub txtCreditoHosp_LostFocus()
Totalizar
End Sub


Private Sub txtDefNacional_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtDefNacional
End Sub



Private Sub txtDefNacional_KeyPress(KeyAscii As Integer)
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If
End Sub

Private Sub txtDefNacional_LostFocus()
Totalizar
End Sub

Private Sub txtDevMerma_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtDevMerma
End Sub

Private Sub txtDevMerma_KeyPress(KeyAscii As Integer)
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If

End Sub

Private Sub txtDevMerma_LostFocus()
Totalizar
End Sub

Private Sub txtDevoluciones_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtDevoluciones
End Sub

Private Sub txtDevoluciones_KeyPress(KeyAscii As Integer)
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If

End Sub







Private Sub txtDevoluciones_LostFocus()
Totalizar
End Sub



Private Sub txtDevVencim_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtDevVencim

End Sub

Private Sub txtDevVencim_KeyPress(KeyAscii As Integer)
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If

End Sub

Private Sub txtDevVencim_LostFocus()
Totalizar
End Sub

Private Sub txtExoneracion_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtExoneracion

End Sub

Private Sub txtExoneracion_KeyPress(KeyAscii As Integer)
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If

End Sub

Private Sub txtExoneracion_LostFocus()
Totalizar
End Sub

Private Sub txtIngresos_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIngresos

End Sub

Private Sub txtIngresos_KeyPress(KeyAscii As Integer)
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If

End Sub





Private Sub txtIngresos_LostFocus()
Totalizar
End Sub

Private Sub txtIntSanit_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtIntSanit
End Sub

Private Sub txtIntSanit_KeyPress(KeyAscii As Integer)
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If

End Sub



Private Sub txtIntSanit_LostFocus()
Totalizar
End Sub

Private Sub txtOtrasSal_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtOtrasSal
End Sub

Private Sub txtOtrasSal_KeyPress(KeyAscii As Integer)
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If

End Sub

Private Sub txtOtrasSal_LostFocus()
Totalizar
End Sub



Private Sub txtOtrDevol_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtOtrDevol
End Sub

Private Sub txtOtrDevol_KeyPress(KeyAscii As Integer)
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If

End Sub

Private Sub txtOtrDevol_LostFocus()
Totalizar
End Sub

Private Sub txtSaldoInicial_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtSaldoInicial
End Sub

Private Sub txtSaldoInicial_KeyPress(KeyAscii As Integer)
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If

End Sub





Private Sub txtSaldoInicial_LostFocus()
Totalizar
End Sub

Private Sub txtSis_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtSis
End Sub

Private Sub txtSis_KeyPress(KeyAscii As Integer)
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If

End Sub



Private Sub txtSis_LostFocus()
Totalizar
End Sub



Private Sub txtSoat_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtSoat
End Sub

Private Sub txtSoat_KeyPress(KeyAscii As Integer)
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If

End Sub

Private Sub txtSoat_LostFocus()
Totalizar
End Sub

Private Sub txtVentas_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtVentas
End Sub

Private Sub txtVentas_KeyPress(KeyAscii As Integer)
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                KeyAscii = 0
            End If

End Sub






Sub Totalizar()
    txtTotalIngresos.Text = Val(Me.txtIngresos.Text) + Val(Me.txtDevoluciones.Text)
    txtTotalSalidas.Text = Val(txtVentas.Text) + Val(txtSis.Text) + _
                          Val(Me.txtSoat.Text) + Val(Me.txtConvenios.Text) + _
                          Val(Me.txtCreditoHosp.Text) + Val(Me.txtExoneracion.Text) + _
                          Val(Me.txtIntSanit.Text) + Val(Me.txtOtrasSal.Text) + _
                          Val(Me.txtDefNacional.Text) + Val(Me.txtOtrDevol.Text) + _
                          Val(Me.txtDevVencim.Text) + Val(Me.txtDevMerma.Text)
    txtSaldoFinal.Text = Val(Me.txtSaldoInicial.Text) + Val(Me.txtTotalIngresos.Text) - Val(Me.txtTotalSalidas.Text)
End Sub

Private Sub txtVentas_LostFocus()
Totalizar
End Sub
