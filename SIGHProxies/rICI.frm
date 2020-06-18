VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form rICI 
   Caption         =   "Formato ICI"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rICI.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTodasFarmacias 
      Caption         =   "Todos las Farmacias"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1110
      Width           =   1965
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
      Height          =   5835
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   11445
      Begin VB.CheckBox chkNOconsiderarSALDOcero 
         Caption         =   "NO considerar SALDO INICIAL y FINAL =0 y sin MOVIMIENTO"
         Height          =   255
         Left            =   75
         TabIndex        =   32
         Top             =   2745
         Visible         =   0   'False
         Width           =   6075
      End
      Begin VB.Frame Frame4 
         Height          =   870
         Left            =   45
         TabIndex        =   24
         Top             =   135
         Width           =   11370
         Begin VB.CheckBox chkSaldoInicialDelHistorico 
            Alignment       =   1  'Right Justify
            Caption         =   "Tomar SALDO INICIAL del HISTORICO DEL MES ANTERIOR"
            Height          =   255
            Left            =   6045
            TabIndex        =   33
            Top             =   525
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   5205
         End
         Begin VB.CheckBox chkHistoricosIci 
            Caption         =   "ICI ya generados"
            Height          =   255
            Left            =   75
            TabIndex        =   31
            Top             =   540
            Visible         =   0   'False
            Width           =   1785
         End
         Begin MSMask.MaskEdBox txtFdesde 
            Height          =   315
            Left            =   1755
            TabIndex        =   25
            Top             =   150
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
            Left            =   8745
            TabIndex        =   26
            Top             =   120
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
            Left            =   3135
            TabIndex        =   27
            Top             =   150
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtHrFin 
            Height          =   315
            Left            =   10125
            TabIndex        =   28
            Top             =   120
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "hasta"
            Height          =   210
            Left            =   8250
            TabIndex        =   30
            Top             =   150
            Width           =   435
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "F.Movimiento"
            Height          =   210
            Left            =   105
            TabIndex        =   29
            Top             =   210
            Width           =   1080
         End
      End
      Begin VB.TextBox txtCodigoItem 
         Height          =   315
         Left            =   10125
         TabIndex        =   22
         Top             =   1980
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Frame fraICI 
         Caption         =   "Formato ICI (con recálculo)"
         Height          =   1905
         Left            =   105
         TabIndex        =   16
         Top             =   3810
         Visible         =   0   'False
         Width           =   11250
         Begin VB.Label Label8 
            Caption         =   "*  En el Reporte ICI no se toma en cuenta NOTA DE INGRESO por DEVOLUCION DEL PACIENTE "
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1560
            Width           =   9585
         End
         Begin VB.Label Label7 
            Caption         =   "*  En la tabla 'FarmAlmacen.codigoSISMED' debe estar definido las farmacias de acuerdo a la codificación SISMEDV2"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1230
            Width           =   9585
         End
         Begin VB.Label Label6 
            Caption         =   "*  Se exporta a la Version del Sismed:  30 de Setiembre del 2011 "
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   930
            Width           =   8595
         End
         Begin VB.Label Label5 
            Caption         =   "* Al imprimir el formato se llena las tablas:    formato.dbf, formdet.dbf, formDetL.dbf, formDetM.dbf"
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   150
            TabIndex        =   18
            Top             =   600
            Width           =   8595
         End
         Begin VB.Label Label1 
            Caption         =   "* Debe existir el ODBC: HIS (visual foxpro, tabla libre) que apunte a:   c:\archivos....\galenhos\archivos"
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   150
            TabIndex        =   17
            Top             =   270
            Width           =   8595
         End
      End
      Begin VB.CheckBox chkTproducto 
         Caption         =   "Diferenciar movimientos de cada Producto segun Tipo: Ventas, Estratégico"
         Height          =   255
         Left            =   90
         TabIndex        =   15
         Top             =   2370
         Value           =   1  'Checked
         Width           =   6705
      End
      Begin VB.CheckBox chkSinMov 
         Caption         =   "Considera aquellos productos sin Movimientos"
         Enabled         =   0   'False
         Height          =   255
         Left            =   90
         TabIndex        =   13
         Top             =   1560
         Width           =   4305
      End
      Begin VB.CheckBox chkConsideraOSH 
         Caption         =   "Considerar Nota Salida (concepto=distribución,  destino=otros servicios hospital)"
         Height          =   255
         Left            =   90
         TabIndex        =   12
         Top             =   1980
         Value           =   1  'Checked
         Width           =   7065
      End
      Begin VB.Frame Frame1 
         Caption         =   "Reporte"
         Height          =   615
         Left            =   105
         TabIndex        =   8
         Top             =   3075
         Width           =   11220
         Begin Threed.SSOption optParteDiario 
            Height          =   255
            Left            =   150
            TabIndex        =   9
            Top             =   240
            Width           =   1365
            _ExtentX        =   2408
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
            Caption         =   "Parte Diario"
            Value           =   -1
         End
         Begin Threed.SSOption optICI 
            Height          =   255
            Left            =   6210
            TabIndex        =   10
            Top             =   240
            Width           =   2625
            _ExtentX        =   4630
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
            Caption         =   "Formato ICI (con recálculo)"
         End
         Begin Threed.SSOption optParteDiarioR 
            Height          =   255
            Left            =   2520
            TabIndex        =   11
            Top             =   240
            Width           =   2685
            _ExtentX        =   4736
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
            Caption         =   "Parte Diario (con recálculo)"
         End
      End
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "En Excel"
         Height          =   315
         Left            =   10245
         Picture         =   "rICI.frx":0CCA
         TabIndex        =   7
         Top             =   1455
         Width           =   1050
      End
      Begin VB.ComboBox cmbAlmacen 
         Height          =   330
         Left            =   2220
         TabIndex        =   1
         Top             =   1080
         Width           =   9150
      End
      Begin VB.ComboBox cmbOrden 
         Height          =   330
         ItemData        =   "rICI.frx":0FDC
         Left            =   5430
         List            =   "rICI.frx":0FE6
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1515
         Width           =   2745
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   210
         Left            =   9555
         TabIndex        =   23
         Top             =   2040
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Orden"
         Height          =   210
         Left            =   4875
         TabIndex        =   14
         Top             =   1590
         Width           =   510
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
      Left            =   30
      TabIndex        =   3
      Top             =   5895
      Width           =   11445
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rICI.frx":1002
         DownPicture     =   "rICI.frx":1462
         Height          =   700
         Left            =   4298
         Picture         =   "rICI.frx":18D7
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rICI.frx":1D4C
         DownPicture     =   "rICI.frx":2210
         Height          =   700
         Left            =   5828
         Picture         =   "rICI.frx":26FC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "rICI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte del Formato ICI
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_cmbAlmacen As New sighentidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim sMensaje As String
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_TextoDelFiltro As String
Const ml_IdPuntoCarga As Integer = 5
Dim lnIdProducto As Long
Dim mo_Formulario As New sighentidades.Formulario
Dim ml_idUsuario As Long
Dim lcAlmacenesParaICI As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lcCodigoSismed As String
Dim lbEsDonaciones As Boolean
Dim lbSiGrabaHistorico As Boolean
Dim lbVerificaSiRangoEsDeUnMesCompletoICI As Boolean
Dim ldHoy As Date

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Function VerificaSiRangoEsDeUnMesCompleto(mda_FechaInicio As Date, mda_FechaFin As Date, lc_CodigoItem As String) As Boolean
    lbVerificaSiRangoEsDeUnMesCompletoICI = False
    chkSaldoInicialDelHistorico.Value = 0
    chkHistoricosIci.Value = 0
    lbSiGrabaHistorico = False
    If lcCodigoSismed = "" Then
       Exit Function
    End If
    If optICI.Value = False Then
       Exit Function
    End If
    If Val(lc_CodigoItem) > 0 Then
       Exit Function
    End If
'    If Format(mda_FechaInicio, "hh:mm:ss") <> "00:00:00" Then
'       Exit Function
'    End If
    If Day(mda_FechaInicio) <> 1 Then
       Exit Function
    End If
'    If Format(mda_FechaFin, "hh:mm:ss") <> "23:59:59" Then
'       Exit Function
'    End If
    If Day(mda_FechaFin) <> DevuelveUltimoDiaDelMes(Month(mda_FechaFin), Year(mda_FechaFin)) Then
       Exit Function
    End If
    If Month(mda_FechaInicio) <> Month(mda_FechaFin) Then
       Exit Function
    End If
    lbVerificaSiRangoEsDeUnMesCompletoICI = True
    chkHistoricosIci.Value = 1
    chkSaldoInicialDelHistorico.Value = 1
    ChequeaSiHaySaldoFinalMesAnterior
End Function
'debb-02/05/2019
Sub ChequeaSiHaySaldoFinalMesAnterior()
    Me.chkSaldoInicialDelHistorico.Value = 0
    If chkHistoricosIci.Value = 1 And lbVerificaSiRangoEsDeUnMesCompletoICI = True And lcCodigoSismed <> "" Then
        Dim ldFechaHistoricoXmes As Date
        Dim oConexion As New Connection
        Dim oRsTmp As New Recordset
        sighentidades.AbreConexionSIGH oConexion
        ldFechaHistoricoXmes = CDate("01" & Format(Me.txtFdesde.Text, "/mm/yyyy") & " " & lcBuscaParametro.SeleccionaFilaParametro(263) & ":59") - 1
        Set oRsTmp = mo_ReglasFarmacia.Farm_formDetSeleccionarUltimoSaldoPorIdproductoXmes("", lcCodigoSismed, ldFechaHistoricoXmes, oConexion)
        If oRsTmp.RecordCount > 0 Then
           Me.chkSaldoInicialDelHistorico.Value = 1
           
        Else
           'MsgBox "No hay SALDO DEL MES ANTERIOR", vbInformation, ""
        End If
        oRsTmp.Close
        ChequeaSiHayHistoricoICI oConexion
        oConexion.Close
        Set oConexion = Nothing
        Set oRsTmp = Nothing
    End If
End Sub
Sub ChequeaSiHayHistoricoICI(Optional oConexion As Connection)
   
    Dim rsErrores As New Recordset
    Dim ldFechaHistoricoXmes As Date, lcNombre As String
    Dim oConexion1 As New Connection
    ldFechaHistoricoXmes = CDate("01" & Format(Me.txtFdesde.Text, "/mm/yyyy"))
    If Val(Format(ldHoy, "yyyymm")) > Val(Format(ldFechaHistoricoXmes, "yyyymm")) Then
        If oConexion Is Nothing Then
           sighentidades.AbreConexionSIGH oConexion1
           Set rsErrores = mo_ReglasFarmacia.Farm_formDetSeleccionarUltimoSaldoPorIdproductoXmes("", lcCodigoSismed, ldFechaHistoricoXmes, oConexion1)
        Else
           Set rsErrores = mo_ReglasFarmacia.Farm_formDetSeleccionarUltimoSaldoPorIdproductoXmes("", lcCodigoSismed, ldFechaHistoricoXmes, oConexion)
        End If
        If rsErrores.RecordCount = 0 Then
           lbSiGrabaHistorico = True
           chkHistoricosIci.Value = 0
        Else
           lbSiGrabaHistorico = False
           chkHistoricosIci.Value = 1
        End If
    Else
        lbSiGrabaHistorico = False
        chkHistoricosIci.Value = 0
    End If
    If oConexion Is Nothing Then
       oConexion1.Close
       Set oConexion1 = Nothing
    End If
    
End Sub


Private Sub btnAceptar_Click()
If wxFranklin = "*" Then Exit Sub
    
    Dim ldDia1 As Date, ldDia2 As Date
    Dim oRsTmp1 As New Recordset
    ldDia1 = CDate(Format(txtFdesde.Text & " " & txtHrInicio, sighentidades.DevuelveFechaSoloFormato_DMY_HMS))
    ldDia2 = CDate(Format(txtFhasta.Text & " " & txtHrFin, sighentidades.DevuelveFechaSoloFormato_DMY_HMS))
sighentidades.ParaAuditoria = "paso dia2"
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
sighentidades.ParaAuditoria = "paso validaDatos"
        If optICI.Value = True And Me.chkHistoricosIci.Value = 1 Then
            'chequea que esten todos los meses procesesados en tabla resumen:farm_formDet
            If mo_ReglasFarmacia.farm_formDetChequeaICIqueFaltanGrabar(Format(ldDia1, "yyyymm"), _
                                                                       Format(ldDia2, "yyyymm"), _
                                                                       lcCodigoSismed, ldDia2) = True Then
sighentidades.ParaAuditoria = "paso chequeaICIque"
                Me.txtFhasta.Text = Format(ldDia2, sighentidades.DevuelveFechaSoloFormato_DMY)
                CargaSubTitulo
sighentidades.ParaAuditoria = "ml_texto"
                ml_TextoDelFiltro = ml_TextoDelFiltro & " (ICI HISTORICO)"
                Dim oRptClaseCry1 As New rCrystal
                oRptClaseCry1.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
                oRptClaseCry1.FechaInicio = ldDia1
                oRptClaseCry1.FechaFin = ldDia2
                oRptClaseCry1.OrdenadoPor = cmbOrden.ListIndex
                oRptClaseCry1.TextoDelFiltro = ml_TextoDelFiltro
                oRptClaseCry1.ConsiderarSinMovimientos = IIf(chkSinMov.Value = 1, True, False)
                oRptClaseCry1.AlmacenesParaICI = lcAlmacenesParaICI
                oRptClaseCry1.ConsiderarRecalculo = IIf(optParteDiario.Value = True, False, True)
                oRptClaseCry1.ConsideraOSH = IIf(chkConsideraOSH.Value = 1, True, False)
                oRptClaseCry1.ConsiderarPAdesdeServidor = IIf(Me.chkConsideraOSH.Value = 1, True, False)
                oRptClaseCry1.VtaYestrategicoSeparado = IIf(Me.chkTproducto.Value = 1, True, False)
                oRptClaseCry1.CodigoSismed = lcCodigoSismed
                oRptClaseCry1.EsDonaciones = lbEsDonaciones
                oRptClaseCry1.CodigoItem = txtCodigoItem.Text
                oRptClaseCry1.TipoReporte = "IciMensual"
                oRptClaseCry1.NOconsiderarSALDOcero = IIf(Me.chkNOconsiderarSALDOcero.Value = 1, True, False)
                oRptClaseCry1.EsUnIciHistorico = True
                oRptClaseCry1.HoraFin = txtHrFin.Text
sighentidades.ParaAuditoria = "antes xfangoMes"
                Set oRptClaseCry1.oRsRecord = mo_ReglasFarmacia.farm_formDetXrangoMeses(Format(ldDia1, "yyyymm"), _
                                                                                        Format(ldDia2, "yyyymm"), _
                                                                                        lcCodigoSismed)

                oRptClaseCry1.Show vbModal
                'oRptClaseCry1.SeCorrigioDato = False
                Set oRptClaseCry1 = Nothing
sighentidades.ParaAuditoria = "okey"
            End If
        Else
sighentidades.ParaAuditoria = "graba ICI"
            Dim oRptClaseCry As New rCrystal
            'oRptClaseCry.IdAlmacen = Val(mo_cmbAlmacen.BoundText)
            oRptClaseCry.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
            oRptClaseCry.FechaInicio = ldDia1
            oRptClaseCry.FechaFin = ldDia2
            oRptClaseCry.OrdenadoPor = cmbOrden.ListIndex
            oRptClaseCry.TextoDelFiltro = ml_TextoDelFiltro
            oRptClaseCry.ConsiderarSinMovimientos = IIf(chkSinMov.Value = 1, True, False)
            oRptClaseCry.AlmacenesParaICI = lcAlmacenesParaICI
            oRptClaseCry.ConsiderarRecalculo = IIf(optParteDiario.Value = True, False, True)
            oRptClaseCry.TipoReporte = IIf(optICI.Value = True, "rICI", "rPdiario")
            oRptClaseCry.ConsideraOSH = IIf(chkConsideraOSH.Value = 1, True, False)
            oRptClaseCry.ConsiderarPAdesdeServidor = IIf(Me.chkConsideraOSH.Value = 1, True, False)
            oRptClaseCry.VtaYestrategicoSeparado = IIf(Me.chkTproducto.Value = 1, True, False)
            oRptClaseCry.CodigoSismed = lcCodigoSismed
            oRptClaseCry.EsDonaciones = lbEsDonaciones
            oRptClaseCry.CodigoItem = txtCodigoItem.Text
            oRptClaseCry.SeGrabaICImensual = lbSiGrabaHistorico   'debb-10/12/2018
            oRptClaseCry.ConsiderarSaldoInicialDelHistorico = IIf(chkSaldoInicialDelHistorico.Value = 1, True, False)   'debb-10/12/2018
            oRptClaseCry.NOconsiderarSALDOcero = IIf(Me.chkNOconsiderarSALDOcero.Value = 1, True, False)
            'oRptClaseCry.OdbcICI = txtOdbc.Text
            oRptClaseCry.Show vbModal
            Set oRptClaseCry = Nothing
            If lbSiGrabaHistorico = True Then ChequeaSiHayHistoricoICI
        End If
        Me.MousePointer = 1
        
    End If
End Sub

Sub CargaSubTitulo()
    ml_TextoDelFiltro = "FILTROS:   " & IIf(chkTodasFarmacias.Value = 0, "Almacén: (" & Trim(cmbAlmacen.Text) & ")", "") & "      F.Movimiento: (" & txtFdesde.Text & " " & txtHrInicio.Text & "   al " & txtFhasta.Text & " " & txtHrFin.Text & ")     Orden: " & cmbOrden.Text & "     " & IIf(chkSinMov.Value = 1, "(Todos los Productos)", "(Sólo Productos con Movimientos)") & "     " & IIf(optParteDiario.Value = True, "(Sin Recalculos)", "(Con recalculos)") & _
                        IIf(chkConsideraOSH.Value = 1, "", " (sin OSH)") & IIf(Me.chkNOconsiderarSALDOcero.Value = 1, " (" & Me.chkNOconsiderarSALDOcero.Caption & ") ", "")

End Sub


Sub DevuelveCodigoSisMed()
     Dim oRsTmp As New Recordset
     lcCodigoSismed = ""
     If chkTodasFarmacias.Value = 0 Then
           If Val(mo_cmbAlmacen.BoundText) > 0 Then
                Set oRsTmp = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' and idAlmacen=" & mo_cmbAlmacen.BoundText)
                If oRsTmp.RecordCount > 0 Then
                   lcCodigoSismed = oRsTmp.Fields!CodigoSismed
                   lbEsDonaciones = IIf(oRsTmp.Fields!idTipoSuministro = "02", True, False)
                End If
                oRsTmp.Close
           End If
      Else
           lcCodigoSismed = Trim(lcBuscaParametro.SeleccionaFilaParametro(208)) & "F01"
      End If
      Set oRsTmp = Nothing
      If lcCodigoSismed <> "" Then
         VerificaSiRangoEsDeUnMesCompleto CDate(txtFdesde.Text), CDate(txtFhasta.Text), txtCodigoItem.Text
      End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    Dim oRsTmp As New Recordset
    sMensaje = ""
    
    CargaSubTitulo
    'ml_TextoDelFiltro = "FILTROS:   " & IIf(chkTodasFarmacias.Value = 0, "Almacén: (" & Trim(cmbAlmacen.Text) & ")", "") & "      F.Movimiento: (" & txtFdesde.Text & " " & txtHrInicio.Text & "   al " & txtFhasta.Text & " " & txtHrFin.Text & ")     Orden: " & cmbOrden.Text & "     " & IIf(chkSinMov.Value = 1, "(Todos los Productos)", "(Sólo Productos con Movimientos)") & "     " & IIf(optParteDiario.Value = True, "(Sin Recalculos)", "(Con recalculos)") & _
     '                   IIf(chkConsideraOSH.Value = 1, "", " (sin OSH)")
    
    lcCodigoSismed = ""
    If chkTodasFarmacias.Value = 0 Then
        If mo_cmbAlmacen.BoundText = "" Then
            sMensaje = sMensaje + "Por favor elija el Almacén" + Chr(13)
            cmbAlmacen.SetFocus
        Else
           lcAlmacenesParaICI = "/" & mo_cmbAlmacen.BoundText & "/"
           Set oRsTmp = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' and idAlmacen=" & mo_cmbAlmacen.BoundText)
           If oRsTmp.RecordCount > 0 Then
              lcCodigoSismed = oRsTmp.Fields!CodigoSismed
              lbEsDonaciones = IIf(oRsTmp.Fields!idTipoSuministro = "02", True, False)
           End If
           oRsTmp.Close
        End If
    Else
        lbEsDonaciones = False
        lcCodigoSismed = Trim(lcBuscaParametro.SeleccionaFilaParametro(208)) & "F01"
        Set oRsTmp = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idEstado=1 and idTipoLocales='F' and " & _
                                   " idtipoSuministro='01' ")
        oRsTmp.MoveFirst
        lcAlmacenesParaICI = "/"
        Do While Not oRsTmp.EOF
           lcAlmacenesParaICI = lcAlmacenesParaICI & Trim(str(oRsTmp.Fields!IdAlmacen)) & "/"
           oRsTmp.MoveNext
        Loop
        oRsTmp.Close
    End If
    If CDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) > CDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
       Exit Function
    End If
    
    
 
    'debb-16/07/2019
    If chkSaldoInicialDelHistorico.Value = 1 And optICI.Value = True Then
        Dim ldFechaHistoricoXmes As Date
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        ldFechaHistoricoXmes = CDate("01" & Format(Me.txtFdesde.Text, "/mm/yyyy") & " " & lcBuscaParametro.SeleccionaFilaParametro(263) & ":59") - 1
        Set oRsTmp = mo_ReglasFarmacia.Farm_formDetSeleccionarUltimoSaldoPorIdproductoXmes("", lcCodigoSismed, ldFechaHistoricoXmes, oConexion)
        If oRsTmp.RecordCount = 0 Then
           MsgBox "Ha marcado el CHECK: " & chkSaldoInicialDelHistorico.Caption & Chr(13) & _
                  "  pero no hay SALDO FINAL del mes anterior " & Chr(13) & _
                  " tiene que procesar el ICI del MES ANTERIOR", vbInformation, ""
           oConexion.Close
           Set oConexion = Nothing
           Exit Function
        End If
        oConexion.Close
        Set oConexion = Nothing
        '
    End If
    lbSiGrabaHistorico = False
   ' If SIGHEntidades.VerificaClaveMesDia(txtClave2.Text) = True Then
''        If sighentidades.VerificaSiRangoEsDeUnMesCompleto(CDate(Me.txtFdesde.Text), CDate(txtFhasta.Text), txtCodigoItem.Text) = False Then
''           MsgBox "       No podrá GRABAR en HISTORICOS de ICI         " & Chr(13) & _
''                  "  porque el RANGO DE FECHAS no corresponde a un mes " & _
''                  "      o porque quiere ver el ICI de un ITEM         ", vbInformation, ""
''        Else
''            lbSiGrabaHistorico = True
''        End If
   ' End If
    
    '
    Set oRsTmp = Nothing
    
    
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









Private Sub chkHistoricosIci_Click()
'    If Me.chkHistoricosIci.Value = 1 Then
'       chkSaldoInicialDelHistorico.Value = 0
'       txtCodigoItem.Text = ""
'    End If
End Sub

'debb-10/12/2018
Private Sub chkSaldoInicialDelHistorico_Click()
'    If chkSaldoInicialDelHistorico.Value = 1 Then
'       chkHistoricosIci.Value = 0
'    End If
End Sub

Private Sub chkTodasFarmacias_Click()
    If chkTodasFarmacias.Value = 1 Then
       cmbAlmacen.Visible = False
    Else
       cmbAlmacen.Visible = True
    End If
    DevuelveCodigoSisMed
End Sub

Private Sub chkTodasFarmacias_LostFocus()
    VerificaSiRangoEsDeUnMesCompleto CDate(txtFdesde.Text), CDate(txtFhasta.Text), txtCodigoItem.Text
End Sub

Private Sub cmbAlmacen_Click()
    DevuelveCodigoSisMed
End Sub

Private Sub cmbAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmacen

End Sub



Private Sub cmbOrden_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbOrden

End Sub

Private Sub Form_Initialize()
    Set mo_cmbAlmacen.MiComboBox = cmbAlmacen
End Sub

Sub InicializaFechaHora()
    txtFdesde.Text = Date
    txtFhasta.Text = Date
    txtHrInicio.Text = lcBuscaParametro.SeleccionaFilaParametro(267) & ":00"
    txtHrFin.Text = lcBuscaParametro.SeleccionaFilaParametro(268) & ":59"

End Sub

Sub CargaAlmacenes(lbConUNIDOSIS As Boolean)
    mo_cmbAlmacen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacen.ListField = "Descripcion"
    Set mo_cmbAlmacen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' " & _
                                  " and (idtipoSuministro='01' or idtipoSuministro='02') ")
End Sub


Private Sub Form_Load()
    ldHoy = CDate(lcBuscaParametro.RetornaFechaServidorSQL)
    InicializaFechaHora
    CargaAlmacenes True
    cmbOrden.ListIndex = 1
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




Private Sub optICI_Click(Value As Integer)
    If optICI.Value = True Then
       CargaAlmacenes False
       Me.chkTproducto.Visible = False
       fraICI.Visible = True
       chkSinMov.Enabled = True
       txtCodigoItem.Visible = True
       lblCodigo.Visible = True
       chkSinMov.Value = 1
       chkHistoricosIci.Visible = True
       chkNOconsiderarSALDOcero.Visible = True
       chkSaldoInicialDelHistorico.Visible = True    'debb-10/12/2018
       
       VerificaSiRangoEsDeUnMesCompleto CDate(txtFdesde.Text), CDate(txtFhasta.Text), txtCodigoItem.Text
    Else
       Me.chkTproducto.Visible = True
    End If
End Sub



Private Sub optParteDiario_Click(Value As Integer)
   If optParteDiario.Value = True Then
      CargaAlmacenes True
      Me.chkTproducto.Visible = True
      fraICI.Visible = False
      chkSinMov.Enabled = False
      txtCodigoItem.Visible = False
      lblCodigo.Visible = False
      chkHistoricosIci.Visible = False
      chkNOconsiderarSALDOcero.Visible = False
      chkSaldoInicialDelHistorico.Visible = False
   End If
End Sub

Private Sub optParteDiarioR_Click(Value As Integer)
    If optParteDiarioR.Value = True Then
       CargaAlmacenes True
       Me.chkTproducto.Visible = True
       fraICI.Visible = False
       chkSinMov.Enabled = False
       txtCodigoItem.Visible = False
       lblCodigo.Visible = False
       chkHistoricosIci.Visible = False
       chkNOconsiderarSALDOcero.Visible = False
       chkSaldoInicialDelHistorico.Visible = False
    End If
End Sub

Private Sub txtFdesde_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFdesde
End Sub

Private Sub txtFdesde_LostFocus()
    If txtFdesde <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFdesde, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
        Else
            VerificaSiRangoEsDeUnMesCompleto CDate(txtFdesde.Text), CDate(txtFhasta.Text), txtCodigoItem.Text
        End If
    End If

End Sub

Private Sub txtFhasta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFhasta

End Sub

Private Sub txtFhasta_LostFocus()
    If txtFhasta <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFhasta, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
        Else
            VerificaSiRangoEsDeUnMesCompleto CDate(txtFdesde.Text), CDate(txtFhasta.Text), txtCodigoItem.Text
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
