VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form HerrRegeneraSaldos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Regenera Saldos en todos los Almacenes"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "HerrRegeneraSaldos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7800
      Left            =   75
      TabIndex        =   4
      Top             =   0
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   13758
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Regenera Saldos"
      TabPicture(0)   =   "HerrRegeneraSaldos.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Actualiza Precios "
      TabPicture(1)   =   "HerrRegeneraSaldos.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Actualiza Fecha Vencimiento"
      TabPicture(2)   =   "HerrRegeneraSaldos.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label13"
      Tab(2).Control(1)=   "grdFarmSaldoDetallado"
      Tab(2).Control(2)=   "txtCodigoSismed"
      Tab(2).Control(3)=   "cmdActualizaFVencimiento"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "ExportaPrecios"
      TabPicture(3)   =   "HerrRegeneraSaldos.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame8"
      Tab(3).Control(1)=   "Frame7"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Exporta INVENTARIO"
      TabPicture(4)   =   "HerrRegeneraSaldos.frx":0D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraICI"
      Tab(4).Control(1)=   "grdInventarios"
      Tab(4).Control(2)=   "ProgressBar5"
      Tab(4).ControlCount=   3
      Begin VB.Frame fraICI 
         Caption         =   "Formato ICI (con recálculo)"
         Height          =   3510
         Left            =   -74880
         TabIndex        =   55
         Top             =   495
         Width           =   11505
         Begin VB.TextBox txtRutaArchivos 
            Height          =   345
            Left            =   1860
            TabIndex        =   60
            Text            =   "C:\Archivos de programa\Digital Works Corporation\GalenHos\Archivos"
            Top             =   3000
            Width           =   7515
         End
         Begin VB.Label Label27 
            Caption         =   "      ++al ACEPTAR el sistema genera una sola Farmacia en el archivo ZIP, tal cual lo presentan el ICI en el SISMEDV2"
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   75
            TabIndex        =   68
            Top             =   2730
            Width           =   10950
         End
         Begin VB.Label Label26 
            Caption         =   "      ++asignar un solo ""N°Inventario"" a todas, asignar una sola ""F.Cierre"" a todas"
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   75
            TabIndex        =   67
            Top             =   2426
            Width           =   10950
         End
         Begin VB.Label Label25 
            Caption         =   "      ++marcar CHECK a todos los INVENTARIOS del Periodo, asignar un solo ""CodigoSismed"" a todas"
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   75
            TabIndex        =   66
            Top             =   2118
            Width           =   10950
         End
         Begin VB.Label Label24 
            Caption         =   "* Si se tienen varias FARMACIAS en SisGalenPlus pero siempre presentan el ICI  en el SISMEDV2 como si fuese una sola, tendrán que:"
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   75
            TabIndex        =   65
            Top             =   1815
            Width           =   11130
         End
         Begin VB.Label Label20 
            Caption         =   "      ++invCon.dbf -> solo una vez: actualizar datos de su HOSPITAL usando el Visual FOX"
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   75
            TabIndex        =   64
            Top             =   1194
            Width           =   10950
         End
         Begin VB.Label Label19 
            Caption         =   "      ++malma.dbf -> se copia desde 'Sismedv2.exe ->...Exportar Tablas->elegir Malmacen.dbf y se copia como malma.dbf en ARCHIVOS"
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   75
            TabIndex        =   63
            Top             =   885
            Width           =   11280
         End
         Begin VB.Label Label56 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Debe tener ODBC: HIS"
            Height          =   315
            Left            =   60
            TabIndex        =   61
            Top             =   3015
            Width           =   1755
         End
         Begin VB.Label Label23 
            Caption         =   "* Debe existir el ODBC: HIS (visual foxpro, tabla libre) que apunte a:   c:\archivos....\galenhos\archivos"
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   75
            TabIndex        =   58
            Top             =   270
            Width           =   10950
         End
         Begin VB.Label Label22 
            Caption         =   "* Debe tener las tablas:    inv.dbf, invCb.dbf, invDe.dbf, invCon.dbf, mAlma.dbf"
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   75
            TabIndex        =   57
            Top             =   578
            Width           =   10950
         End
         Begin VB.Label Label21 
            Caption         =   "* Se exporta a la Version del Sismed:  30 de Setiembre del 2011 "
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   75
            TabIndex        =   56
            Top             =   1502
            Width           =   10950
         End
      End
      Begin VB.Frame Frame7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   -74880
         TabIndex        =   47
         Top             =   1110
         Width           =   11115
         Begin VB.TextBox Año 
            Height          =   315
            Left            =   2610
            MaxLength       =   4
            TabIndex        =   49
            Top             =   600
            Width           =   915
         End
         Begin VB.TextBox Mes 
            Height          =   315
            Left            =   1380
            MaxLength       =   2
            TabIndex        =   48
            Top             =   600
            Width           =   435
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Año"
            Height          =   210
            Left            =   2160
            TabIndex        =   51
            Top             =   630
            Width           =   330
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Mes"
            Height          =   210
            Left            =   1020
            TabIndex        =   50
            Top             =   630
            Width           =   315
         End
      End
      Begin VB.Frame Frame8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   -74895
         TabIndex        =   45
         Top             =   2445
         Width           =   11115
         Begin MSComctlLib.ProgressBar ProgressBarExportaPreciosSismed 
            Height          =   345
            Left            =   75
            TabIndex        =   46
            Top             =   1335
            Width           =   10875
            _ExtentX        =   19182
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Crea la carpeta 'c:\ExportaPreciosSismedTem', si ya existe ELIMINARLA antes de procesar"
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   150
            TabIndex        =   54
            Top             =   930
            Width           =   7335
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "(Debe existir el ODBC: HIS que apunte a c:\archivos..\digit...\archivos)"
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   105
            TabIndex        =   53
            Top             =   615
            Width           =   5775
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "(Debe existir el archivo: expPrec.dbf en c:\archivos..\digital...\archivos )"
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   120
            TabIndex        =   52
            Top             =   285
            Width           =   5865
         End
      End
      Begin VB.CommandButton cmdActualizaFVencimiento 
         Caption         =   "Actualiza columna:  ""Nueva F.Venc.""   de la lista en las tablas de Farmacia"
         Height          =   705
         Left            =   -74940
         TabIndex        =   38
         Top             =   6720
         Width           =   11415
      End
      Begin VB.TextBox txtCodigoSismed 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -72810
         TabIndex        =   37
         Top             =   6090
         Width           =   1425
      End
      Begin VB.Frame Frame6 
         Height          =   6975
         Left            =   -74850
         TabIndex        =   22
         Top             =   720
         Width           =   11385
         Begin VB.CheckBox chkSoloExiste1 
            Caption         =   "En el archivo con precios (ZIP) sólo existe 1 archivo (Pr???.dbf)"
            Height          =   285
            Left            =   180
            TabIndex        =   42
            Top             =   6000
            Width           =   7335
         End
         Begin VB.CheckBox chkSoloAgregaItems 
            Caption         =   "Solo agrega nuevos medicamentos/insumos (no actualiza precios- 'con check') (actualiza precios - 'sin check')"
            Height          =   285
            Left            =   180
            TabIndex        =   41
            Top             =   5640
            Width           =   10605
         End
         Begin VB.CommandButton cmdDescomprimir 
            Caption         =   "1) Pulse clic para descomprimir archivo ZIP"
            Height          =   1095
            Left            =   150
            TabIndex        =   32
            Top             =   3240
            Width           =   7335
         End
         Begin VB.TextBox txtRutaGalenhos 
            Height          =   315
            Left            =   2040
            TabIndex        =   28
            Top             =   1110
            Width           =   5415
         End
         Begin VB.TextBox txtZipClave 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2040
            PasswordChar    =   "*"
            TabIndex        =   25
            Top             =   720
            Width           =   2025
         End
         Begin VB.TextBox txtZipArchivo 
            Height          =   315
            Left            =   2040
            TabIndex        =   23
            Text            =   "c:\Precios17017A01050711.zip"
            Top             =   330
            Width           =   5415
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "4) Pulsar Clic en el botón ""Aceptar (F2)"""
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   150
            TabIndex        =   36
            Top             =   5010
            Width           =   3300
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Deberá realizar los siguientes pasos:"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   180
            TabIndex        =   35
            Top             =   2910
            Width           =   2895
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "3) El archivo nuePrec.dbf debe de estar ubicado en 'Descomprime en'"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   150
            TabIndex        =   34
            Top             =   4710
            Width           =   5760
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "2) Renombrar archivo Pr????.dbf como PR.DBF (ubicado en 'Descomprime en')"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   150
            TabIndex        =   33
            Top             =   4410
            Width           =   6450
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "(Microsoft Visual Foxpro Driver --> tabla Libre)"
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   2040
            TabIndex        =   31
            Top             =   1740
            Width           =   3750
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "(Debe tener el ODBC llamado HIS que apunte a 'Ruta GalenHos')"
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   2040
            TabIndex        =   30
            Top             =   1470
            Width           =   5325
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Descomprime en"
            Height          =   210
            Left            =   180
            TabIndex        =   29
            Top             =   1170
            Width           =   1365
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(del archivo mconfig.dbf del Sismedv2)"
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   4230
            TabIndex        =   27
            Top             =   750
            Width           =   3180
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Clave (archivo Zip)"
            Height          =   210
            Left            =   180
            TabIndex        =   26
            Top             =   780
            Width           =   1500
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Archivo con Precios"
            Height          =   210
            Left            =   180
            TabIndex        =   24
            Top             =   390
            Width           =   1590
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Consideraciones:"
         Height          =   7095
         Left            =   135
         TabIndex        =   5
         Top             =   615
         Width           =   11445
         Begin VB.Frame Frame5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   90
            TabIndex        =   19
            Top             =   4830
            Width           =   11175
            Begin VB.CommandButton cmdSaldosNegativos 
               Caption         =   "..."
               Height          =   510
               Left            =   7485
               TabIndex        =   69
               Top             =   180
               Width           =   3555
            End
            Begin VB.ComboBox cmbAlmacen 
               Height          =   330
               Left            =   2100
               TabIndex        =   21
               Top             =   270
               Width           =   5310
            End
            Begin VB.CheckBox chkTodasFarmacias 
               Caption         =   "Todos las Farmacias"
               Height          =   255
               Left            =   150
               TabIndex        =   20
               Top             =   330
               Width           =   1965
            End
         End
         Begin VB.Frame Frame4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1995
            Left            =   150
            TabIndex        =   11
            Top             =   2820
            Width           =   11145
            Begin VB.TextBox txtAnio 
               Height          =   315
               Left            =   2610
               TabIndex        =   13
               Text            =   "Text1"
               Top             =   600
               Width           =   915
            End
            Begin VB.TextBox txtMes 
               Height          =   315
               Left            =   1380
               TabIndex        =   12
               Text            =   "Text1"
               Top             =   600
               Width           =   435
            End
            Begin Threed.SSOption optRegeneraMes 
               Height          =   345
               Left            =   150
               TabIndex        =   14
               Top             =   210
               Width           =   4815
               _ExtentX        =   8493
               _ExtentY        =   609
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
               Caption         =   "Regenera Saldos tomando SALDOS DEL ULTIMO MES"
               Value           =   -1
            End
            Begin Threed.SSOption optRegeneraInicio 
               Height          =   345
               Left            =   150
               TabIndex        =   15
               Top             =   1050
               Width           =   5715
               _ExtentX        =   10081
               _ExtentY        =   609
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
               Caption         =   "Regenera Saldos tomando SALDOS desde el primer INVENTARIO"
            End
            Begin Threed.SSOption optSoloSaldosMensualesDesdeInicio 
               Height          =   465
               Left            =   150
               TabIndex        =   16
               Top             =   1410
               Width           =   7305
               _ExtentX        =   12885
               _ExtentY        =   820
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
               Caption         =   "Regenera Saldos tomando SALDOS desde el primer INVENTARIO (solo los Mensuales)"
            End
            Begin VB.Label lblHoraFinal 
               Caption         =   "...."
               Height          =   225
               Left            =   8850
               TabIndex        =   44
               Top             =   660
               Width           =   1755
            End
            Begin VB.Label lblHoraInicio 
               Caption         =   "...."
               Height          =   225
               Left            =   8850
               TabIndex        =   43
               Top             =   240
               Width           =   1755
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Año"
               Height          =   210
               Left            =   2160
               TabIndex        =   18
               Top             =   630
               Width           =   330
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Mes"
               Height          =   210
               Left            =   1020
               TabIndex        =   17
               Top             =   630
               Width           =   315
            End
         End
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
            Height          =   1365
            Left            =   90
            TabIndex        =   7
            Top             =   5580
            Width           =   11175
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   345
               Left            =   120
               TabIndex        =   8
               Top             =   210
               Width           =   10875
               _ExtentX        =   19182
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   1
            End
            Begin MSComctlLib.ProgressBar ProgressBar2 
               Height          =   345
               Left            =   120
               TabIndex        =   9
               Top             =   600
               Width           =   10875
               _ExtentX        =   19182
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   1
            End
            Begin VB.Label lblProducto 
               Caption         =   "....."
               Height          =   210
               Left            =   120
               TabIndex        =   10
               Top             =   1020
               Width           =   7260
            End
         End
         Begin VB.ListBox cmbConsideraciones 
            BackColor       =   &H80000003&
            ForeColor       =   &H80000004&
            Height          =   2580
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   4455
         End
         Begin UltraGrid.SSUltraGrid grdAuditoriasRS 
            Height          =   2535
            Left            =   4590
            TabIndex        =   0
            Top             =   240
            Width           =   6690
            _ExtentX        =   11800
            _ExtentY        =   4471
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   71303188
            BorderStyle     =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   "HerrRegeneraSaldos.frx":0D56
            Caption         =   "20 últimas REGENERACIONES DE SALDOS"
         End
      End
      Begin UltraGrid.SSUltraGrid grdFarmSaldoDetallado 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   40
         Top             =   720
         Width           =   11430
         _ExtentX        =   20161
         _ExtentY        =   9128
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   71303188
         BorderStyle     =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   "HerrRegeneraSaldos.frx":0D92
         Caption         =   ".."
      End
      Begin UltraGrid.SSUltraGrid grdInventarios 
         Height          =   2985
         Left            =   -74880
         TabIndex        =   59
         Top             =   4080
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   5265
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   71303188
         BorderStyle     =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   "HerrRegeneraSaldos.frx":0DCE
         Caption         =   "Lista de Inventarios"
      End
      Begin MSComctlLib.ProgressBar ProgressBar5 
         Height          =   345
         Left            =   -74880
         TabIndex        =   62
         Top             =   7110
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label13 
         Caption         =   "Ingrese el CODIGO sismed y pulse ENTER"
         Height          =   465
         Left            =   -74880
         TabIndex        =   39
         Top             =   6030
         Width           =   1905
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
      Height          =   825
      Left            =   60
      TabIndex        =   1
      Top             =   7830
      Width           =   11715
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "HerrRegeneraSaldos.frx":0E0A
         DownPicture     =   "HerrRegeneraSaldos.frx":126A
         Height          =   645
         Left            =   4425
         Picture         =   "HerrRegeneraSaldos.frx":16DF
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   135
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HerrRegeneraSaldos.frx":1B54
         DownPicture     =   "HerrRegeneraSaldos.frx":2018
         Height          =   645
         Left            =   5970
         Picture         =   "HerrRegeneraSaldos.frx":2504
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   135
         Width           =   1335
      End
   End
End
Attribute VB_Name = "HerrRegeneraSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Herramienta para Regenerar Saldos, Actualizar Precios
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim lcDias As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_idUsuario As Long
Dim mo_lcNombrePc  As String
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
Dim ldUltimaDiaMesActual As Date
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasComunes As New ReglasComunes
Dim mo_ReglasFarmacia As New ReglasFarmacia
Dim mo_ReglasFacturacion As New ReglasFacturacion
Dim mo_cmbAlmacen As New sighentidades.ListaDespleglable
Dim oRsFarmSaldoDetallado As New ADODB.Recordset
Dim oRsInventarios As New Recordset
Dim mrs_Tmp As New ADODB.Recordset
Dim lcSql As String
Dim ml_FormularioUsadoDesdeOtroFrm As Boolean
Dim mo_IdAlmacenAregenerar As Long
Dim ml_RegeneraDesdeUltimoMes As Boolean
Dim ldHoy As Date
Dim ml_EsperaApulsarACEPTAR As Boolean
Property Let EsperaApulsarACEPTAR(lValue As Boolean)
    ml_EsperaApulsarACEPTAR = lValue
End Property

Property Let RegeneraDesdeUltimoMes(lIdValue As Boolean)
    ml_RegeneraDesdeUltimoMes = lIdValue
End Property

Property Let FormularioUsadoDesdeOtroFrm(lIdValue As Boolean)
    ml_FormularioUsadoDesdeOtroFrm = lIdValue
End Property
Property Let IdAlmacenAregenerar(lValue As Long)
   mo_IdAlmacenAregenerar = lValue
End Property



Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let idUsuario(lIdValue As Long)
    ml_idUsuario = lIdValue
End Property



Private Sub btnAceptar_Click()
    Dim lbInicioProceso As Boolean
    If SSTab1.Tab = 1 Then
       ActualizaPreciosParaFarmacia
       Exit Sub
    End If
    If SSTab1.Tab = 2 Then
       cmdActualizaFVencimiento_Click
       Exit Sub
    End If
    If SSTab1.Tab = 3 Then
       ExportaPreciosSismed
       Exit Sub
    End If
    If SSTab1.Tab = 4 Then
       ExportaInventariosSismed
       Exit Sub
    End If
    On Error GoTo ErrAcp
    If chkTodasFarmacias.Value = 0 Then
       If Me.cmbAlmacen.Text = "" Then
          MsgBox "Por favor elija el Almacén o Farmacia a procesar", vbInformation, Me.Caption
          Exit Sub
       End If
    End If
    lbInicioProceso = True
    If ml_FormularioUsadoDesdeOtroFrm <> True Then
       If MsgBox("Esta seguro que desea REGENERAR SALDOS", vbQuestion + vbYesNo, "Farmacia") <> vbYes Then
          lbInicioProceso = False
       End If
    End If
    
    If lbInicioProceso = True Then
        btnCancelar.Visible = False
        btnAceptar.Visible = False
        lblHoraInicio.Caption = lcBuscaParametro.RetornaHoraServidorSQL1
        Dim oRsTmp As New ADODB.Recordset
        Dim oRsTmp1 As New ADODB.Recordset
        Dim lcSql As String
        Dim lbContinua As Boolean
        Dim lnCant As Long
        Dim lcErrores As String
        Dim lnTotal As Long
        Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
        Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
        Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
        Dim lnAlmacen As Long: Dim lcAlmacen As String
        Dim oMovimientoDetalle As New farmMovimientoDetalle
        Dim ldUltimaDiaMesAnterior As Date
        Dim lnMes As Integer, lnAnio As Integer, lnMes1 As Integer, lnAnio1 As Integer
        Dim lbFin As Boolean, lnIdProducto As Long
        Dim lnIdAlmacen1 As Long, lnIdProducto1 As Long, lcLote1 As String, ldFechaVencimiento1 As Date
        Dim lnCantidad1 As Long, lnPrecio1 As Double, ldFechaMovimiento1 As Date, lbChequea As Boolean
        Dim ldUltimaDiaMesNow As Date
        Dim lcWhere As String, lcWhere1 As String, lbProsigue As Boolean
        Dim lnidTipoSalidaBienInsumo1 As Long, lbUsandoProcesoEnServidor As Boolean, lcWhere2 As String
        Me.MousePointer = 11
        lbUsandoProcesoEnServidor = False
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        oConexion.BeginTrans
        lcErrores = ""
        'Where si eligio solo un ALMACEN
        lcWhere = IIf(chkTodasFarmacias.Value = 1, "", " Where idAlmacen=" & mo_cmbAlmacen.BoundText)
        lcWhere1 = IIf(chkTodasFarmacias.Value = 1, "", " and idAlmacen=" & mo_cmbAlmacen.BoundText)
        'Elimina datos de las tablas: FarmSaldo y FarmSaldoDetallado
        If optSoloSaldosMensualesDesdeInicio.Value = False Then
            mo_ReglasFarmacia.SaldoEliminarTodosSegunFiltro lcWhere, oConexion
        End If
        ldUltimaDiaMesAnterior = sighentidades.DevuelveUltimaFechaDelMesAnteriorDelMovimiento(Now)
        ldUltimaDiaMesNow = sighentidades.DevuelveFechaHoraFinalDelMesDelMovimiento(Now)
        If optRegeneraInicio.Value = True Then
            mo_ReglasFarmacia.FarmMovimientoDetalleActualizaEspaciosDeCadaLote oConexion
            mo_ReglasFarmacia.SaldoMensualEliminarTodosSegunFiltro lcWhere, oConexion
            'Filtra todos los movimientos e/s
            lcWhere2 = IIf(chkTodasFarmacias.Value = 1, "", " and (dbo.farmMovimiento.idAlmacenDestino=" & mo_cmbAlmacen.BoundText & "  or dbo.farmMovimiento.idAlmacenOrigen=" & mo_cmbAlmacen.BoundText & ")")
            Set oRsTmp = mo_ReglasFarmacia.FarmaciaFiltraTodosMovimientos(lcWhere2, oConexion)
        ElseIf optSoloSaldosMensualesDesdeInicio.Value = True Then
            mo_ReglasFarmacia.FarmMovimientoDetalleActualizaEspaciosDeCadaLote oConexion
            mo_ReglasFarmacia.SaldoMensualEliminarTodosSegunFiltro lcWhere, oConexion
            'Filtra todos los movimientos e/s
            lcWhere2 = IIf(chkTodasFarmacias.Value = 1, "", " and (dbo.farmMovimiento.idAlmacenDestino=" & mo_cmbAlmacen.BoundText & "  or dbo.farmMovimiento.idAlmacenOrigen=" & mo_cmbAlmacen.BoundText & ")")
            Set oRsTmp = mo_ReglasFarmacia.FarmaciaFiltraTodosMovimientos(lcWhere2, oConexion)
        Else
            ldUltimaDiaMesActual = sighentidades.DevuelveFechaHoraFinalDelMesDelMovimiento(CDate("01/" & Right("0" & Me.txtMes.Text, 2) & "/" & Me.txtAnio.Text))
            ldUltimaDiaMesAnterior = sighentidades.DevuelveUltimaFechaDelMesAnteriorDelMovimiento(ldUltimaDiaMesActual)
            'elimina Saldos Mensuales, del mes actual
            mo_ReglasFarmacia.SaldoMensualEliminarTodosSegunFiltro " where SaldoFecha>'" & Format(ldUltimaDiaMesActual, sighentidades.DevuelveFechaSoloFormato_DMY_HMS) & "'" & lcWhere1, oConexion
            'Actualiza saldos iniciales (FarmSaldo/FarmSaldoDetallado)
            Set oRsTmp = mo_ReglasFarmacia.FarmSaldoMensualDetalladoSeleccionarPorSAldoFecha(Format(ldUltimaDiaMesActual, sighentidades.DevuelveFechaSoloFormato_DMY_HMS), lcWhere1, oConexion)
            If oRsTmp.RecordCount > 0 Then
               oRsTmp.MoveFirst
               Do While Not oRsTmp.EOF
                  If ActualizaSaldosPorProducto("E", oRsTmp.Fields!IdAlmacen, oRsTmp.Fields!idProducto, oRsTmp.Fields!Lote, oRsTmp.Fields!fechaVencimiento, oRsTmp.Fields!idTipoSalidaBienInsumo, oRsTmp.Fields!saldo, oRsTmp.Fields!precio, oRsTmp.Fields!SaldoFecha, False) = False Then
                       lcErrores = lcErrores & "Error en el Tipo Mov: (saldo inicial)    Almacen: " & oRsTmp.Fields!IdAlmacen & " Producto: " & oRsTmp.Fields!idProducto & " Lote: " & oRsTmp.Fields!Lote & " F.Vencimiento: " & oRsTmp.Fields!fechaVencimiento
                  End If
                  oRsTmp.MoveNext
               Loop
            End If
            oRsTmp.Close
            'Filtra los movimientos e/s a partir del 1 dia del mes actual
            Set oRsTmp = mo_ReglasFarmacia.FarmaciaFiltraTodosMovimientos(" and  (dbo.farmMovimiento.fechaCreacion Between CONVERT(DATETIME,'" & Format(ldUltimaDiaMesActual, "dd/mm/yyyy hh:mm:ss") & "',103) and CONVERT(DATETIME,'01/01/3000 23:23:59',103))", oConexion)
        End If
        'Proceso
        lbFin = False
        lnTotal = oRsTmp.RecordCount
        If lnTotal > 0 Then
            Set oMovimientoDetalle.Conexion = oConexion
            ProgressBar1.Min = 0
            ProgressBar1.Max = lnTotal
            lnCant = 1
            oRsTmp.MoveFirst
            lnMes = Month(oRsTmp.Fields!fechaCreacion)
            lnAnio = Year(oRsTmp.Fields!fechaCreacion)
            Do While Not oRsTmp.EOF
                DoEvents
                ProgressBar1.Value = lnCant
                Me.Refresh
                lnCant = lnCant + 1
                lbProsigue = True
                If chkTodasFarmacias.Value = 0 Then
                   If oRsTmp.Fields!MovTipo = "E" Then
                      If oRsTmp.Fields!IdAlmacenDestino <> Val(mo_cmbAlmacen.BoundText) Then
                         lbProsigue = False
                      End If
                   Else
                      If oRsTmp.Fields!IdAlmacenOrigen <> Val(mo_cmbAlmacen.BoundText) Then
                         lbProsigue = False
                      End If
                   End If
                End If
                If lbProsigue = True Then
                    If (Month(oRsTmp.Fields!fechaCreacion) + (Year(oRsTmp.Fields!fechaCreacion) * 12)) < (lnMes + (lnAnio * 12)) Then
                        lnMes = Month(oRsTmp.Fields!fechaCreacion)
                        lnAnio = Year(oRsTmp.Fields!fechaCreacion)
                    End If
                    If oRsTmp.Fields!MovTipo = "E" Then
                       lnAlmacen = oRsTmp.Fields!IdAlmacenDestino
                       lcAlmacen = oRsTmp.Fields!almDestino
                    Else
                       lnAlmacen = oRsTmp.Fields!IdAlmacenOrigen
                       lcAlmacen = oRsTmp.Fields!almOrigen
                    End If
                    If optSoloSaldosMensualesDesdeInicio.Value = True Then
                        If Not FarmActualizaSaldosMensualRS(oRsTmp.Fields!MovTipo, lnAlmacen, oRsTmp.Fields!idProducto, oRsTmp.Fields!fechaCreacion, oRsTmp.Fields!Cantidad, oRsTmp.Fields!Lote, oRsTmp.Fields!fechaVencimiento, oRsTmp.Fields!idTipoSalidaBienInsumo, oRsTmp.Fields!precio) = True Then
                           lcErrores = lcErrores & "Error en el Tipo Mov: " & oRsTmp.Fields!MovTipo & " N° Movimiento: " & oRsTmp.Fields!movNumero & " Almacen: " & oRsTmp.Fields!almDestino & " Producto: " & oRsTmp.Fields!codigo & " Lote: " & oRsTmp.Fields!Lote & " F.Vencimiento: " & oRsTmp.Fields!fechaVencimiento
                        End If
                    Else
                        If Not ActualizaSaldosPorProducto(oRsTmp.Fields!MovTipo, lnAlmacen, oRsTmp.Fields!idProducto, oRsTmp.Fields!Lote, oRsTmp.Fields!fechaVencimiento, oRsTmp.Fields!idTipoSalidaBienInsumo, oRsTmp.Fields!Cantidad, oRsTmp.Fields!precio, oRsTmp.Fields!fechaCreacion, True) Then
                           lcErrores = lcErrores & "Error en el Tipo Mov: " & oRsTmp.Fields!MovTipo & " N° Movimiento: " & oRsTmp.Fields!movNumero & " Almacen: " & oRsTmp.Fields!almDestino & " Producto: " & oRsTmp.Fields!codigo & " Lote: " & oRsTmp.Fields!Lote & " F.Vencimiento: " & oRsTmp.Fields!fechaVencimiento
                        End If
                    End If
                    lblProducto.Caption = Trim(oRsTmp.Fields!codigo) & "/" & Trim(lcAlmacen) & "/" & oRsTmp.Fields!MovTipo & "/" & oRsTmp.Fields!movNumero
                End If
                oRsTmp.MoveNext
            Loop
        End If
        oRsTmp.Close
        'En este momento solo estan registrados los movimientos del mes
        'Este proceso asignara el SALDO al final del mes
        If optRegeneraInicio.Value = True Then
           lcSql = lcWhere
        ElseIf optSoloSaldosMensualesDesdeInicio.Value = True Then
           lcSql = lcWhere
        Else
           lcSql = " Where SaldoFecha>=CONVERT(DATETIME,'" & Format(ldUltimaDiaMesActual, sighentidades.DevuelveFechaSoloFormato_DMY_HMS) & "',103)" & " and " & lcWhere1
        End If
        lblProducto.Caption = "....Espere....Procesando Totales Mensuales...."
        Set oRsTmp1 = mo_ReglasFarmacia.FarmSaldoMensualDetalladoSeleccionarPorFiltro(lcSql, oConexion)
        lnTotal = oRsTmp1.RecordCount
        If lnTotal > 0 Then
           ProgressBar2.Min = 0
           ProgressBar2.Max = lnTotal
           lnCant = 1
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              lnIdAlmacen1 = oRsTmp1.Fields!IdAlmacen
              lnIdProducto1 = oRsTmp1.Fields!idProducto
              Do While Not oRsTmp1.EOF And lnIdAlmacen1 = oRsTmp1.Fields!IdAlmacen And lnIdProducto1 = oRsTmp1.Fields!idProducto
                    lcLote1 = oRsTmp1.Fields!Lote
                    ldFechaVencimiento1 = oRsTmp1.Fields!fechaVencimiento
                    lnidTipoSalidaBienInsumo1 = oRsTmp1.Fields!idTipoSalidaBienInsumo
                    lnMes1 = Month(oRsTmp1.Fields!SaldoFecha)
                    lnAnio1 = Year(oRsTmp1.Fields!SaldoFecha)
                    lnCantidad1 = oRsTmp1.Fields!saldo
                    lnPrecio1 = oRsTmp1.Fields!precio
                    Do While Not oRsTmp1.EOF And lnIdAlmacen1 = oRsTmp1.Fields!IdAlmacen And lnIdProducto1 = oRsTmp1.Fields!idProducto And lcLote1 = oRsTmp1.Fields!Lote And ldFechaVencimiento1 = oRsTmp1.Fields!fechaVencimiento And lnidTipoSalidaBienInsumo1 = oRsTmp1.Fields!idTipoSalidaBienInsumo
                       If (lnMes1 + (lnAnio1 * 12)) <> (Month(oRsTmp1.Fields!SaldoFecha) + (Year(oRsTmp1.Fields!SaldoFecha) * 12)) Then
                            Do While True
                                If lnMes1 = 12 Then
                                     lnMes1 = 1
                                     lnAnio1 = lnAnio1 + 1
                                Else
                                     lnMes1 = lnMes1 + 1
                                End If
                                If (lnMes1 + (lnAnio1 * 12)) = (Month(oRsTmp1.Fields!SaldoFecha) + (Year(oRsTmp1.Fields!SaldoFecha) * 12)) Then
                                    ldFechaMovimiento1 = CDate("01/" & Right("0" + Trim(str(Month(oRsTmp1.Fields!SaldoFecha))), 2) + "/" + Trim(str(Year(oRsTmp1.Fields!SaldoFecha))))
                                    If FarmActualizaSaldosMensualRS("E", lnIdAlmacen1, lnIdProducto1, ldFechaMovimiento1, lnCantidad1, lcLote1, ldFechaVencimiento1, lnidTipoSalidaBienInsumo1, lnPrecio1) = True Then
                                         ms_MensajeError = ""
                                    Else
                                        lcErrores = lcErrores & ms_MensajeError
                                    End If
                                    lnCantidad1 = lnCantidad1 + oRsTmp1.Fields!saldo
                                    lnPrecio1 = oRsTmp1.Fields!precio
                                    Exit Do
                                Else
                                    ldFechaMovimiento1 = CDate("01/" & Right("0" + Trim(str(lnMes1)), 2) + "/" + Trim(str(lnAnio1)))
                                    If FarmActualizaSaldosMensualRS("E", lnIdAlmacen1, lnIdProducto1, ldFechaMovimiento1, lnCantidad1, lcLote1, ldFechaVencimiento1, lnidTipoSalidaBienInsumo1, lnPrecio1) = True Then
                                         ms_MensajeError = ""
                                    Else
                                        lcErrores = lcErrores & ms_MensajeError
                                    End If
                                End If
                                If (lnMes1 + (lnAnio1 * 12)) = (Month(ldUltimaDiaMesNow) + (Year(ldUltimaDiaMesNow) * 12)) Then
                                    Exit Do
                                End If
                            Loop
                       Else
                            lnCantidad1 = oRsTmp1.Fields!saldo
                            lnPrecio1 = oRsTmp1.Fields!precio
                       End If
                       DoEvents
                       ProgressBar2.Value = lnCant
                       Me.Refresh
                       lnCant = lnCant + 1
                       oRsTmp1.MoveNext
                       If oRsTmp1.EOF Then
                          Exit Do
                       End If
                    Loop
                    If (lnMes1 + (lnAnio1 * 12)) < (Month(ldUltimaDiaMesNow) + (Year(ldUltimaDiaMesNow) * 12)) Then
                        Do While True
                            If lnMes1 = 12 Then
                                 lnMes1 = 1
                                 lnAnio1 = lnAnio1 + 1
                            Else
                                 lnMes1 = lnMes1 + 1
                            End If
                            ldFechaMovimiento1 = CDate("01/" & Right("0" + Trim(str(lnMes1)), 2) + "/" + Trim(str(lnAnio1)))
                            If FarmActualizaSaldosMensualRS("E", lnIdAlmacen1, lnIdProducto1, ldFechaMovimiento1, lnCantidad1, lcLote1, ldFechaVencimiento1, lnidTipoSalidaBienInsumo1, lnPrecio1) = True Then
                                 ms_MensajeError = ""
                            Else
                                lcErrores = lcErrores & ms_MensajeError
                            End If
                            If (lnMes1 + (lnAnio1 * 12)) = (Month(ldUltimaDiaMesNow) + (Year(ldUltimaDiaMesNow) * 12)) Then
                                Exit Do
                            End If
                        Loop
                    End If
                    If oRsTmp1.EOF Then
                       Exit Do
                    End If
              Loop
           Loop
        End If
        '
        If ml_idUsuario > 0 Then
           Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "M", 0, "FarmSaldo", oConexion, 500 + 15, mo_lcNombrePc, _
                     "Reg.Saldos: " & IIf(optRegeneraMes.Value = True, "ULTIMO MES", _
                                      IIf(optRegeneraInicio.Value = True, "PRIMER INV", "MENSUALES")) & _
                     "/" & IIf(chkTodasFarmacias.Value = 1, "TODAS.FARM", Left(cmbAlmacen.Text, 30)))
        End If
        '
        oConexion.CommitTrans
        'arregla PRECIOS PONDERADOS negativos
        mo_ReglasArchivoClinico.farmActualizaPrecioPonderadoMenorAcero oConexion
        'arregla SALDOS LOTES NEGATIVOS y lo inserta en LOTES POSITIVOS
        If cmbAlmacen.Text = "" Then
           Dim lnFor As Integer
           For lnFor = 1 To cmbAlmacen.ListCount
               cmbAlmacen.ListIndex = lnFor - 1
               cmdSaldosNegativos_Click
           Next
        Else
           cmdSaldosNegativos_Click
        End If
        '
        oConexion.Close
        Set oRsTmp = Nothing
        Set oConexion = Nothing
        Set mo_ReglasFarmacia = Nothing
        Set oMovimientoDetalle = Nothing
        Set mo_ReglasArchivoClinico = Nothing
        If lcErrores <> "" Then
           MsgBox lcErrores, vbInformation, "Cierre"
        End If
        lblProducto.Caption = "............Generando impresión de ERRORES............."
        '
        'Dim mo_Imprime As New RegeneraSaldos
        'mo_Imprime.CrearReporte_excel lbUsandoProcesoEnServidor, Me.hwnd
        If ml_FormularioUsadoDesdeOtroFrm = True Then
            LimpiarVariablesDeMemoria
            Unload Me
        Else
            CrearReporte_excel lbUsandoProcesoEnServidor, Me.hwnd
            '
            Me.MousePointer = 1
            lblHoraFinal.Caption = lcBuscaParametro.RetornaHoraServidorSQL1
            lcErrores = DateDiff("s", CDate(lblHoraInicio.Caption), CDate(lblHoraFinal.Caption))
    '        MsgBox "Termino el Proceso OK" & Chr(13) & Chr(13) & "Segundos de Diferencia: " & lcErrores, vbInformation, "Mensaje"
            Me.Visible = False
            LimpiarVariablesDeMemoria
        End If
    End If
    Exit Sub
ErrAcp:
    oConexion.RollbackTrans
    MsgBox Err.Description
   Resume
End Sub



Sub CrearReporte_excel(lbUsandoProcesoEnServidor As Boolean, lnHwnd As Long)
Dim rsReporte As New Recordset
Dim rsReporte1 As New Recordset
Dim iFila As Long
Dim lnTotal As Double
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReporteUtil As New sighentidades.ReporteUtil
Dim oConexion As New ADODB.Connection
Dim lcSql As String, lcHoraInicio As String, lcHoraFinal As String
Dim lbEsOpenOffice As Boolean

lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
    On Error GoTo ManejadorErrorExcel
   
    If lbEsOpenOffice = True Then
        Dim ServiceManager As Object
        Dim Desktop As Object
        Dim Document As Object
        Dim Feuille As Object
        Dim Plage As Object
        Dim args()
        Dim Chemin As String
        Dim Fichier As String
        Dim lcArchivoExcel As String
        Dim PrintArea(0)
        Dim Style As Object
        Dim Border As Object
        'encabezado
        Dim PageStyles As Object
        Dim Sheet As Object
        Dim StyleFamilies As Object
        Dim DefPage As Object
        Dim Htext As Object
        Dim Hcontent As Object
        Dim ret As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    GenerarRecordsetTemporal
    '***************Error- Cabecera sin  Detalle
    lcHoraInicio = lcBuscaParametro.RetornaHoraServidorSQL1
    '**********Errores de Stock con Cantidades negativas
    Set rsReporte = mo_ReglasFarmacia.SaldosNegativos
    If rsReporte.RecordCount > 0 Then
        rsReporte.MoveFirst
        Do While Not rsReporte.EOF
            mrs_Tmp.AddNew
            mrs_Tmp.Fields("codigo").Value = "Existe LOTE con Saldos NEGATIVOS"
            mrs_Tmp.Fields!descrip = "Almacen: " & Trim(rsReporte.Fields!Descripcion) & "    Producto: " & rsReporte.Fields!codigo & " " & Trim(rsReporte.Fields!nombre) & "    Cantidad: " & rsReporte.Fields!Cantidad
            mrs_Tmp.Fields!solucion = "Chequee los Movimientos de E/S (KARDEX), debe faltar algun Ingreso"
            mrs_Tmp.Update
            rsReporte.MoveNext
        Loop
    End If
    rsReporte.Close

        Set rsReporte = FarmNiNsConCabeceraSinDetalle(oConexion)
        rsReporte.Filter = "cantidad=0"
        If rsReporte.RecordCount > 0 Then
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
                  mrs_Tmp.AddNew
                  mrs_Tmp.Fields!codigo = "Cabecera sin Detalle"
                  mrs_Tmp.Fields!descrip = IIf(rsReporte.Fields!MovTipo = "E", "Nota Ingreso: ", "Nota Salida: ") & rsReporte.Fields!movNumero & "    Fecha: " & rsReporte.Fields!fechaCreacion
                  mrs_Tmp.Fields!solucion = "Fijese si el Dcto existe fisicamente, Anulelo o  Registre los Productos"
                  mrs_Tmp.Update
               rsReporte.MoveNext
            Loop
        End If
        rsReporte.Close
'
    lcHoraFinal = lcBuscaParametro.RetornaHoraServidorSQL1
    If mrs_Tmp.RecordCount > 0 Then
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\farmErrores.ods"
'            FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'            Chemin = "file:///" & App.Path & "\Plantillas\"
'            Chemin = Replace(Chemin, "\", "/")
'            Fichier = Chemin & "/OpenOffice.ods"
            Fichier = Format(Time, "hhmmss") & ".ods"
            FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
            lcArchivoExcel = Fichier
            Chemin = "file:///" & App.Path & "\Plantillas\"
            Chemin = Replace(Chemin, "\", "/")
            Fichier = Chemin & "/" & lcArchivoExcel
            '
            Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
            Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
            Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
            Set Feuille = Document.getSheets().getByIndex(0)
            'Encabezado de Pagina
            mo_CabeceraReportes.CabeceraReportes Document, True
            ' Pone la ventana en primer plano, pasándole el Hwnd
            ret = SetForegroundWindow(lnHwnd)
        Else
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            'Crea nueva hoja
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\farmErrores.xls")
            oWorkBookPlantilla.Worksheets("farmErrores").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
        iFila = 6: lnTotal = 0
        mrs_Tmp.MoveFirst
        Do While Not mrs_Tmp.EOF
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula(mrs_Tmp.Fields("codigo").Value)
                Call Feuille.getcellbyposition(2, iFila - 1).setFormula(mrs_Tmp.Fields("descrip").Value)
                Call Feuille.getcellbyposition(3, iFila - 1).setFormula(mrs_Tmp.Fields("solucion").Value)
            Else
                oWorkSheet.Cells(iFila, 2).Value = mrs_Tmp.Fields("codigo").Value
                oWorkSheet.Cells(iFila, 3).Value = mrs_Tmp.Fields("descrip").Value
                oWorkSheet.Cells(iFila, 4).Value = mrs_Tmp.Fields("solucion").Value
            End If
           iFila = iFila + 1
           mrs_Tmp.MoveNext
        Loop
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":D" & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Nº Errores: " + Trim(str(mrs_Tmp.RecordCount)))
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 4
            oWorkSheet.Cells(iFila, 2).Value = "Nº Errores: " + Trim(str(mrs_Tmp.RecordCount))
        End If
        If lbEsOpenOffice = True Then
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
            oExcel.Visible = True
            oWorkSheet.PrintPreview
            'oWorkSheet.PrintOut
        End If
    End If
    If lbEsOpenOffice = True Then
        'Liberar Memoria
        Set Plage = Nothing
        Set Feuille = Nothing
        Set Document = Nothing
        Set Desktop = Nothing
        Set ServiceManager = Nothing
        Set Style = Nothing
        Set Border = Nothing
        'encabezado de pagina
        Set PageStyles = Nothing
        Set Sheet = Nothing
        Set StyleFamilies = Nothing
        Set DefPage = Nothing
        Set Htext = Nothing
        Set Hcontent = Nothing
    Else
        'Liberar memoria
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
Exit Sub
ManejadorErrorExcel:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub
Resume
End Sub
Sub GenerarRecordsetTemporal()
    With mrs_Tmp
          .Fields.Append "codigo", adVarChar, 50, adFldIsNullable
          .Fields.Append "Descrip", adVarChar, 200, adFldIsNullable
          .Fields.Append "Solucion", adVarChar, 200, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
End Sub
Function FarmNiNsConCabeceraSinDetalle(mo_Conexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
Dim oConexion As New ADODB.Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set FarmNiNsConCabeceraSinDetalle = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmNiNsConCabeceraSinDetalle"
        Set oRecordset = .Execute
   End With
   Set FarmNiNsConCabeceraSinDetalle = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub



Private Sub chkTodasFarmacias_Click()
    If chkTodasFarmacias.Value = 1 Then
       cmbAlmacen.Visible = False
    Else
       cmbAlmacen.Visible = True
    End If
End Sub



Private Sub cmdActualizaFVencimiento_Click()
        On Error GoTo errActFV
        If oRsFarmSaldoDetallado.RecordCount = 0 Then
           MsgBox "No hay Items en la LISTA"
           Exit Sub
        End If
        If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
           Dim oRsTmp As New ADODB.Recordset
           oRsFarmSaldoDetallado.MoveFirst
           Do While Not oRsFarmSaldoDetallado.EOF
              If Not IsNull(oRsFarmSaldoDetallado.Fields!fechaVencimientoN) Then
                 mo_ReglasFarmacia.FarmaciaActualizaFechaDeVencimiento oRsFarmSaldoDetallado.Fields!fechaVencimientoN, _
                                                                     oRsFarmSaldoDetallado.Fields!idProducto, _
                                                                     oRsFarmSaldoDetallado.Fields!Lote, _
                                                                     oRsFarmSaldoDetallado.Fields!fechaVencimiento
                    
              End If
              oRsFarmSaldoDetallado.MoveNext
           Loop
           Unload Me
        End If
        Exit Sub
errActFV:
    MsgBox Err.Description

End Sub

Private Sub cmdDescomprimir_Click()
        sighentidades.DescomprimeArchivoZIP Me.txtZipClave.Text, Me.txtZipArchivo.Text, Me.txtRutaGalenhos.Text, True

End Sub

Private Sub cmdSaldosNegativos_Click()
        If Val(mo_cmbAlmacen.BoundText) = 0 Then
           MsgBox "Tiene que elegir el ALMACEN o FARMACIA a poner SALDOS POSITIVOS en tabla: FarmSaldosDetallados"
           Exit Sub
        End If
        On Error GoTo errSalNweg
        Dim oRsTmp1 As New Recordset
        Dim oRsTmp2 As New Recordset
        Dim oRsTmp3 As New Recordset
        Dim oConexion As New Connection
        Dim lnIdProducto  As Long, lnIdAlmacen As Long, lnIdTipoSalidaBienInsumo As Long
        Dim lnCantidadTotal  As Double, lnSaldoNegativo As Double, lnSaldoPositivo As Double
        Dim lcSql As String
        oConexion.CommandTimeout = 900
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        lcSql = "select * from farmSaldoDetallado where cantidad<0 and idAlmacen=" & mo_cmbAlmacen.BoundText & " order by idProducto,idAlmacen,idTipoSalidaBienInsumo"
        oRsTmp1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
        If oRsTmp1.RecordCount > 0 Then
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
             lnIdProducto = oRsTmp1!idProducto
             lnIdAlmacen = oRsTmp1!IdAlmacen
             lnIdTipoSalidaBienInsumo = oRsTmp1!idTipoSalidaBienInsumo
             lnSaldoNegativo = 0
             Do While Not oRsTmp1.EOF And lnIdProducto = oRsTmp1!idProducto And lnIdAlmacen = oRsTmp1!IdAlmacen And lnIdTipoSalidaBienInsumo = oRsTmp1!idTipoSalidaBienInsumo
                  lnSaldoNegativo = lnSaldoNegativo + oRsTmp1!Cantidad
                  oRsTmp1.MoveNext
                  If oRsTmp1.EOF Then
                     Exit Do
                  End If
              Loop
              
              lnCantidadTotal = 0
              lcSql = "select sum(cantidad) as tot from farmSaldoDetallado where idProducto =" & lnIdProducto & " and idAlmacen =" & lnIdAlmacen & " and IdTipoSalidaBienInsumo =" & lnIdTipoSalidaBienInsumo
              If oRsTmp2.State = 1 Then oRsTmp2.Close
              oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
              If oRsTmp2.RecordCount > 0 Then
                 lnCantidadTotal = oRsTmp2!tot
              End If
              lnSaldoPositivo = lnCantidadTotal + lnSaldoNegativo
              If lnCantidadTotal >= 0 Then
                    lcSql = "select  * from farmSaldoDetallado where idProducto =" & lnIdProducto & " and idAlmacen =" & lnIdAlmacen & " and IdTipoSalidaBienInsumo =" & lnIdTipoSalidaBienInsumo & _
                               "order by cantidad"
                    If oRsTmp3.State = 1 Then oRsTmp3.Close
                    oRsTmp3.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                    lnSaldoPositivo = -lnSaldoNegativo
                    Do While Not oRsTmp3.EOF
                          If oRsTmp3!Cantidad >= lnSaldoPositivo Then
                                oRsTmp3!Cantidad = oRsTmp3!Cantidad - lnSaldoPositivo
                                oRsTmp3.Update
                                Exit Do
                          ElseIf oRsTmp3!Cantidad <= 0 Then
                                oRsTmp3!Cantidad = 0
                                 oRsTmp3.Update
                          Else
                                 lnSaldoPositivo = lnSaldoPositivo - oRsTmp3!Cantidad
                                 oRsTmp3!Cantidad = 0
                                 oRsTmp3.Update
                          End If
                          oRsTmp3.MoveNext
                    Loop
              End If
        Loop
        End If
        oRsTmp1.Close
        MsgBox "Procesó correctamente", vbInformation, ""
errSalNweg:
        Set oRsTmp1 = Nothing
        Set oRsTmp2 = Nothing
        Set oRsTmp3 = Nothing
        Set oConexion = Nothing
        'MsgBox "Terminó de procesar"
        'Me.Visible = False

End Sub

Private Sub Form_Activate()
    If ml_FormularioUsadoDesdeOtroFrm = True Then
       SSTab1.Tab = 0
       If ml_RegeneraDesdeUltimoMes = True Then
          optRegeneraMes.Value = True
       Else
          optRegeneraInicio.Value = True
       End If
       mo_cmbAlmacen.BoundText = mo_IdAlmacenAregenerar
       Me.Refresh
       If ml_EsperaApulsarACEPTAR = False Then
          btnAceptar_Click
       End If
    End If
End Sub

Sub ExportaInventariosSismed()
    oRsInventarios.Filter = "exportar=true"
    If oRsInventarios.RecordCount = 0 Then
       MsgBox "Tiene que elegir UNO o GRUPO DE INVENTARIOS", vbInformation, ""
    Else
       Dim lbSigue As Boolean, lcCodigoSismed As String, lcNumeroInventario As String
       Dim lcInventarios As String, lcAlmCod As String, lcInvNumero As String, lcCodMed As String
       Dim ldFechCie As Date, ldFechReg As Date, ldfechUlt As Date, lcLote As String, ldFechaVencimiento As Date
       Dim lnSobranteItem As Long, lnFaltanteItem As Long, lnCantidadSaldoItem As Long, lnCantidadItem As Long, lnInvPrecio As Double
       Dim ldFechPrc As Date
       Dim lcInvTipo As String, lcInvindiprc As String, lnInvnumelst As Integer, lcInvindicie As String, lcInvsitua As String
       Dim lcLabCod As String
       lbSigue = True
       Me.MousePointer = 11
'       If oRsInventarios.RecordCount > 1 Then
'          oRsInventarios.MoveFirst
'          lcAlmCod = oRsInventarios!CodigoSismed
'          lcInvNumero = oRsInventarios!NumeroInventario
'          lcCodigoSismed = oRsInventarios!CodigoSismed
'          lcNumeroInventario = oRsInventarios!NumeroInventario
'          ldFechCie = oRsInventarios!FechaCierre
'          lcInventarios = "dbo.farmInventarioDetalle.idInventario=" & oRsInventarios!IdInventario
'          ldfechUlt = oRsInventarios!fechaModificacion
'          lcInvTipo = IIf(oRsInventarios!idTipoInventario = 1, "M", "A")
'          ldFechPrc = oRsInventarios!fechaCreacion
'          ldFechReg = oRsInventarios!fechaCreacion
'          lcInvindiprc = IIf(Month(oRsInventarios!fechaCreacion) = 1 Or Month(oRsInventarios!fechaCreacion) = 12, "F", "I")
'          lnInvnumelst = 1
'          lcInvindicie = "C"
'          lcInvsitua = "1"
'          lcLabCod = Space(5)
'          Do While Not oRsInventarios.EOF
'             If lcCodigoSismed <> oRsInventarios!CodigoSismed Then
'                lbSigue = False
'                MsgBox "Eligió varios inventarios, todos deben tener el mismo CODIGO SISMED", vbInformation, ""
'                Exit Do
'             End If
'             If lcNumeroInventario <> oRsInventarios!NumeroInventario Then
'                lbSigue = False
'                MsgBox "Eligió varios inventarios, todos deben tener el mismo NUMERO DE INVENTARIO", vbInformation, ""
'                Exit Do
'             End If
'             If ldFechCie <> oRsInventarios!FechaCierre Then
'                lbSigue = False
'                MsgBox "Eligió varios inventarios, todos deben tener la misma FECHA DE CIERRE", vbInformation, ""
'                Exit Do
'             End If
'             oRsInventarios.MoveNext
'             If Not oRsInventarios.EOF Then
'                lcInventarios = lcInventarios & " and dbo.farmInventarioDetalle.idInventario=" & oRsInventarios!IdInventario
'             End If
'          Loop
'       Else
          lcInventarios = "dbo.farmInventarioDetalle.idInventario=" & oRsInventarios!IdInventario
          lcAlmCod = oRsInventarios!CodigoSismed
          lcInvNumero = oRsInventarios!NumeroInventario
          lcCodigoSismed = oRsInventarios!CodigoSismed
          lcNumeroInventario = oRsInventarios!NumeroInventario
          ldFechCie = oRsInventarios!FechaCierre
          lcInventarios = "dbo.farmInventarioDetalle.idInventario=" & oRsInventarios!IdInventario
          ldfechUlt = oRsInventarios!fechaModificacion
          lcInvTipo = IIf(oRsInventarios!idTipoInventario = 1, "M", "A")
          ldFechPrc = oRsInventarios!fechaCreacion
          ldFechReg = oRsInventarios!fechaCreacion
          lcInvindiprc = IIf(Month(oRsInventarios!fechaCreacion) = 1 Or Month(oRsInventarios!fechaCreacion) = 12, "F", "I")
          lnInvnumelst = 1
          lcInvindicie = "C"
          lcInvsitua = "1"
          lcLabCod = Space(5)
'       End If
       If lbSigue = True Then
          On Error GoTo errExInv
          Dim oConexionFox As New Connection
          Dim oRsInvent As New Recordset
          Dim oRsInventCB As New Recordset
          Dim oRsInventDE As New Recordset
          Dim oRsInventConf As New Recordset
          Dim oRsDe As New Recordset
          Dim lcNombreArchivo As String, lcTabla As String, lcRegSan As String, lnCantidad As Long
          Dim lnTotalReg As Long, lnCantidadSaldo As Long
          oConexionFox.CommandTimeout = 300
          oConexionFox.Open "DSN=his"
          '
          lcTabla = "tinv"
          lcNombreArchivo = txtRutaArchivos.Text & "\" & lcTabla & ".dbf"
          FileCopy txtRutaArchivos.Text & "\inv.dbf", lcNombreArchivo
          lcSql = "select * from " & lcTabla
          oRsInvent.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
          '
          lcTabla = "tinvCb"
          lcNombreArchivo = txtRutaArchivos.Text & "\" & lcTabla & ".dbf"
          FileCopy txtRutaArchivos.Text & "\invCb.dbf", lcNombreArchivo
          lcSql = "select * from " & lcTabla
          oRsInventCB.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
          '
          lcTabla = "tinvDe"
          lcNombreArchivo = txtRutaArchivos.Text & "\" & lcTabla & ".dbf"
          FileCopy txtRutaArchivos.Text & "\invDe.dbf", lcNombreArchivo
          lcSql = "select * from " & lcTabla
          oRsInventDE.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
          '
          lcTabla = "confi"
          lcNombreArchivo = txtRutaArchivos.Text & "\" & lcTabla & ".dbf"
          FileCopy txtRutaArchivos.Text & "\invCon.dbf", lcNombreArchivo
          lcSql = "select * from " & lcTabla
          oRsInventConf.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
          If oRsInventConf.RecordCount = 0 Then
             MsgBox "El archivo confi.dbf del SISMEDV2 debe tener al menos 1 registro", vbInformation, ""
          Else
            oRsInventConf.MoveFirst
            Do While Not oRsInventConf.EOF
               oRsInventConf!fecha = Format(ldHoy, sighentidades.DevuelveFechaSoloFormato_DMY)
               oRsInventConf!hora = Format(ldHoy, sighentidades.DevuelveHoraSoloFormato_HM)
               oRsInventConf.Update
               oRsInventConf.MoveNext
            Loop
            '
            Set oRsDe = mo_ReglasFarmacia.farmInventarioDetallePorFiltro(lcInventarios)
            lnTotalReg = oRsDe.RecordCount
            If lnTotalReg > 0 Then
               ProgressBar5.Min = 0
               ProgressBar5.Max = lnTotalReg
               ProgressBar5.Value = 0
               oRsDe.MoveFirst
               Do While Not oRsDe.EOF
                  lcCodMed = oRsDe!codigo
                  lnCantidadItem = 0
                  lnCantidadSaldoItem = 0
                  lnSobranteItem = 0
                  lnFaltanteItem = 0
                  Do While Not oRsDe.EOF And lcCodMed = oRsDe!codigo
                     lcLote = oRsDe!Lote
                     ldFechaVencimiento = oRsDe!fechaVencimiento
                     lcRegSan = Trim(oRsDe!registroSanitario)
                     lnCantidad = 0
                     lnCantidadSaldo = 0
                     lnInvPrecio = oRsDe!precio
                     Do While Not oRsDe.EOF And lcCodMed = oRsDe!codigo And lcLote = oRsDe!Lote And ldFechaVencimiento = oRsDe!fechaVencimiento
                        DoEvents: ProgressBar5.Value = ProgressBar5.Value + 1: Me.Refresh
                        lnCantidad = lnCantidad + oRsDe!Cantidad
                        lnCantidadSaldo = lnCantidadSaldo + oRsDe!cantidadSaldo
                        lnCantidadItem = lnCantidadItem + oRsDe!Cantidad
                        lnCantidadSaldoItem = lnCantidadSaldoItem + oRsDe!cantidadSaldo
                        lnSobranteItem = lnSobranteItem + oRsDe!cantidadSobrante
                        lnFaltanteItem = lnFaltanteItem + oRsDe!cantidadFaltante
                        oRsDe.MoveNext
                        If oRsDe.EOF Then
                           Exit Do
                        End If
                     Loop
                     oRsInventDE.AddNew
                     oRsInventDE!almcod = lcAlmCod
                     oRsInventDE!invnumero = lcInvNumero
                     oRsInventDE!medcod = lcCodMed
                     oRsInventDE!labcod = lcLabCod
                     oRsInventDE!invlote = lcLote
                     oRsInventDE!invnumelst = Right("0" & Trim(str(lnInvnumelst)), 2)
                     oRsInventDE!invfechvto = ldFechaVencimiento
                     oRsInventDE!invregsan = lcRegSan
                     oRsInventDE!invcantid = lnCantidad
                     oRsInventDE!invcontfis = lnCantidadSaldo
                     oRsInventDE!invindiprc = lcInvindiprc
                     oRsInventDE!invsitua = lcInvsitua
                     oRsInventDE!invalter = 0
                     oRsInventDE!salajus = 0
                     oRsInventDE.Update
                     If oRsDe.EOF Then
                        Exit Do
                     End If
                  Loop
                  oRsInventCB.AddNew
                  oRsInventCB!almcod = lcAlmCod
                  oRsInventCB!invnumero = lcInvNumero
                  oRsInventCB!medcod = lcCodMed
                  oRsInventCB!invnumelst = Right("0" & Trim(str(lnInvnumelst)), 2)
                  oRsInventCB!invcantid = lnCantidadItem
                  oRsInventCB!invprecio = lnInvPrecio
                  oRsInventCB!invtotact = Round(lnCantidadItem * lnInvPrecio, 6)
                  oRsInventCB!invcontfis = lnCantidadSaldoItem
                  oRsInventCB!invcantfal = lnFaltanteItem
                  oRsInventCB!invcantsob = lnSobranteItem
                  oRsInventCB!invtotal = Round(lnCantidadSaldoItem * lnInvPrecio, 6)
                  oRsInventCB!invindiprc = Space(1)  'lcInvindiprc
                  oRsInventCB!invfechreg = ldFechReg
                  oRsInventCB!invfechult = ldfechUlt
                  oRsInventCB!invsitua = lcInvsitua
                  oRsInventCB!invalter = 0
                  oRsInventCB.Update
               Loop
               oRsInvent.AddNew
               oRsInvent!almcod = lcAlmCod
               oRsInvent!invnumero = lcInvNumero
               oRsInvent!invtipo = lcInvTipo
               oRsInvent!invindiprc = lcInvindiprc
               oRsInvent!invfechprc = ldFechPrc
               oRsInvent!invnumelst = lnInvnumelst
               oRsInvent!invindicie = lcInvindicie
               oRsInvent!invfechcie = ldFechCie
               oRsInvent!invfechreg = ldFechReg
               oRsInvent!invfechult = ldfechUlt
               oRsInvent!invsitua = lcInvsitua
               oRsInvent!invindiaju = Space(1)
               oRsInvent.Update
            Else
               MsgBox "No existe datos", vbInformation, ""
            End If
          End If
          Set oConexionFox = Nothing
          Set oRsInvent = Nothing
          Set oRsInventCB = Nothing
          Set oRsInventDE = Nothing
          Set oRsDe = Nothing
          Set oRsInventConf = Nothing
          '
          ExportaArchivoZIP 1
          '
          Me.Visible = False
          Me.MousePointer = 1
          Exit Sub
       End If
    End If
    oRsInventarios.Filter = ""
    Exit Sub
errExInv:
    MsgBox Err.Description
    Me.MousePointer = 1
    Exit Sub
    Resume
End Sub

Sub ExportaArchivoZIP(lbDesde As Integer)
    Dim lcRutaExportar As String
    Dim lcTempo As Object
    Dim lcFolder As String
    Dim EXPORTAR_RUTA As String
    Dim lcArchivoExpZip As String
    Dim oCrypKey As New CrypKey.Util
    EXPORTAR_RUTA = lcBuscaParametro.SeleccionaFilaParametro(313)
    Set lcTempo = CreateObject("Scripting.FileSystemObject")
    lcFolder = EXPORTAR_RUTA & "ExportaPreciosSismedTem"
    lcTempo.CreateFolder lcFolder
    Select Case lbDesde
    Case 1
          lcArchivoExpZip = "inv" & lcBuscaParametro.SeleccionaFilaParametro(280) & "F01" & Format(ldHoy, "ddmmyy") & ".zip"
          FileCopy txtRutaArchivos.Text & "\confi.dbf", lcFolder & "\confi.dbf"
          FileCopy txtRutaArchivos.Text & "\malma.dbf", lcFolder & "\malma.dbf"
          FileCopy txtRutaArchivos.Text & "\tinv.dbf", lcFolder & "\tinv.dbf"
          FileCopy txtRutaArchivos.Text & "\tinvcb.dbf", lcFolder & "\tinvcb.dbf"
          FileCopy txtRutaArchivos.Text & "\tinvde.dbf", lcFolder & "\tinvde.dbf"
    Case 2
    End Select
    sighentidades.ComprimeArchivoZip EXPORTAR_RUTA & lcArchivoExpZip, _
                  oCrypKey.DecryptString(lcBuscaParametro.SeleccionaFilaParametro(350)), lcFolder & "\*.*"
    MsgBox "Se creó el archivo: " & EXPORTAR_RUTA & lcArchivoExpZip
    lcTempo.DeleteFolder lcFolder  'elimino carpeta de creacion de textos
    Set oCrypKey = Nothing
End Sub

Sub CreaTemporalYcargaIventarios()
  With oRsInventarios
      .Fields.Append "idInventario", adInteger
      .Fields.Append "Exportar", adBoolean
      .Fields.Append "CodigoSismed", adVarChar, 11
      .Fields.Append "numeroInventario", adVarChar, 20
      .Fields.Append "FechaCierre", adDate
      .Fields.Append "Almacen", adVarChar, 50
      .Fields.Append "fechaModificacion", adDate
      .Fields.Append "idTipoInventario", adInteger
      .Fields.Append "fechaCreacion", adDate
      .LockType = adLockOptimistic
      .Open
  End With
  Dim oRsTmp1 As New Recordset
  Set oRsTmp1 = mo_ReglasFarmacia.farmInventarioParaExportar
  If oRsTmp1.RecordCount > 0 Then
     oRsTmp1.MoveFirst
     Do While Not oRsTmp1.EOF
        oRsInventarios.AddNew
        oRsInventarios!IdInventario = oRsTmp1!IdInventario
        oRsInventarios!CodigoSismed = Left(oRsTmp1!CodigoSismed, 9) & oRsTmp1!idTipoSuministro
        oRsInventarios!NumeroInventario = oRsTmp1!NumeroInventario
        oRsInventarios!FechaCierre = CDate(Format(oRsTmp1!FechaCierre, sighentidades.DevuelveFechaSoloFormato_DMY))
        oRsInventarios!Almacen = oRsTmp1!Almacen
        oRsInventarios!fechaModificacion = IIf(IsNull(oRsTmp1!fechaModificacion), oRsTmp1!fechaCreacion, oRsTmp1!fechaModificacion)
        oRsInventarios!idTipoInventario = oRsTmp1!idTipoInventario
        oRsInventarios!fechaCreacion = oRsTmp1!fechaCreacion
        oRsInventarios.Update
        oRsTmp1.MoveNext
     Loop
  End If
  oRsTmp1.Close
  Set oRsTmp1 = Nothing
  Set grdInventarios.DataSource = oRsInventarios
  mo_Apariencia.ConfigurarFilasBiColores Me.grdInventarios, sighentidades.GrillaConFilasBicolor


End Sub

Private Sub Form_Load()
  Set grdAuditoriasRS.DataSource = mo_CabeceraReportes.AuditoriaPorTabla("FarmSaldo")
  mo_Apariencia.ConfigurarFilasBiColores Me.grdAuditoriasRS, sighentidades.GrillaConFilasBicolor

  'franklin 2017
  If lcBuscaParametro.SeleccionaFilaParametro(282) <> "S" Then  'Es un Hospital, solo agrega nuevos medicamentos
     chkSoloAgregaItems.Value = 1
  Else
     chkSoloExiste1.Value = 1
  End If
  '
  Set mo_cmbAlmacen.MiComboBox = cmbAlmacen
  mo_cmbAlmacen.BoundColumn = "IdAlmacen"
  mo_cmbAlmacen.ListField = "Descripcion"
  Set mo_cmbAlmacen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarTodosMenosExternos
  '
  mo_ReglasComunes.LlenaListBoxConTablaMensajesEnVentana cmbConsideraciones, "HerrRegeneraSaldos"
  ldUltimaDiaMesActual = sighentidades.DevuelveFechaHoraFinalDelMesDelMovimiento(Now)
  '
  Me.Año.Text = Year(ldUltimaDiaMesActual)
  Me.Mes.Text = Right("0" & Trim(str(Month(ldUltimaDiaMesActual))), 2)
  '
  txtMes.Text = Month(ldUltimaDiaMesActual)
  Me.txtAnio.Text = Year(ldUltimaDiaMesActual)
  If Val(txtMes.Text) - 1 = 0 Then
     txtMes.Text = "12"
     Me.txtAnio.Text = Trim(str(Val(Me.txtAnio.Text) - 1))
  Else
     txtMes.Text = Trim(str(Val(txtMes.Text) - 1))
  End If
  '
  Me.txtRutaGalenhos.Text = App.Path & "\archivos"
  '
  With oRsFarmSaldoDetallado
      .Fields.Append "IdProducto", adInteger
      .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
      .Fields.Append "nombre", adVarChar, 150, adFldIsNullable
      .Fields.Append "Cantidad", adInteger
      .Fields.Append "Lote", adChar, 15
      .Fields.Append "FechaVencimiento", adDate, , adFldIsNullable
      .Fields.Append "FechaVencimientoN", adDate, , adFldIsNullable
      .Fields.Append "Almacen", adVarChar, 150, adFldIsNullable
      .LockType = adLockOptimistic
      .Open
  End With
  Set grdFarmSaldoDetallado.DataSource = oRsFarmSaldoDetallado
  mo_Apariencia.ConfigurarFilasBiColores Me.grdFarmSaldoDetallado, sighentidades.GrillaConFilasBicolor
  '
  Dim oCrypKey As New CrypKey.Util
  txtZipClave.Text = oCrypKey.DecryptString(lcBuscaParametro.SeleccionaFilaParametro(350))
  Set oCrypKey = Nothing
  '
  CreaTemporalYcargaIventarios
  ldHoy = lcBuscaParametro.RetornaFechaHoraServidorSQL
  
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


Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set lcBuscaParametro = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_cmbAlmacen = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub



Function FarmActualizaSaldosMensualRS(lcEntradaOsalida As String, lnIdAlmacen As Long, lnIdProducto As Long, ldFechaMovimiento As Date, lnCantidad As Long, lcLote As String, ldFechaVencimiento As Date, lnIdTipoSalidaBienInsumo As Long, lnPrecio As Double) As Boolean
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim ldFechaMov As Date
    FarmActualizaSaldosMensualRS = False
    ldFechaMov = sighentidades.DevuelveFechaHoraFinalDelMesDelMovimiento(ldFechaMovimiento)
    '
    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = oConexion
       .CommandText = "FarmActualizaSaldosMensualRS"
       Set oParameter = .CreateParameter("@lcEntradaSalida", adVarChar, adParamInput, 1, lcEntradaOsalida): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, lnCantidad): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@FechaMov", adDate, adParamInput, 10, ldFechaMov): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Lote", adVarChar, adParamInput, 15, lcLote): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@FechaVencimiento", adDate, adParamInput, 10, ldFechaVencimiento): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Precio", adDouble, adParamInput, 0, lnPrecio): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@idTipoSalidaBienInsumo", adInteger, adParamInput, 0, lnIdTipoSalidaBienInsumo): .Parameters.Append oParameter
       .Execute
    End With
    FarmActualizaSaldosMensualRS = True
    ms_MensajeError = ""
    Exit Function
ManejadorDeError:
       ms_MensajeError = Err.Number & " " + Err.Description
End Function



Function ActualizaSaldosPorProducto(lcEntradaOsalida As String, lnIdAlmacen As Long, lnIdProducto As Long, lcLote As String, ldFechaVencimiento As Date, lnIdTipoSalidaBienInsumo As Long, lnCantidad As Long, lnPrecio As Double, ldFechaMovimiento As Date, lbActualizaSaldosMensuales As Boolean) As Boolean
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    
    ActualizaSaldosPorProducto = False
    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = oConexion
       .CommandText = "FarmActualizaSaldosPorProducto"
       Set oParameter = .CreateParameter("@lcEntradaSalida", adVarChar, adParamInput, 1, lcEntradaOsalida): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Lote", adVarChar, adParamInput, 15, lcLote): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@FechaVencimiento", adDate, adParamInput, 10, ldFechaVencimiento): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, lnCantidad): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Precio", adDouble, adParamInput, 0, lnPrecio): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@idTipoSalidaBienInsumo", adInteger, adParamInput, 0, lnIdTipoSalidaBienInsumo): .Parameters.Append oParameter
       .Execute
    End With
    If lbActualizaSaldosMensuales = True Then
        If FarmActualizaSaldosMensualRS(lcEntradaOsalida, lnIdAlmacen, lnIdProducto, ldFechaMovimiento, lnCantidad, lcLote, ldFechaVencimiento, lnIdTipoSalidaBienInsumo, lnPrecio) = True Then
           ms_MensajeError = ""
           ActualizaSaldosPorProducto = True
        End If
    Else
        ActualizaSaldosPorProducto = True
    End If
    Exit Function
ManejadorDeError:
       ms_MensajeError = Err.Number & " " + Err.Description

End Function



Sub ActualizaPreciosParaFarmacia()
    If MsgBox("Esta seguro que desea ACTUALIZAR PRECIOS ?", vbQuestion + vbYesNo, "Farmacia") = vbYes Then
       On Error GoTo ErrActPRe
       Dim lcDias1 As String
       lcDias1 = ""
       Me.MousePointer = 11
       lcDias = "1"
       Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
            
            
            lcDias = "2"
            Dim oConexionFox As New ADODB.Connection
            Dim oRsFox1 As New Recordset
            Dim oRsFox2 As New Recordset
            Dim oRsTmp1 As New Recordset
            Dim oRsTmp2 As New Recordset
            Dim mrs_Precios As New Recordset
            Dim mo_CatalogoBienesInsumos As New DOCatalogoBienesInsumos
            Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
            Dim lnIdProducto As Long
            Dim lcSql As String
            Dim lnPrecioVenta As Double, lnPrecioCompra As Double, lnPrecioDistribucion As Double, lcNombreItem As String
            '
            oConexionFox.CommandTimeout = 300
            oConexionFox.Open "DSN=his"
            '
            
            oConexion.CommandTimeout = 300
            oConexion.CursorLocation = adUseClient
            oConexion.Open sighentidades.CadenaConexion
            
            '
            If chkSoloExiste1.Value = 1 Then
                'Debe ser un Hospital, que genero el archivo ZIP desde el SISMEDV2, luego migrará a Galenhos (debe tener 1 archivo)
                Set oRsFox1 = mo_ReglasFacturacion.Sismedv2SeleccionarPrecios(oConexionFox, "pr", "")
                If oRsFox1.RecordCount > 0 Then
                   lcDias = "3"
                   'crea temporal y llena de tipos finanacimientos
                    With mrs_Precios
                          .Fields.Append "IdTipoFinanciamiento", adInteger, 4, adFldIsNullable
                          .Fields.Append "TipoFinanciamiento", adVarChar, 50, adFldIsNullable
                          .Fields.Append "PrecioVenta", adDouble
                          .Fields.Append "PrecioDistribucion", adDouble
                          .Fields.Append "PrecioCompra", adDouble
                          .Fields.Append "PrecioDonacion", adDouble
                          .CursorType = adOpenKeyset
                          .LockType = adLockOptimistic
                          .Open
                    End With
                    Set oRsTmp1 = mo_ReglasFacturacion.TiposFinanciamientoSeleccionarSoloIngresaPrecios
                    If oRsTmp1.RecordCount > 0 Then
                       oRsTmp1.MoveFirst
                       Do While Not oRsTmp1.EOF
                          mrs_Precios.AddNew
                          mrs_Precios.Fields!IdTipoFinanciamiento = oRsTmp1.Fields!IdTipoFinanciamiento
                          mrs_Precios.Fields!TipoFinanciamiento = oRsTmp1.Fields!Descripcion
                          mrs_Precios.Fields!PrecioVenta = 0
                          mrs_Precios.Fields!PrecioCompra = 0
                          mrs_Precios.Fields!PrecioDistribucion = 0
                          mrs_Precios.Fields!PrecioDonacion = 0
                          mrs_Precios.Update
                          oRsTmp1.MoveNext
                       Loop
                   End If
                   oRsTmp1.Close
                   
                   oRsFox1.MoveFirst
                   'lcDias = "5"
                   'lnPrecioVenta = oRsFox1.Fields!Preope
                   'lcDias = "6"
                   'lnPrecioCompra = oRsFox1.Fields!Preadj
                   'lcDias = "7"
                   'lnPrecioDistribucion = oRsFox1.Fields!Predist
                   'lcDias = "8"
                   'lcNombreItem = Left(Trim(oRsFox1.Fields!descr), 290)
                   'lcDias = "9"
                   Do While Not oRsFox1.EOF
If Val(oRsFox1.Fields!medcod) = 808 Then
lcSql = ""
End If
                     lcDias = "5"
                     lnPrecioVenta = oRsFox1.Fields!Preope
                     lcDias = "6"
                     lnPrecioCompra = oRsFox1.Fields!Preadj
                    lcDias = "7"
                     lnPrecioDistribucion = oRsFox1.Fields!Predist
                     lcDias = "8"
                     lcNombreItem = Left(Trim(oRsFox1.Fields!descr), 290)
                     lcDias = "9"

                      mrs_Precios.MoveFirst
                      Do While Not mrs_Precios.EOF
                          mrs_Precios.Fields!PrecioVenta = lnPrecioVenta
                          mrs_Precios.Fields!PrecioCompra = lnPrecioCompra
                          mrs_Precios.Fields!PrecioDistribucion = lnPrecioDistribucion
                          mrs_Precios.Fields!PrecioDonacion = 0
                          mrs_Precios.Update
                          mrs_Precios.MoveNext
                      Loop
                      lcDias = "6"
                      '
                      Set oRsTmp1 = mo_ReglasComunes.CatalogoBienesInsumosSeleccionarPorCodigo(Trim(oRsFox1.Fields!medcod))
                      If oRsTmp1.RecordCount = 0 Then
                            'NO EXISTE Medicamento se crea
                            With mo_CatalogoBienesInsumos
                               .codigo = Trim(oRsFox1.Fields!medcod)
                               .nombre = lcNombreItem
                               '.PrecioUnitario = CCur(Me.txtPrecioUnitario.Text)
                               .NombreComercial = ""
                               '.IdClasificacionBienInsumo = Val(mo_cmbIdClasificacionBienInsumo.BoundText)
                               .IdGrupoFarmacologico = 999
                               .IdSubGrupoFarmacologico = 999
                               .IdPartida = 1
                               .IdCentroCosto = 999
                               If chkSoloAgregaItems.Value = 0 Then
                                    .PrecioCompra = lnPrecioCompra
                                    .PrecioDistribucion = lnPrecioDistribucion
                                    .PrecioDonacion = 0
                               End If
                               .idTipoSalidaBienInsumo = 1
                               .TipoProducto = IIf(oRsFox1.Fields!Medtip = "M", 0, 1)
                               .IdUsuarioAuditoria = ml_idUsuario
                               '.Denominacion = Left(oRsFox1.Fields!medNom, 100)
                               '.Concentracion = Left(oRsFox1.Fields!medCnc, 100)
                               '.Presentacion = Left(oRsFox1.Fields!medPres, 100)
                               '.FormaFarmaceutica = Left(oRsFox1.Fields!medff, 10)
                             End With
                             If mo_AdminComun.CatalogoBienesInsumosAgregar(mo_CatalogoBienesInsumos, 500 + 15, mo_lcNombrePc, mo_CatalogoBienesInsumos.codigo & " " & mo_CatalogoBienesInsumos.nombre) = False Then
                                MsgBox "Error al Crear Medicamento/Insumo: " & mo_CatalogoBienesInsumos.codigo & " " & mo_CatalogoBienesInsumos.nombre & Chr(13) & mo_AdminComun.MensajeError
                                Exit Sub
                             End If
                             lnIdProducto = mo_CatalogoBienesInsumos.idProducto
                      Else
                             lnIdProducto = oRsTmp1.Fields!idProducto
                             If chkSoloAgregaItems.Value = 0 Then
                                mo_AdminComun.CatalogoBienesInsumosActualizaPRECIOS lnPrecioCompra, lnPrecioDistribucion, _
                                                                                    lnIdProducto, oConexion
                             End If
                      End If
                      oRsTmp1.Close
                      'actualiza Precios
                      If chkSoloAgregaItems.Value = 0 Then
                         mo_AdminComun.ActualizaPreciosParaFarmacia mrs_Precios, lnIdProducto
                      End If
                      '
                      oRsFox1.MoveNext
                   Loop
                End If
            Else
                'Debe ser un CS, que generó ZIP desde la DISA los nuevos precios (debe tener 2 archivos)
                Set oRsFox1 = mo_ReglasFacturacion.Sismedv2SeleccionarPrecios(oConexionFox, "nueprec", "")
                If oRsFox1.RecordCount > 0 Then
                   lcDias = "3"
                   'crea temporal y llena de tipos finanacimientos
                    With mrs_Precios
                          .Fields.Append "IdTipoFinanciamiento", adInteger, 4, adFldIsNullable
                          .Fields.Append "TipoFinanciamiento", adVarChar, 50, adFldIsNullable
                          .Fields.Append "PrecioVenta", adDouble
                          .Fields.Append "PrecioDistribucion", adDouble
                          .Fields.Append "PrecioCompra", adDouble
                          .Fields.Append "PrecioDonacion", adDouble
                          .CursorType = adOpenKeyset
                          .LockType = adLockOptimistic
                          .Open
                    End With
                    Set oRsTmp1 = mo_ReglasFacturacion.TiposFinanciamientoSeleccionarSoloIngresaPrecios
                    If oRsTmp1.RecordCount > 0 Then
                       oRsTmp1.MoveFirst
                       Do While Not oRsTmp1.EOF
                          mrs_Precios.AddNew
                          mrs_Precios.Fields!IdTipoFinanciamiento = oRsTmp1.Fields!IdTipoFinanciamiento
                          mrs_Precios.Fields!TipoFinanciamiento = oRsTmp1.Fields!Descripcion
                          mrs_Precios.Fields!PrecioVenta = 0
                          mrs_Precios.Fields!PrecioCompra = 0
                          mrs_Precios.Fields!PrecioDistribucion = 0
                          mrs_Precios.Fields!PrecioDonacion = 0
                          mrs_Precios.Update
                          oRsTmp1.MoveNext
                       Loop
                   End If
                   oRsTmp1.Close
                   'Barre archivo FOX con nuevos precios
                   oRsFox1.MoveFirst
                   Do While Not oRsFox1.EOF
                      mrs_Precios.MoveFirst
                      Do While Not mrs_Precios.EOF
                          mrs_Precios.Fields!PrecioVenta = oRsFox1.Fields!TnNuePop
                          mrs_Precios.Fields!PrecioCompra = oRsFox1.Fields!TnNuePre
                          mrs_Precios.Fields!PrecioDistribucion = oRsFox1.Fields!tnNuePds
                          mrs_Precios.Fields!PrecioDonacion = 0
                          mrs_Precios.Update
                          mrs_Precios.MoveNext
                      Loop
                      '
                      Set oRsTmp1 = mo_ReglasComunes.CatalogoBienesInsumosSeleccionarPorCodigo(Trim(oRsFox1.Fields!medcod))
                      If oRsTmp1.RecordCount = 0 Then
                         'NO EXISTE Medicamento se crea
                         lcDias = "4"
                         Set oRsFox2 = mo_ReglasFacturacion.Sismedv2SeleccionarPrecios(oConexionFox, "pr", " where medcod='" & Trim(oRsFox1.Fields!medcod) & "'")
                         lcDias = "3"
                         If oRsFox2.RecordCount > 0 Then
                            With mo_CatalogoBienesInsumos
                               .codigo = Trim(oRsFox2.Fields!medcod)
                               .nombre = Left(Trim(oRsFox2.Fields!MedNom) & " " & Trim(oRsFox2.Fields!MedPres) & " " & Trim(oRsFox2.Fields!MedCnc), 290) & " " & Trim(oRsFox2.Fields!Medff)
                               '.PrecioUnitario = CCur(Me.txtPrecioUnitario.Text)
                               .NombreComercial = ""
                               '.IdClasificacionBienInsumo = Val(mo_cmbIdClasificacionBienInsumo.BoundText)
                               .IdGrupoFarmacologico = 999
                               .IdSubGrupoFarmacologico = 999
                               .IdPartida = 1
                               .IdCentroCosto = 999
                               If chkSoloAgregaItems.Value = 0 Then
                                    .PrecioCompra = oRsFox1.Fields!TnNuePre
                                    .PrecioDistribucion = oRsFox1.Fields!tnNuePds
                                    .PrecioDonacion = oRsFox2.Fields!preOpe_do
                               End If
                               .idTipoSalidaBienInsumo = 1
                               .TipoProducto = IIf(oRsFox2.Fields!Medtip = "M", 0, 1)
                               .IdUsuarioAuditoria = ml_idUsuario
                               .denominacion = Left(oRsFox2.Fields!MedNom, 100)
                               .Concentracion = Left(oRsFox2.Fields!MedCnc, 100)
                               .Presentacion = Left(oRsFox2.Fields!MedPres, 100)
                               .FormaFarmaceutica = Left(oRsFox2.Fields!Medff, 10)
                             End With

                             If mo_AdminComun.CatalogoBienesInsumosAgregar(mo_CatalogoBienesInsumos, 500 + 15, mo_lcNombrePc, mo_CatalogoBienesInsumos.codigo & " " & mo_CatalogoBienesInsumos.nombre) = False Then
                                MsgBox "Error al Crear Medicamento/Insumo: " & mo_CatalogoBienesInsumos.codigo & " " & mo_CatalogoBienesInsumos.nombre & Chr(13) & mo_AdminComun.MensajeError
                                Exit Sub
                             End If
                             lnIdProducto = mo_CatalogoBienesInsumos.idProducto
                          Else
                             lnIdProducto = 0
                          End If
                          oRsFox2.Close
                      Else
                          lnIdProducto = oRsTmp1.Fields!idProducto
                          If chkSoloAgregaItems.Value = 0 Then
                             mo_ReglasComunes.CatalogoBienesInsumosActualizaPRECIOS oRsFox1.Fields!TnNuePre, oRsFox1.Fields!tnNuePds, lnIdProducto, oConexion
                          End If
                      End If
                      oRsTmp1.Close
                      'actualiza Precios
                      If chkSoloAgregaItems.Value = 0 Then
                         mo_AdminComun.ActualizaPreciosParaFarmacia mrs_Precios, lnIdProducto
                      End If
                      '
                      oRsFox1.MoveNext
                   Loop
                Else
                    oRsFox1.Close
                    Set oRsFox1 = mo_ReglasFacturacion.Sismedv2SeleccionarPrecios(oConexionFox, "pr", " ")
                    If oRsFox1.RecordCount > 0 Then
                       lcDias = "3"
                       'crea temporal y llena de tipos finanacimientos
                        With mrs_Precios
                              .Fields.Append "IdTipoFinanciamiento", adInteger, 4, adFldIsNullable
                              .Fields.Append "TipoFinanciamiento", adVarChar, 50, adFldIsNullable
                              .Fields.Append "PrecioVenta", adDouble
                              .Fields.Append "PrecioDistribucion", adDouble
                              .Fields.Append "PrecioCompra", adDouble
                              .Fields.Append "PrecioDonacion", adDouble
                              .CursorType = adOpenKeyset
                              .LockType = adLockOptimistic
                              .Open
                        End With
                        Set oRsTmp1 = mo_ReglasFacturacion.TiposFinanciamientoSeleccionarSoloIngresaPrecios
                        If oRsTmp1.RecordCount > 0 Then
                           oRsTmp1.MoveFirst
                           Do While Not oRsTmp1.EOF
                              mrs_Precios.AddNew
                              mrs_Precios.Fields!IdTipoFinanciamiento = oRsTmp1.Fields!IdTipoFinanciamiento
                              mrs_Precios.Fields!TipoFinanciamiento = oRsTmp1.Fields!Descripcion
                              mrs_Precios.Fields!PrecioVenta = 0
                              mrs_Precios.Fields!PrecioCompra = 0
                              mrs_Precios.Fields!PrecioDistribucion = 0
                              mrs_Precios.Fields!PrecioDonacion = 0
                              mrs_Precios.Update
                              oRsTmp1.MoveNext
                           Loop
                       End If
                       oRsTmp1.Close
                       oRsFox1.MoveFirst
                       Do While Not oRsFox1.EOF
                          mrs_Precios.MoveFirst
                          Do While Not mrs_Precios.EOF
                              mrs_Precios.Fields!PrecioVenta = oRsFox1.Fields!PrdPreOpe
                              mrs_Precios.Fields!PrecioCompra = oRsFox1.Fields!prdPreadq
                              mrs_Precios.Fields!PrecioDistribucion = oRsFox1.Fields!prdPredist
                              mrs_Precios.Fields!PrecioDonacion = 0
                              mrs_Precios.Update
                              mrs_Precios.MoveNext
                          Loop
                          '
                          Set oRsTmp1 = mo_ReglasComunes.CatalogoBienesInsumosSeleccionarPorCodigo(Trim(oRsFox1.Fields!medcod))
                          If oRsTmp1.RecordCount = 0 Then
                                'NO EXISTE Medicamento se crea
                                With mo_CatalogoBienesInsumos
                                   .codigo = Trim(oRsFox1.Fields!medcod)
                                   .nombre = Left(Trim(oRsFox1.Fields!MedNom), 290)
                                   '.PrecioUnitario = CCur(Me.txtPrecioUnitario.Text)
                                   .NombreComercial = ""
                                   '.IdClasificacionBienInsumo = Val(mo_cmbIdClasificacionBienInsumo.BoundText)
                                   .IdGrupoFarmacologico = 999
                                   .IdSubGrupoFarmacologico = 999
                                   .IdPartida = 1
                                   .IdCentroCosto = 999
                                   If chkSoloAgregaItems.Value = 0 Then
                                        .PrecioCompra = oRsFox1.Fields!prdPreadq
                                        .PrecioDistribucion = oRsFox1.Fields!prdPredist
                                        .PrecioDonacion = 0
                                   End If
                                   .idTipoSalidaBienInsumo = 1
                                   .TipoProducto = IIf(oRsFox1.Fields!Medtip = "M", 0, 1)
                                   .IdUsuarioAuditoria = ml_idUsuario
                                   '.Denominacion = Left(oRsFox1.Fields!medNom, 100)
                                   '.Concentracion = Left(oRsFox1.Fields!medCnc, 100)
                                   '.Presentacion = Left(oRsFox1.Fields!medPres, 100)
                                   '.FormaFarmaceutica = Left(oRsFox1.Fields!medff, 10)
                                 End With
                                 If mo_AdminComun.CatalogoBienesInsumosAgregar(mo_CatalogoBienesInsumos, 500 + 15, mo_lcNombrePc, mo_CatalogoBienesInsumos.codigo & " " & mo_CatalogoBienesInsumos.nombre) = False Then
                                    MsgBox "Error al Crear Medicamento/Insumo: " & mo_CatalogoBienesInsumos.codigo & " " & mo_CatalogoBienesInsumos.nombre & Chr(13) & mo_AdminComun.MensajeError
                                    Exit Sub
                                 End If
                                 lnIdProducto = mo_CatalogoBienesInsumos.idProducto
                          Else
                                 lnIdProducto = oRsTmp1.Fields!idProducto
                                 If chkSoloAgregaItems.Value = 0 Then
                                    mo_AdminComun.CatalogoBienesInsumosActualizaPRECIOS oRsFox1.Fields!prdPreadq, oRsFox1.Fields!prdPredist, lnIdProducto, oConexion
                                 End If
                          End If
                          oRsTmp1.Close
                          'actualiza Precios
                          If chkSoloAgregaItems.Value = 0 Then
                             mo_AdminComun.ActualizaPreciosParaFarmacia mrs_Precios, lnIdProducto
                          End If
                          '
                          oRsFox1.MoveNext
                       Loop
                    End If
                End If
            End If
            oRsFox1.Close
            oConexionFox.Close
            '
            Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "M", 0, "FactCatalogoBienesInsumos", oConexion, 500 + 15, mo_lcNombrePc, "Actualizó Precios para Farmacia")   '500+ ListBarReporte.idReporte
            
            '
            Set oRsFox1 = Nothing
            Set oConexionFox = Nothing
            Me.Visible = False
            LimpiarVariablesDeMemoria
          
    End If
    Unload Me
    Exit Sub
ErrActPRe:

    Select Case lcDias
    Case "1"
          MsgBox Err.Description & Chr(13) & Chr(13) & "Problemas al momento de descromprir", vbInformation, Me.Caption
    Case "2"
          MsgBox Err.Description & Chr(13) & Chr(13) & "Ya descomprimió, pero hay problemas con el ODBC y/o la ruta", vbInformation, Me.Caption
    Case "3"
          MsgBox Err.Description & Chr(13) & Chr(13) & "Ya descomprimió, ya abrió DBF,pero hay problemas", vbInformation, Me.Caption
    Case "4"
          MsgBox Err.Description & Chr(13) & Chr(13) & "Ya descomprimió, ya abrió DBF,pero hay problemas con el archivo PR?????.DBF renombrelo como PR.DBF", vbInformation, Me.Caption
    Case "5"
         If Err.Number = 3265 Then
            lnPrecioVenta = oRsFox1.Fields!PrdPreOpe
            Resume Next
         End If
    Case "6"
         If Err.Number = 3265 Then
            lnPrecioCompra = oRsFox1.Fields!prdPreadq
            Resume Next
         End If
    Case "7"
         If Err.Number = 3265 Then
            lnPrecioDistribucion = oRsFox1.Fields!prdPredist
            Resume Next
         End If
    Case "8"
         If Err.Number = 3265 Then
            lcNombreItem = Left(Trim(oRsFox1.Fields!MedNom) & " " & Trim(oRsFox1.Fields!MedNom) & _
                                             " " & Trim(oRsFox1.Fields!MedPres) & " " & Trim(oRsFox1.Fields!MedCnc) & " " & _
                                             Trim(oRsFox1.Fields!Medff), 290)
            Resume Next
         End If
    End Select
    Exit Sub
    Me.MousePointer = 1
    Resume
End Sub



Private Sub grdAuditoriasRS_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdAuditoriasRS.Bands(0).Columns("observaciones").Width = 4000
End Sub

Private Sub grdFarmSaldoDetallado_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti

    grdFarmSaldoDetallado.Bands(0).Columns("idProducto").Hidden = True
    '
    grdFarmSaldoDetallado.Bands(0).Columns("Codigo").Header.Caption = "Codigo"
    grdFarmSaldoDetallado.Bands(0).Columns("Codigo").Width = 750
    grdFarmSaldoDetallado.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
    '
    grdFarmSaldoDetallado.Bands(0).Columns("nombre").Header.Caption = "Medicamento/Insumo"
    grdFarmSaldoDetallado.Bands(0).Columns("nombre").Width = 6000
    grdFarmSaldoDetallado.Bands(0).Columns("nombre").Activation = ssActivationActivateNoEdit
    '
    grdFarmSaldoDetallado.Bands(0).Columns("Cantidad").Header.Caption = "Cantidad"
    grdFarmSaldoDetallado.Bands(0).Columns("Cantidad").Width = 750
    grdFarmSaldoDetallado.Bands(0).Columns("Cantidad").Activation = ssActivationActivateNoEdit
    '
    grdFarmSaldoDetallado.Bands(0).Columns("Lote").Header.Caption = "Lote"
    grdFarmSaldoDetallado.Bands(0).Columns("Lote").Width = 1000
    grdFarmSaldoDetallado.Bands(0).Columns("Lote").Activation = ssActivationActivateNoEdit
    '
    grdFarmSaldoDetallado.Bands(0).Columns("FechaVencimiento").Header.Caption = "FVencimiento"
    grdFarmSaldoDetallado.Bands(0).Columns("FechaVencimiento").Width = 1000
    grdFarmSaldoDetallado.Bands(0).Columns("FechaVencimiento").Activation = ssActivationActivateNoEdit
    '
    grdFarmSaldoDetallado.Bands(0).Columns("FechaVencimientoN").Header.Caption = "Nueva F.Venc"
    grdFarmSaldoDetallado.Bands(0).Columns("FechaVencimientoN").Width = 1000
    grdFarmSaldoDetallado.Bands(0).Columns("FechaVencimientoN").Activation = ssActivationAllowEdit
    '
    grdFarmSaldoDetallado.Bands(0).Columns("Almacen").Header.Caption = "Almacen"
    grdFarmSaldoDetallado.Bands(0).Columns("Almacen").Width = 3000
    grdFarmSaldoDetallado.Bands(0).Columns("Almacen").Activation = ssActivationActivateNoEdit

End Sub



Private Sub grdInventarios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdInventarios.Bands(0).Columns("idInventario").Hidden = True
    grdInventarios.Bands(0).Columns("Exportar").Width = 1000
    grdInventarios.Bands(0).Columns("Exportar").Header.Appearance.ForeColor = vbWhite
    grdInventarios.Bands(0).Columns("Exportar").Header.Appearance.BackColor = vbRed
    grdInventarios.Bands(0).Columns("CodigoSismed").Width = 1200
    grdInventarios.Bands(0).Columns("CodigoSismed").Header.Appearance.ForeColor = vbWhite
    grdInventarios.Bands(0).Columns("CodigoSismed").Header.Appearance.BackColor = vbRed
    grdInventarios.Bands(0).Columns("numeroInventario").Header.Caption = "N°Inventario"
    grdInventarios.Bands(0).Columns("numeroInventario").Width = 1200
    grdInventarios.Bands(0).Columns("numeroInventario").Header.Appearance.ForeColor = vbWhite
    grdInventarios.Bands(0).Columns("numeroInventario").Header.Appearance.BackColor = vbRed
    grdInventarios.Bands(0).Columns("FechaCierre").Header.Caption = "F.Cierre Inventario"
    grdInventarios.Bands(0).Columns("FechaCierre").Width = 1200
    grdInventarios.Bands(0).Columns("FechaCierre").Activation = ssActivationActivateNoEdit
    grdInventarios.Bands(0).Columns("FechaCierre").Header.Appearance.ForeColor = vbWhite
    grdInventarios.Bands(0).Columns("FechaCierre").Header.Appearance.BackColor = vbRed
    grdInventarios.Bands(0).Columns("Almacen").Width = 5000
    grdInventarios.Bands(0).Columns("Almacen").Activation = ssActivationActivateNoEdit
End Sub

Private Sub txtCodigoSismed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCodigoSismed.Text <> "" Then
       Dim oRsTmp As New ADODB.Recordset
       Set oRsTmp = mo_ReglasFarmacia.FarmaciaMuestraSaldosPorCodigo(txtCodigoSismed.Text)
       If oRsFarmSaldoDetallado.RecordCount > 0 Then
          oRsFarmSaldoDetallado.MoveFirst
          Do While Not oRsFarmSaldoDetallado.EOF
             oRsFarmSaldoDetallado.Delete
             oRsFarmSaldoDetallado.Update
             oRsFarmSaldoDetallado.MoveNext
          Loop
       End If
       If oRsTmp.RecordCount > 0 Then
          oRsTmp.MoveFirst
          Do While Not oRsTmp.EOF
             oRsFarmSaldoDetallado.AddNew
             oRsFarmSaldoDetallado.Fields!idProducto = oRsTmp.Fields!idProducto
             oRsFarmSaldoDetallado.Fields!codigo = oRsTmp.Fields!codigo
             oRsFarmSaldoDetallado.Fields!nombre = oRsTmp.Fields!nombre
             oRsFarmSaldoDetallado.Fields!Cantidad = oRsTmp.Fields!Cantidad
             oRsFarmSaldoDetallado.Fields!Lote = oRsTmp.Fields!Lote
             oRsFarmSaldoDetallado.Fields!fechaVencimiento = oRsTmp.Fields!fechaVencimiento
             oRsFarmSaldoDetallado.Fields!Almacen = oRsTmp.Fields!Almacen
             oRsFarmSaldoDetallado.Update
             oRsTmp.MoveNext
          Loop
          Set grdFarmSaldoDetallado.DataSource = oRsFarmSaldoDetallado
       Else
          MsgBox "No hay datos para ese CODIGO"
          Set grdFarmSaldoDetallado.DataSource = Nothing
       End If
    End If
End Sub



Sub ExportaPreciosSismed()
    Dim oRsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
    Dim oConexion As New ADODB.Connection
    Dim oConexionFox As New ADODB.Connection
    Dim oCrypKey As New CrypKey.Util
    Dim oCommand As New ADODB.Command
    Dim oParameter As New ADODB.Command
    Dim lnContadorDetalle As Integer
    Dim lnTotal As Integer
    Dim lcRutaExportar As String
    Dim lcTempo As Object
    Dim oRsFox As New Recordset
    Dim lcNombre_Lote As String
    Dim lcSql As String
    Dim lnNroReg As Integer
    Dim NombreArchivo As String
    Dim NombreTablaExportar As String
    Dim lcArchivoExpZip As String
    Dim prueba As Integer, lnIdProducto As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim EXPORTAR_RUTA As String
    Dim CODIGOUNIEJEC As String
    Dim CODIGO_DISA As String
    Dim lcFolder As String
    Dim ldFechaMod As Date
    Dim ldReg_sanit As String
    Dim Fechaisnull As Boolean

    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    oConexionFox.CommandTimeout = 300
    oConexionFox.Open "DSN=his"
    
    Set oRsTmp = mo_ReglasFarmacia.ExportaPreciosSismed(oConexion)
    EXPORTAR_RUTA = lcBuscaParametro.SeleccionaFilaParametro(313)
    CODIGOUNIEJEC = lcBuscaParametro.SeleccionaFilaParametro(208)
    CODIGO_DISA = lcBuscaParametro.SeleccionaFilaParametro(239)
    
 'ProgressBarExportaPreciosSismed
    If oRsTmp.RecordCount > 0 Then
        ProgressBarExportaPreciosSismed.Min = 0
        ProgressBarExportaPreciosSismed.Max = oRsTmp.RecordCount
   Else
        ProgressBarExportaPreciosSismed.Min = 0
        ProgressBarExportaPreciosSismed.Max = 1
        ProgressBarExportaPreciosSismed.Value = 1
        Me.Refresh
   End If
    


    If oRsTmp.RecordCount > 0 Then
        Set lcTempo = CreateObject("Scripting.FileSystemObject")
        lcFolder = EXPORTAR_RUTA & "ExportaPreciosSismedTem"
        '
        On Error Resume Next
        lcTempo.CreateFolder lcFolder
        '
    
        'Inicializa tabla fox
            
        lcNombre_Lote = App.Path            '"C:\Archivos de programa\Digital Works Corporation\GalenHos"  '
        NombreArchivo = lcNombre_Lote & "\Archivos\ExpPrec.dbf"
        'If EstadoDeArchivo(NombreArchivo) Then Kill (NombreArchivo)
        NombreTablaExportar = "pr" & Me.Año & Me.Mes
        lcRutaExportar = lcNombre_Lote & "\Archivos\" & NombreTablaExportar & ".dbf"
        FileCopy lcNombre_Lote & "\Archivos\ExpPrec.dbf", lcRutaExportar
        'lcSql = "delete from " & NombreTablaExportar
        'oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        lcSql = "select * from " & NombreTablaExportar
        oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic

        oRsTmp.MoveFirst
        lnNroReg = 1
        Do While Not oRsTmp.EOF
            oRsFox.AddNew
            oRsFox.Fields!Cdisa = CODIGO_DISA
            oRsFox.Fields!Cperiodo = Me.Año & Me.Mes
            oRsFox.Fields!Ccodue = Right(CODIGOUNIEJEC, 5)
            oRsFox.Fields!medcod = oRsTmp.Fields!codigo
            oRsFox.Fields!descr = oRsTmp.Fields!nombre
            oRsFox.Fields!Medtip = IIf(oRsTmp.Fields!TipoProducto = 0, "M", "I")
            oRsFox.Fields!Preadj = oRsTmp.Fields!PrecioCompra
            oRsFox.Fields!Predist = oRsTmp.Fields!PrecioDistribucion
            oRsFox.Fields!Preope = oRsTmp.Fields!PrecioVenta
            lnIdProducto = oRsTmp.Fields!idProducto
            Set oRsTmp1 = mo_ReglasFarmacia.ExportaPreciosSismedRegSant(lnIdProducto, oConexion)
            Set oRsTmp2 = mo_ReglasFarmacia.FarmHistPrecioPorIdProducto(lnIdProducto, oConexion)
            '
            Fechaisnull = True
            ldFechaMod = vbNull
            If oRsTmp2.RecordCount > 0 Then
                If Not IsNull(oRsTmp2!fecha) Then
                    ldFechaMod = oRsTmp2!fecha
                    Fechaisnull = False
                End If
            End If
            If Fechaisnull Then
                If oRsTmp1.RecordCount > 0 Then
                    ldFechaMod = oRsTmp1!fechaCreacion
                End If
            End If
            '
            ldReg_sanit = "S/R"
            If oRsTmp1.RecordCount > 0 Then
               ldReg_sanit = oRsTmp1!registroSanitario
            End If
            '
            oRsFox.Fields!Fechamod = ldFechaMod
            oRsFox.Fields!Reg_sanit = Trim(ldReg_sanit)
            oRsFox.Update
            lnContadorDetalle = lnContadorDetalle + 1
            ProgressBarExportaPreciosSismed.Value = lnContadorDetalle
            Me.Refresh
             
          oRsTmp.MoveNext
       Loop
    End If
    If oRsFox.State = 1 Then oRsFox.Close
    oRsTmp.Close
    oConexionFox.Close
    lcArchivoExpZip = "Catprec_" & Me.Año & Me.Mes & "_" & CODIGOUNIEJEC & "_" & Format(Date, "yyyymmdd") & ".zip"
    fso.CopyFile lcRutaExportar, lcFolder & "\" & NombreTablaExportar & ".dbf"

    sighentidades.ComprimeArchivoZip EXPORTAR_RUTA & lcArchivoExpZip, oCrypKey.DecryptString(lcBuscaParametro.SeleccionaFilaParametro(350)), lcFolder & "\*.*"
    MsgBox "Se creó el archivo: " & EXPORTAR_RUTA & lcArchivoExpZip
    lcTempo.DeleteFolder lcFolder  'elimino carpeta de creacion de textos
    Unload Me
End Sub

Public Function EstadoDeArchivo(ByVal Archivo As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(Archivo)) Then
    EstadoDeArchivo = True
    Else
    EstadoDeArchivo = False
    End If
End Function
