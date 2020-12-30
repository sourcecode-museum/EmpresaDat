VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form frmEmpresa 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Dados da Empresa"
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   Icon            =   "frmEmpresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin rdActiveText.ActiveText txtCampo 
      Height          =   285
      Index           =   3
      Left            =   1380
      TabIndex        =   7
      Top             =   1710
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   503
      Appearance      =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   40
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      Height          =   285
      Index           =   4
      Left            =   1380
      TabIndex        =   9
      Top             =   2025
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   503
      Alignment       =   2
      Appearance      =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   18
      TextMask        =   8
      RawText         =   8
      Mask            =   "##.###.###/####-##"
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      Height          =   285
      Index           =   5
      Left            =   1380
      TabIndex        =   11
      Top             =   2340
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   503
      Appearance      =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   20
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      Height          =   285
      Index           =   6
      Left            =   1380
      TabIndex        =   13
      Top             =   2655
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   503
      Appearance      =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   40
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      Height          =   285
      Index           =   7
      Left            =   1380
      TabIndex        =   15
      Top             =   2970
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   503
      Appearance      =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   40
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      Height          =   285
      Index           =   8
      Left            =   1380
      TabIndex        =   17
      Top             =   3285
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   503
      Appearance      =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   30
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      Height          =   285
      Index           =   9
      Left            =   4245
      TabIndex        =   18
      Top             =   3285
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   503
      Alignment       =   2
      Appearance      =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   2
      TextMask        =   9
      TextCase        =   1
      RawText         =   9
      Mask            =   "??"
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      Height          =   285
      Index           =   10
      Left            =   1380
      TabIndex        =   20
      Top             =   3600
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      Alignment       =   2
      Appearance      =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   9
      TextMask        =   6
      RawText         =   6
      Mask            =   "#####-###"
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      Height          =   285
      Index           =   12
      Left            =   1380
      TabIndex        =   24
      Top             =   4230
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      Alignment       =   2
      Appearance      =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   13
      TextMask        =   9
      RawText         =   9
      Mask            =   "(##)####-####"
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      Height          =   285
      Index           =   11
      Left            =   1380
      TabIndex        =   22
      Top             =   3915
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      Alignment       =   2
      Appearance      =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   13
      TextMask        =   9
      RawText         =   9
      Mask            =   "(##)####-####"
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Sempre no &Topo"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3210
      TabIndex        =   25
      Top             =   4320
      Width           =   1500
   End
   Begin rdActiveText.ActiveText txtCampo 
      Height          =   285
      Index           =   2
      Left            =   1380
      TabIndex        =   5
      Top             =   1395
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   503
      Appearance      =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   40
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      Height          =   285
      Index           =   0
      Left            =   1380
      TabIndex        =   1
      Top             =   660
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   503
      Appearance      =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   10
      TextCase        =   1
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      Height          =   285
      Index           =   1
      Left            =   1380
      TabIndex        =   3
      Top             =   975
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   503
      Appearance      =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   150
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      Locked          =   -1  'True
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Excluir"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   5
      Left            =   1665
      TabIndex        =   33
      Tag             =   "ButtonLabel"
      Top             =   4830
      Width           =   495
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Incluir"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   600
      OLEDropMode     =   1  'Manual
      TabIndex        =   32
      Tag             =   "ButtonLabel"
      Top             =   4830
      Width           =   435
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3045
      TabIndex        =   31
      Top             =   690
      Width           =   1140
   End
   Begin VB.Image imgNav 
      Height          =   240
      Index           =   3
      Left            =   4425
      Picture         =   "frmEmpresa.frx":0442
      Top             =   690
      Width           =   240
   End
   Begin VB.Image imgNav 
      Height          =   240
      Index           =   2
      Left            =   4185
      Picture         =   "frmEmpresa.frx":07B4
      Top             =   690
      Width           =   240
   End
   Begin VB.Image imgNav 
      Height          =   240
      Index           =   1
      Left            =   2790
      Picture         =   "frmEmpresa.frx":0B1B
      Top             =   690
      Width           =   240
   End
   Begin VB.Image imgNav 
      Height          =   240
      Index           =   0
      Left            =   2550
      Picture         =   "frmEmpresa.frx":0E82
      Top             =   690
      Width           =   240
   End
   Begin VB.Label lblPathFile 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   4425
      TabIndex        =   30
      Top             =   1005
      Width           =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   60
      X2              =   4815
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lblCaptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   270
      TabIndex        =   0
      Tag             =   "Label"
      Top             =   720
      Width           =   540
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   4530
      TabIndex        =   29
      Top             =   150
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   285
      Index           =   2
      Left            =   4485
      Top             =   135
      Width           =   330
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   585
      X2              =   2685
      Y1              =   390
      Y2              =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dados da Empresa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   630
      TabIndex        =   28
      Top             =   135
      Width           =   1770
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   105
      Picture         =   "frmEmpresa.frx":11F3
      Top             =   15
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      Height          =   510
      Index           =   0
      Left            =   15
      Top             =   15
      Width           =   4905
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FAX:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   270
      TabIndex        =   23
      Tag             =   "Label"
      Top             =   4290
      Width           =   345
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefone:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   270
      TabIndex        =   21
      Tag             =   "Label"
      Top             =   3975
      Width           =   675
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CEP:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   270
      TabIndex        =   19
      Tag             =   "Label"
      Top             =   3660
      Width           =   360
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade - UF:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   270
      TabIndex        =   16
      Tag             =   "Label"
      Top             =   3345
      Width           =   885
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   270
      TabIndex        =   14
      Tag             =   "Label"
      Top             =   3030
      Width           =   450
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   270
      TabIndex        =   12
      Tag             =   "Label"
      Top             =   2715
      Width           =   735
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insc. Est.:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   270
      TabIndex        =   10
      Tag             =   "Label"
      Top             =   2400
      Width           =   705
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.G.C.:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   270
      TabIndex        =   8
      Tag             =   "Label"
      Top             =   2085
      Width           =   510
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome Fantasia:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   270
      TabIndex        =   6
      Tag             =   "Label"
      Top             =   1770
      Width           =   1110
   End
   Begin VB.Image imgLabel 
      Height          =   270
      Index           =   8
      Left            =   135
      Picture         =   "frmEmpresa.frx":1635
      Top             =   3930
      Width           =   1545
   End
   Begin VB.Image imgLabel 
      Height          =   270
      Index           =   9
      Left            =   135
      Picture         =   "frmEmpresa.frx":2C67
      Top             =   4245
      Width           =   1545
   End
   Begin VB.Image imgLabel 
      Height          =   270
      Index           =   6
      Left            =   135
      Picture         =   "frmEmpresa.frx":4299
      Top             =   3615
      Width           =   1545
   End
   Begin VB.Image imgLabel 
      Height          =   270
      Index           =   5
      Left            =   135
      Picture         =   "frmEmpresa.frx":58CB
      Top             =   3300
      Width           =   1545
   End
   Begin VB.Image imgLabel 
      Height          =   270
      Index           =   4
      Left            =   135
      Picture         =   "frmEmpresa.frx":6EFD
      Top             =   2985
      Width           =   1545
   End
   Begin VB.Image imgLabel 
      Height          =   270
      Index           =   3
      Left            =   135
      Picture         =   "frmEmpresa.frx":852F
      Top             =   2670
      Width           =   1545
   End
   Begin VB.Image imgLabel 
      Height          =   270
      Index           =   2
      Left            =   135
      Picture         =   "frmEmpresa.frx":9B61
      Top             =   2355
      Width           =   1545
   End
   Begin VB.Image imgLabel 
      Height          =   270
      Index           =   1
      Left            =   135
      Picture         =   "frmEmpresa.frx":B193
      Top             =   2040
      Width           =   1545
   End
   Begin VB.Image imgLabel 
      Height          =   270
      Index           =   0
      Left            =   135
      Picture         =   "frmEmpresa.frx":C7C5
      Top             =   1725
      Width           =   1545
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Salvar"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   2715
      OLEDropMode     =   1  'Manual
      TabIndex        =   26
      Tag             =   "ButtonLabel"
      Top             =   4830
      Width           =   555
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   4
      Tag             =   "Label"
      Top             =   1455
      Width           =   465
   End
   Begin VB.Image imgButton 
      Height          =   375
      Index           =   1
      Left            =   3570
      Picture         =   "frmEmpresa.frx":DDF7
      Top             =   3585
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   0
      Left            =   3570
      Picture         =   "frmEmpresa.frx":F4E1
      Stretch         =   -1  'True
      Top             =   3945
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Cancelar"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   3765
      TabIndex        =   27
      Tag             =   "ButtonLabel"
      Top             =   4830
      Width           =   645
   End
   Begin VB.Image imgButton 
      Height          =   375
      Index           =   2
      Left            =   2445
      Picture         =   "frmEmpresa.frx":10F03
      Stretch         =   -1  'True
      Top             =   4740
      Width           =   1095
   End
   Begin VB.Image imgButton 
      Height          =   375
      Index           =   3
      Left            =   3525
      Picture         =   "frmEmpresa.frx":12925
      Stretch         =   -1  'True
      Top             =   4740
      Width           =   1095
   End
   Begin VB.Image imgLabel 
      Height          =   270
      Index           =   7
      Left            =   135
      Picture         =   "frmEmpresa.frx":14347
      Top             =   1410
      Width           =   1545
   End
   Begin VB.Image imgLabel 
      Height          =   270
      Index           =   10
      Left            =   135
      Picture         =   "frmEmpresa.frx":15979
      Top             =   675
      Width           =   1545
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      Height          =   285
      Index           =   3
      Left            =   4380
      Top             =   975
      Width           =   330
   End
   Begin VB.Label lblCaptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caminho DB:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   270
      TabIndex        =   2
      Tag             =   "Label"
      Top             =   1035
      Width           =   930
   End
   Begin VB.Image imgLabel 
      Height          =   270
      Index           =   11
      Left            =   135
      Picture         =   "frmEmpresa.frx":16FAB
      Top             =   990
      Width           =   1545
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   285
      Left            =   2520
      Top             =   660
      Width           =   2190
   End
   Begin VB.Image imgButton 
      Height          =   375
      Index           =   5
      Left            =   1350
      Picture         =   "frmEmpresa.frx":185DD
      Stretch         =   -1  'True
      Top             =   4740
      Width           =   1095
   End
   Begin VB.Image imgButton 
      Height          =   375
      Index           =   4
      Left            =   270
      Picture         =   "frmEmpresa.frx":19FFF
      Stretch         =   -1  'True
      Top             =   4740
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      Height          =   510
      Index           =   1
      Left            =   15
      Top             =   4695
      Width           =   4905
   End
End
Attribute VB_Name = "frmEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************
'Últimas Alterações:  28/01/2002 às 10:25:41 horas
'                     15/01/2003 às 10:52:32 horas
'*********************************************************************
Option Explicit

'For Dragging Borderless Forms...
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

'AlwaysOnTop
Private Declare Function SetWindowPos Lib "user32" (ByVal h&, ByVal hb&, ByVal X&, ByVal Y&, ByVal cx&, ByVal cy&, ByVal f&) As Long
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private sPathDAT  As String

Private WithEvents RS  As ADODB.Recordset
Attribute RS.VB_VarHelpID = -1



Private Sub Form_Load()
  If App.PrevInstance = True Then Unload Me
  
  sPathDAT = App.Path & "\" & "Empresa.dat"
  
  Call AbrirDat   'Abre os Dados apartir de uma arquivo independete empresa.dat
End Sub

Private Sub AbrirDat()
  
  On Error GoTo Destruir
  Set RS = New ADODB.Recordset
  With RS
    .CursorType = adOpenDynamic
    .LockType = adLockBatchOptimistic
    
    .Open sPathDAT, , , , adCmdFile
  End With
  On Error GoTo 0
  
  'Não deu Erro Set a Colecao
  Dim i As Integer
  
  With RS
    If .RecordCount > 0 Then
      For i = 0 To .Fields.Count - 1
        txtCampo(i).DataField = .Fields(i).Name
        Set txtCampo(i).DataSource = RS
      Next
    End If

  End With
  
  Exit Sub
  
Destruir:
  Set RS = Nothing
End Sub

Private Sub SalvarDat()
             
  Set RS = New ADODB.Recordset
  
  With RS.Fields
    .Append "Codigo", adChar, 10
    .Append "PathFile", adChar, 150
    .Append "Nome", adChar, 40
    .Append "Fantasia", adChar, 40
    .Append "CGC", adChar, 18
    .Append "Insc", adChar, 20
    .Append "End", adChar, 40
    .Append "Bairro", adChar, 40
    .Append "Cidade", adChar, 30
    .Append "UF", adChar, 2
    .Append "CEP", adChar, 9
    .Append "Fone", adChar, 13
    .Append "FAX", adChar, 13
  End With
  
  With RS
    .Open
    .AddNew Array("Codigo", "PathFile", "Nome", "Fantasia", "CGC", "Insc", "End", "Bairro", "Cidade", "UF", "CEP", "Fone", "Fax"), _
            Array(txtCampo(0).Text, txtCampo(1).Text, _
                  txtCampo(2).Text, txtCampo(3).Text, _
                  txtCampo(4).Text, txtCampo(5).Text, _
                  txtCampo(6).Text, txtCampo(7).Text, _
                  txtCampo(8).Text, txtCampo(9).Text, _
                  txtCampo(10).Text, txtCampo(11).Text, _
                  txtCampo(12).Text)

    .Save sPathDAT & "~Temp", adPersistADTG
  End With

  FileCopy sPathDAT & "~Temp", sPathDAT
  On Error Resume Next
  Kill sPathDAT & "~Temp"
  On Error GoTo 0
  
  RS.Close
  Set RS = Nothing
End Sub

Private Sub Check1_Click()
  Call AlwaysOnTop(CBool(Check1.Value))
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call DragForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set frmEmpresa = Nothing
End Sub

Private Sub imgButton_Click(Index As Integer)
  Select Case Index
    Case Is = 0, 1 'Imagens
    Case Else
      lblButton_Click Index
  End Select
End Sub

Private Sub imgNav_Click(Index As Integer)
  Dim nP As Long, nC As Long
  
  On Error Resume Next
    nP = RS.AbsolutePosition
    nC = RS.RecordCount
  On Error GoTo 0
  
  Select Case Index
    Case Is = 0
      If nP > 1 And (nC > 1) Then RS.MoveFirst
    Case Is = 1
      If nP > 1 Then RS.MovePrevious
    Case Is = 2
      If nP < nC Then RS.MoveNext
    Case Is = 3
      If nP < nC And (nC > 1) Then RS.MoveLast
  End Select
End Sub

Private Sub lblButton_Click(Index As Integer)
  On Error GoTo TrataErro
  Select Case Index
    Case Is = 2 'Salva
      If Not RS Is Nothing Then
        RS.Save
      Else
        Call SalvarDat  'Salva as informacoes em um arquivo independete Empresa.dat
      End If

    Case Is = 3 'Cancelar
      If Not RS Is Nothing Then RS.Cancel
    Case Is = 4 'Incluir
      If Not RS Is Nothing Then RS.AddNew
    Case Is = 5 'Excluir
      If Not RS Is Nothing Then
        RS.Delete adAffectCurrent
        If RS.RecordCount > 0 Then RS.MoveFirst
      End If
  End Select
  On Error GoTo 0
  Exit Sub
TrataErro:
  MsgBox "Erro: " & Err.Description, vbCritical, "Cadastro de Empresas"
  On Error GoTo 0
End Sub

Private Sub lblButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
    imgButton(Index).Picture = imgButton(1).Picture
    lblButton(Index).ForeColor = vbBlack
  End If
End Sub

Private Sub lblButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgButton(Index).Picture = imgButton(0).Picture
  lblButton(Index).ForeColor = &HC0C0C0
End Sub

Private Sub lblClose_Click()
  Unload Me
End Sub

Private Sub lblPathFile_Click()
  Dim sPath As String
  
  sPath = mBrowseFolders.BrowseFolders(Me.hWnd, "Caminho do Banco de Dados", App.Path)
  If sPath <> "" Then
    txtCampo(1).Text = sPath
  End If
End Sub

Private Sub DragForm()
  On Local Error Resume Next
  'Move the borderless form...
  Call ReleaseCapture
  Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub AlwaysOnTop(iPosition As Boolean)
  Dim lFlag As Long
  
  'On top or not on top...
  If iPosition Then
    lFlag = -1
  Else
    lFlag = -2
  End If
  
  Call SetWindowPos(Me.hWnd, lFlag, Me.Left / Screen.TwipsPerPixelX, _
                                    Me.Top / Screen.TwipsPerPixelY, _
                                    Me.Width / Screen.TwipsPerPixelX, _
                                    Me.Height / Screen.TwipsPerPixelY, FLAGS)
End Sub

Private Sub RS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  Dim nP As Integer
  On Error Resume Next
  nP = pRecordset.AbsolutePosition
  If Err.Number <> 0 Then nP = 1
  On Error GoTo 0
  lblCount.Caption = nP & "/" & pRecordset.RecordCount
End Sub

