VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Relatorio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Transações"
   ClientHeight    =   6585
   ClientLeft      =   3750
   ClientTop       =   3195
   ClientWidth     =   7005
   Icon            =   "frm_Relatorio.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   7005
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1140
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   2011
      ButtonWidth     =   1455
      ButtonHeight    =   1852
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Limpar"
            Key             =   "Limpar"
            Object.ToolTipText     =   "Limpar"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Gerar"
            Key             =   "Gerar"
            Object.ToolTipText     =   "Gerar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Fechar"
            Key             =   "Fechar"
            Object.ToolTipText     =   "Fechar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      OLEDropMode     =   1
   End
   Begin VB.Frame Frame4 
      Caption         =   "Categoria:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   4
      Top             =   3000
      Width           =   5895
      Begin VB.ComboBox cmbCategoria 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frm_Relatorio.frx":1084A
         Left            =   1920
         List            =   "frm_Relatorio.frx":1085D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   5895
      Begin VB.ComboBox cmbStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frm_Relatorio.frx":10899
         Left            =   1920
         List            =   "frm_Relatorio.frx":108A9
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
   End
   Begin Crystal.CrystalReport cr 
      Left            =   6720
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordenar Por : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   1
      Top             =   4680
      Width           =   5895
      Begin VB.ComboBox cmbOrdem 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frm_Relatorio.frx":108E3
         Left            =   1920
         List            =   "frm_Relatorio.frx":108F0
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.PictureBox ImageList1 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   12360
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   3120
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ordem:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   -360
      TabIndex        =   9
      Top             =   4200
      Width           =   1575
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   7080
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Relatorio.frx":1090D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Relatorio.frx":11823
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Relatorio.frx":126B5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Filtros:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   -480
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "frm_Relatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    Me.Caption = "XYZ - Administradora de Cartões de Crédito - " + Me.Caption
    Call fLimparCampos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
On Error GoTo TratamentoDeErro

Select Case Button.Key
    Case "Limpar"
        Call fLimparCampos
     
    Case "Gerar"
            Call fGerarRelatorio
            
    Case "Fechar"
        Unload Me
    
End Select

Exit Sub
TratamentoDeErro:
    ' Monta a mensagem de log com detalhes do erro
    Dim strErroDetails As String
    strErroDetails = "Erro na rotina MinhaRotinaQuePodeGerarErro - " & _
                     "Número: " & Err.Number & " | " & _
                     "Descrição: " & Err.Description & " | " & _
                     "Fonte: " & Err.Source & " | " & _
                     "ÚltimaDLL: " & Err.HelpFile & " | " & _
                     "Contexto: Linha do erro/Estado da aplicação" ' Adicione contexto se possível

    ' Chama a rotina de log do módulo1
    Call EscreverLogErro(strErroDetails)

    ' Opcional: Avisar o usuário de forma amigável (sem mostrar detalhes técnicos)
    MsgBox "Ocorreu um erro inesperado. O problema foi registrado e será investigado.", vbCritical, "Erro"


End Sub


Private Function fLimparCampos()

    cmbStatus.ListIndex = 0
    cmbCategoria.ListIndex = 0
    cmbOrdem.ListIndex = 0
    

End Function
Private Function fGerarRelatorio()

    cr.ReportFileName = xCaminhoRPT & "\RelTransacoes.rpt"
    cr.WindowTitle = "Relatório de Transações"
    cr.Destination = crptToWindow
    cr.WindowState = 2
    
    cr.SelectionFormula = ""
    cr.Formulas(1) = ""
    
 
    'Filtro por Status
        If cmbStatus.ListIndex = "1" Then
           cr.SelectionFormula = "{VW_ConsultaTransacoesPorCategoria.Status}='Aprovada'"
        ElseIf cmbStatus.ListIndex = "2" Then
            cr.SelectionFormula = "{VW_ConsultaTransacoesPorCategoria.Status}='Pendente'"
         ElseIf cmbStatus.ListIndex = "3" Then
            cr.SelectionFormula = "{VW_ConsultaTransacoesPorCategoria.Status}='Cancelada'"
        End If

        cr.Formulas(1) = "Status = 'Filtro por Status: '" + cmbStatus.Text

    'Filtro por Categoria
        If cmbCategoria.ListIndex = "1" Then
           cr.SelectionFormula = "{VW_ConsultaTransacoesPorCategoria.Categoria}='Média'"
        ElseIf cmbCategoria.ListIndex = "2" Then
            cr.SelectionFormula = "{VW_ConsultaTransacoesPorCategoria.Categoria}='Premium'"
        ElseIf cmbCategoria.ListIndex = "3" Then
            cr.SelectionFormula = "{VW_ConsultaTransacoesPorCategoria.Categoria}='Baixa'"
        ElseIf cmbCategoria.ListIndex = "4" Then
            cr.SelectionFormula = "{VW_ConsultaTransacoesPorCategoria.Categoria}='Alta'"
        End If

        cr.Formulas(1) = "Status = 'Filtro por Status: '" + cmbStatus.Text

        
    'Ordem por
    If cmbOrdem.Text = "Data" Then
            cr.SortFields(0) = "+{VW_ConsultaTransacoesPorCategoria.Data_Transacao}"
        ElseIf cmbOrdem.Text = "Status" Then
            cr.SortFields(0) = "+{VW_ConsultaTransacoesPorCategoria.Status}"
        ElseIf cmbOrdem.Text = "Categoria" Then
            cr.SortFields(0) = "+{VW_ConsultaTransacoesPorCategoria.Categoria}"
        End If

      ' cr.Formulas(1) = "F002 = 'Ordenar por : '" + cmbOrdem.Text


    cr.Connect = "DSN=" & xServer & ";UID=" & xUserName & ";PWD= " & xPassword & ";"
    cr.Action = 1
    Screen.MousePointer = vbDefault

End Function

