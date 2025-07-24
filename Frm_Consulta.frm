VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Frm_Consulta 
   Caption         =   "Form1"
   ClientHeight    =   10140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16995
   Icon            =   "Frm_Consulta.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10140
   ScaleWidth      =   16995
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1140
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   16995
      _ExtentX        =   29977
      _ExtentY        =   2011
      ButtonWidth     =   1455
      ButtonHeight    =   1852
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
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
            Caption         =   "Fechar"
            Key             =   "Fechar"
            Object.ToolTipText     =   "Fechar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
      OLEDropMode     =   1
   End
   Begin VB.TextBox txtPaginaAtual 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11160
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   8520
      Width           =   735
   End
   Begin VB.CommandButton cmdUltima 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   11
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton cmdProxima 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   10
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Picture         =   "Frm_Consulta.frx":1084A
      TabIndex        =   9
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrimeira 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "Frm_Consulta.frx":11914
      TabIndex        =   8
      Top             =   8640
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4455
      Left            =   0
      TabIndex        =   7
      Top             =   3960
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Consulta"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   12975
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7080
         Picture         =   "Frm_Consulta.frx":2215E
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Buscar Transações"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox cboRegistros 
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
         Left            =   8640
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   240
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtData 
         Height          =   495
         Left            =   3240
         TabIndex        =   15
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_NumeroCartao 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         MaxLength       =   16
         TabIndex        =   3
         ToolTipText     =   "16 dígitos"
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txt_Valor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         ToolTipText     =   "Somente decimais positivos"
         Top             =   1320
         Width           =   2775
      End
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
         ItemData        =   "Frm_Consulta.frx":22391
         Left            =   3240
         List            =   "Frm_Consulta.frx":2239E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Limite de Registros:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6480
         TabIndex        =   17
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Transação:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   " Número Cartão :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor (R$) :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Status :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   1920
         Width           =   1695
      End
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   13200
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Frm_Consulta.frx":223CD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   9960
      Top             =   8880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Frm_Consulta.frx":22B8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Frm_Consulta.frx":23AA1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      Caption         =   "Página"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   18
      Top             =   8520
      Width           =   855
   End
   Begin VB.Label lblTotalPaginas 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12000
      TabIndex        =   13
      Top             =   8520
      Width           =   1455
   End
End
Attribute VB_Name = "Frm_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CurrentPage As Long    ' Página atual que o usuário está vendo
Private TotalRecords As Long   ' Total de registros na tabela (para cálculo do total de páginas)
Private TotalPages As Long     ' Total de páginas calculadas

Dim rs As New ADODB.Recordset
Dim SQL As String


Private Sub cmdBuscar_Click()
    Dim sSQL As String
    Dim sWhereClause As String
    Dim sNumeroCartao As String
    Dim sValor As String
    Dim dData As Date
    Dim sStatus As String


On Error GoTo TratamentoDeErro

        ' --- 1. Obter os valores dos controles da tela ---
        sNumeroCartao = Trim(txt_NumeroCartao.Text)
        sValor = Trim(txt_Valor.Text)
        
        '' Validação de Datas
        If IsDate(txtData.Text) Then
            dData = CDate(txtData.Text)
        Else
            ' Se a data não for válida, trate como se não tivesse sido informada
            ' ou avise o usuário. Aqui, vamos considerar como não informada.
            dData = CDate("01/01/0001") ' Um valor "nulo" para Data para indicar que não foi preenchida
        End If
        
        ' Para ComboBox, pegue o valor selecionado (assumindo que "Todos" ou vazio signifique não filtrar)
        If cmbStatus.ListIndex >= 0 Then ' Se algum item foi selecionado
            sStatus = cmbStatus.ListIndex + 1
            If sStatus = "Todos" Then sStatus = "" ' Se tiver uma opção "Todos", trate como vazio
        Else
            sStatus = "" ' Nenhum item selecionado ou vazio
        End If
    
        ' --- 2. Construir a cláusula WHERE dinamicamente ---
        sWhereClause = " WHERE 1=1 " ' Inicia com uma condição sempre verdadeira
    
        ' Filtro por Número do Cartão (assumindo que é um LIKE para pesquisa parcial)
        If sNumeroCartao <> "" Then
            ' Use LIKE para pesquisa parcial. Para segurança, idealmente use parâmetros ADO.
            sWhereClause = sWhereClause & " AND Numero_Cartao LIKE '%" & sNumeroCartao & "%'"
        End If
    
        ' Filtro por Valor (assumindo que é um campo numérico no DB)
        If IsNumeric(sValor) Then
            sWhereClause = sWhereClause & " AND Valor_transacao = " & Replace(CDbl(sValor), ",", ".") ' Use CDbl para valores decimais
        End If
    
        ' Filtro por Data (considerando apenas a data, sem a hora)
        ' Ajuste a função de data conforme seu SGBD (SQL Server, MySQL, Oracle, PostgreSQL)
        If dData <> CDate("01/01/0001") Then ' Se a data foi preenchida e é válida
            ' Exemplo para SQL Server: CAST(DataCampoDB AS DATE) = 'yyyy-mm-dd'
            ' (Para SQL Server, use estilo 120 para 'yyyy-mm-dd')
            sWhereClause = sWhereClause & " AND CAST(Data_Transacao AS DATE) = '" & Format(dData, "yyyy-mm-dd") & "'"
            ' Outros SGBDs:
            ' MySQL/PostgreSQL: AND DATE(DataTransacao) = '" & Format(dData, "yyyy-mm-dd") & "'"
            ' Oracle: AND TRUNC(DataTransacao) = TO_DATE('" & Format(dData, "yyyy-mm-dd") & "', 'YYYY-MM-DD')"
        End If
    
        ' Filtro por Status
        If sStatus <> "" Then
            ' Assumindo que o campo Status no DB é um texto
            sWhereClause = sWhereClause & " AND Status = '" & sStatus & "'"
        End If
    
    

        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
        
        
        sSQL = "select id_transacao, Numero_Cartao, Valor_transacao, Data_Transacao,descricao, Status from tb_Transacoes "
        sSQL = sSQL & sWhereClause
        sSQL = sSQL & " ORDER BY Data_Transacao DESC"
        rs.Open sSQL, cn, 3, 3
        
        
         Set DataGrid1.DataSource = rs

        ' Ajusta o layout do DataGrid automaticamente (opcional)
        If DataGrid1.Columns.Count > 0 Then
            DataGrid1.Refresh
        End If
    
    

Exit Sub
TratamentoDeErro:
    '' Monta a mensagem de log com detalhes do erro
    Dim strErroDetails As String
    strErroDetails = "Erro na rotina MinhaRotinaQuePodeGerarErro - " & _
                     "Número: " & Err.Number & " | " & _
                     "Descrição: " & Err.Description & " | " & _
                     "Fonte: " & Err.Source & " | " & _
                     "ÚltimaDLL: " & Err.HelpFile & " | " & _
                     "Contexto: Linha do erro/Estado da aplicação" ' Adicione contexto se possível

    '' Chama a rotina de log do módulo1
    Call EscreverLogErro(strErroDetails)

    '' Opcional: Avisar o usuário de forma amigável (sem mostrar detalhes técnicos)
    MsgBox "Ocorreu um erro inesperado. O problema foi registrado e será investigado.", vbCritical, "Erro"
    
End Sub



Private Sub Form_Load()

    '' Renomeia título da Tela
    Me.Caption = "XYZ - Administradora de Cartões de Crédito - " + Me.Caption
    
    cmbStatus.ListIndex = 1
    '' 1. Inicializa variáveis de paginação
    CurrentPage = 1 ' Começa na primeira página

    '' Configura a ComboBox de registros por página (opcional)
    cboRegistros.Clear
    cboRegistros.AddItem "10"
    cboRegistros.AddItem "20" ' Define um padrão inicial
    cboRegistros.AddItem "50"
    cboRegistros.Text = "20" ' Valor padrão
    
    '' 3. Carrega a primeira página de dados
    Call LoadPage
End Sub


Private Sub LoadPage()
    Dim sSQLCount As String
    Dim sSQLData As String
    Dim lOffset As Long
    Dim lRecordsPerPage As Long

    On Error GoTo ErrorHandler

    lRecordsPerPage = CLng(cboRegistros.Text) ' Pega o valor da ComboBox

    ' 1. Fecha o Recordset existente se estiver aberto
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If

    ' 2. Obtém o Total de Registros (para calcular TotalPages)
    '    Isso é feito uma vez ou sempre que o filtro mudar.
    sSQLCount = "SELECT COUNT(*) FROM tb_Transacoes;" ' <-- ALtere para sua tabela real
    Set rs = New ADODB.Recordset
    rs.Open sSQLCount, cn, adOpenStatic, adLockReadOnly
    TotalRecords = rs.Fields(0).Value
    rs.Close
    Set rs = Nothing ' Libera o recordset de contagem

    ' 3. Calcula o total de páginas
    TotalPages = Int(TotalRecords / lRecordsPerPage)
    If (TotalRecords Mod lRecordsPerPage) > 0 Then
        TotalPages = TotalPages + 1 ' Adiciona uma página extra se houver "sobras"
    End If

    ' 4. Ajusta CurrentPage se for inválida (ex: após filtros ou exclusões)
    If CurrentPage < 1 And TotalPages > 0 Then CurrentPage = 1
    If CurrentPage > TotalPages And TotalPages > 0 Then CurrentPage = TotalPages
    If TotalPages = 0 Then CurrentPage = 0 ' Se não houver registros, não há página 1

    ' 5. Calcula o OFFSET para a query paginada
    lOffset = (CurrentPage - 1) * lRecordsPerPage

    ' 6. Monta a SQL para buscar os dados da página atual
    '    AJUSTE ESSA QUERY PARA SEU SGBD E COLUNAS!
    sSQLData = "SELECT ID_transacao, Numero_Cartao, Valor_Transacao,Data_Transacao, Descricao,status FROM tb_Transacoes " & _
               "ORDER BY ID_transacao " & _
               "OFFSET " & lOffset & " ROWS " & _
               "FETCH NEXT " & lRecordsPerPage & " ROWS ONLY;"

    ' 7. Abre o Recordset com os dados da página
    Set rs = New ADODB.Recordset
    rs.Open sSQLData, cn, adOpenKeyset, adLockOptimistic ' adOpenKeyset ou adOpenStatic são bons para paginação

    ' 8. Conecta o Recordset ao DataGrid
    Set DataGrid1.DataSource = rs

    ' 9. Atualiza os indicadores de paginação na interface
    txtPaginaAtual.Text = CStr(CurrentPage)
    lblTotalPaginas.Caption = "de " & CStr(TotalPages)

    ' 10. Habilita/Desabilita botões de navegação
    cmdPrimeira.Enabled = (CurrentPage > 1)
    cmdAnterior.Enabled = (CurrentPage > 1)
    cmdProxima.Enabled = (CurrentPage < TotalPages) And (TotalPages > 0)
    cmdUltima.Enabled = (CurrentPage < TotalPages) And (TotalPages > 0)

    ' Desabilita tudo se não houver registros
    If TotalRecords = 0 Then
        cmdPrimeira.Enabled = False
        cmdAnterior.Enabled = False
        cmdProxima.Enabled = False
        cmdUltima.Enabled = False
        txtPaginaAtual.Text = "0"
        lblTotalPaginas.Caption = "de 0"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Erro ao carregar dados: " & Err.Description & vbCrLf & "SQL da página: " & sSQLData, vbCritical
    If Not rs Is Nothing Then If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

' --- Eventos dos Botões de Navegação ---
Private Sub cmdPrimeira_Click()
    If CurrentPage > 1 Then
        CurrentPage = 1
        Call LoadPage
    End If
End Sub

Private Sub cmdAnterior_Click()
    If CurrentPage > 1 Then
        CurrentPage = CurrentPage - 1
        Call LoadPage
    End If
End Sub

Private Sub cmdProxima_Click()
    If CurrentPage < TotalPages Then
        CurrentPage = CurrentPage + 1
        Call LoadPage
    End If
End Sub

Private Sub cmdUltima_Click()
    If CurrentPage < TotalPages Then
        CurrentPage = TotalPages
        Call LoadPage
    End If
End Sub

' --- Evento da ComboBox de Registros por Página (se mudar o valor) ---
Private Sub cboRegistros_Change()
    ' Verifica se é um número válido e maior que zero
    If IsNumeric(cboRegistros.Text) Then
        If CLng(cboRegistros.Text) > 0 Then
            CurrentPage = 1 ' Redefine para a primeira página ao mudar o tamanho da página
            Call LoadPage
    End If
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

On Error GoTo TratamentoDeErro

Select Case Button.Key

    Case "Limpar"
            Call fLimparCampos
     
    Case "Fechar"
            Unload Me
    
End Select

Exit Sub
TratamentoDeErro:
    '' Monta a mensagem de log com detalhes do erro
    Dim strErroDetails As String
    strErroDetails = "Erro na rotina MinhaRotinaQuePodeGerarErro - " & _
                     "Número: " & Err.Number & " | " & _
                     "Descrição: " & Err.Description & " | " & _
                     "Fonte: " & Err.Source & " | " & _
                     "ÚltimaDLL: " & Err.HelpFile & " | " & _
                     "Contexto: Linha do erro/Estado da aplicação" ' Adicione contexto se possível

    '' Chama a rotina de log do módulo1
    Call EscreverLogErro(strErroDetails)

    '' Opcional: Avisar o usuário de forma amigável (sem mostrar detalhes técnicos)
    MsgBox "Ocorreu um erro inesperado. O problema foi registrado e será investigado.", vbCritical, "Erro"




End Sub


Private Function fLimparCampos()

    txt_NumeroCartao.Text = ""
    txt_Valor.Text = ""
    
    
End Function



' --- Evento da TextBox de Página Atual (para ir para uma página específica) ---
Private Sub txtPaginaAtual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ' Se o usuário pressionar Enter
        If IsNumeric(txtPaginaAtual.Text) Then
            Dim NewPage As Long
            NewPage = CLng(txtPaginaAtual.Text)
            ' Valida se a página digitada está dentro dos limites
            If NewPage >= 1 And NewPage <= TotalPages Then
                CurrentPage = NewPage
                Call LoadPage
            ElseIf TotalPages = 0 And NewPage = 0 Then
                CurrentPage = 0 ' Permite 0/0 se não houver registros
            Else
                MsgBox "Número de página inválido. Digite entre 1 e " & TotalPages, vbExclamation
                txtPaginaAtual.Text = CStr(CurrentPage) ' Volta para a página anterior
            End If
        Else
            MsgBox "Por favor, digite um número para a página.", vbExclamation
            txtPaginaAtual.Text = CStr(CurrentPage)
        End If
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
End Sub

