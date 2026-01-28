Attribute VB_Name = "Auth"
' Módulo de autenticação de login do sistema
' ================================================
' Objetivo: Centralizar a validação de login do sistema com base na tabela de usuários armazenada no Excel. Este módulo _
' lê os dados de credenciais diretamente da planilha "DB_USUARIOS", na tabela "TB_USUARIOS".
'
' Colunas que devem existir na TB_USUARIOS:
'   - Id_Usuario       (numérico) : Identificador do usuário
'   - Nome             (texto)    : Nome completo / apelido
'   - Usuario          (texto)    : Login (comparação case-insensitive)
'   - Senha            (texto)    : Senha em texto puro (ATENÇÃO: menos seguro)
'   - Perfil_Acesso    (texto)    : Perfil do usuário (ex.: Admin, Operador)
'   - Ativo            (texto)    : "SIM" para permitir login (case-insensitive)
'   - Ultimo_Login     (data/hora): Atualizado no login bem sucedido
'
' DEPENDÊNCIAS:
'   - Variáveis de sessão (devem existir em outro módulo, ex.: modSessao):
'       Public IdUsuarioLogado As Long
'       Public UsuarioLogado   As String
'       Public NomeLogado      As String
'       Public PerfilLogado    As String
'
' FUNÇÃO: Autenticar_Login
'
' PARÂMETROS:
'   Usuario        : Login digitado pelo usuário (normalizado para minúsculas)
'   SenhaDigitada  : Senha digitada pelo usuário (comparação literal)
'   msgErro        : Mensagem de retorno em caso de falha (saída por referência)
'
' RETORNO:
'   True  -> Autenticação bem sucedida; sessão preenchida e Ultimo_Login atualizado
'   False -> Falha; msgErro descreve o motivo (campos vazios, inativo, etc.)
'
' REGRAS DE NEGÓCIO / VALIDAÇÕES
'   1) Usuário e senha são obrigatórios (não aceita vazio)
'   2) Procura o usuário na tabela (case-insensitive)
'   3) Verifica se o usuário está ativo (Ativo = "SIM")
'   4) Verifica se existe senha cadastrada
'   5) Compara SenhaDigitada com Senha (texto puro)
'   6) Em caso de sucesso:
'        - Preenche variáveis globais de sessão (IdUsuarioLogado, etc.)
'        - Atualiza Ultimo_Login com Now
'
' TRATAMENTO DE ERROS
'   - Se a planilha/tabela/colunas não existirem ou não puderem ser acessadas,
'     retorna False e msgErro = "Erro ao acessar DB_USUARIOS / TB_USUARIOS ..."
' ================================================================


Option Explicit

Private Const NOME_PLANILHA As String = "DB_USUARIOS"
Private Const NOME_TABELA As String = "TB_USUARIOS"

Public Function Autenticar_Login(ByVal Usuario As String, ByVal SenhaDigitada As String, ByRef msgErro As String) As Boolean
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim r As ListRow

    Usuario = LCase$(Trim$(Usuario))
    SenhaDigitada = CStr(SenhaDigitada)

    If Usuario = "" Or SenhaDigitada = "" Then
        msgErro = "Usuário e senha são obrigatórios."
        Autenticar_Login = False
        Exit Function
    End If

    On Error GoTo ErroEstrutura
    Set ws = ThisWorkbook.Worksheets(NOME_PLANILHA)
    Set tbl = ws.ListObjects(NOME_TABELA)
    On Error GoTo 0

    For Each r In tbl.ListRows
        If LCase$(Trim$(r.Range.Columns(tbl.ListColumns("Usuario").Index).Value)) = Usuario Then

            If UCase$(Trim$(r.Range.Columns(tbl.ListColumns("Ativo").Index).Value)) <> "SIM" Then
                msgErro = "Usuário inativo!"
                Autenticar_Login = False
                Exit Function
            End If

            Dim senhaBase As String
            senhaBase = CStr(r.Range.Columns(tbl.ListColumns("Senha").Index).Value)

            If senhaBase = "" Then
                msgErro = "Usuário sem senha configurada."
                Autenticar_Login = False
                Exit Function
            End If

            If senhaBase = SenhaDigitada Then
                ' Login OK
                IdUsuarioLogado = CLng(r.Range.Columns(tbl.ListColumns("Id_Usuario").Index).Value)
                UsuarioLogado = CStr(r.Range.Columns(tbl.ListColumns("Usuario").Index).Value)
                NomeLogado = CStr(r.Range.Columns(tbl.ListColumns("Nome").Index).Value)
                PerfilLogado = CStr(r.Range.Columns(tbl.ListColumns("Perfil_Acesso").Index).Value)

                r.Range.Columns(tbl.ListColumns("Ultimo_Login").Index).Value = Now

                msgErro = vbNullString
                Autenticar_Login = True
                Exit Function
            Else
                msgErro = "Senha incorreta."
                Autenticar_Login = False
                Exit Function
            End If
        End If
    Next r

    msgErro = "Usuário não encontrado."
    Autenticar_Login = False
    Exit Function

ErroEstrutura:
    msgErro = "Erro ao acessar DB_USUARIOS / TB_USUARIOS (planilha/tabela/colunas)."
    Autenticar_Login = False
End Function

