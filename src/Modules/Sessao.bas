Attribute VB_Name = "Sessao"
'===============================================================================
' MÓDULO: Sessao (Gerenciamento de Sessão)
'-------------------------------------------------------------------------------
' OBJETIVO
'   Armazenar e controlar as informações da sessão do usuário atualmente
'   autenticado no sistema. Este módulo centraliza os dados do usuário logado
'   para que possam ser acessados por outros módulos, formulários e rotinas
'   durante a execução da aplicação.
'
' VARIÁVEIS DE SESSÃO (ESCOPO PÚBLICO)
'   IdUsuarioLogado : Identificador único do usuário autenticado
'   UsuarioLogado   : Login do usuário autenticado
'   NomeLogado      : Nome/apelido do usuário autenticado
'   PerfilLogado    : Perfil de acesso do usuário (ex.: Admin, Operador)
'
' USO TÍPICO
'   - Preenchidas após autenticação bem sucedida (modAuth.Autenticar_Login)
'   - Utilizadas para:
'       • Controle de permissões por perfil
'       • Exibição de informações do usuário na interface
'       • Auditoria e registro de ações
'
' PROCEDIMENTO: Sessao_Limpar
'-------------------------------------------------------------------------------
' OBJETIVO
'   Limpar todas as variáveis de sessão, retornando o sistema ao estado
'   "não autenticado". Deve ser chamado em cenários de:
'       • Logout
'       • Encerramento do sistema
'       • Falha crítica que exija reinício da sessão
'
' COMPORTAMENTO
'   - Zera o identificador do usuário
'   - Remove todas as informações de login armazenadas em memória

Option Explicit

Public IdUsuarioLogado As Long
Public UsuarioLogado As String
Public NomeLogado As String
Public PerfilLogado As String

Public Sub Sessao_Limpar()
    IdUsuarioLogado = 0
    UsuarioLogado = vbNullString
    NomeLogado = vbNullString
    PerfilLogado = vbNullString
End Sub

