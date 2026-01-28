Attribute VB_Name = "Usuarios"
'===============================================================================
' MÓDULO: modUsuarios (Cadastro e Manutenção de Usuários)
'-------------------------------------------------------------------------------
' OBJETIVO
'   Fornecer rotinas auxiliares para inserir novos usuários na base de dados
'   do sistema (Excel), utilizando uma tabela estruturada (ListObject).
'
' FONTE DE DADOS
'   Planilha: DB_USUARIOS
'   Tabela (ListObject): TB_USUARIOS
'
' COLUNAS ESPERADAS NA TB_USUARIOS
'   - Id_Usuario       (numérico) : Identificador único do usuário
'   - Nome             (texto)    : Nome completo / apelido
'   - E-mail           (texto)    : Email do usuário (armazenado em minúsculas)
'   - Usuario          (texto)    : Login do usuário (armazenado em minúsculas)
'   - Senha            (texto)    : Senha em texto puro (ATENÇÃO: menos seguro)
'   - Perfil_Acesso    (texto)    : Perfil (ex.: Admin, Operador)
'   - Ativo            (texto)    : Status do usuário ("Sim"/"Não")
'   - Ultimo_Login     (data/hora): Inicialmente vazio; atualizado no login
'
' PROCEDIMENTO: Usuarios_Adicionar
'-------------------------------------------------------------------------------
' OBJETIVO
'   Adicionar um novo usuário no final da tabela TB_USUARIOS, preenchendo os
'   campos principais e definindo valores padrão para Ativo e Ultimo_Login.
'
' ASSINATURA
'   Public Sub Usuarios_Adicionar( _
'       ByVal Nome As String, _
'       ByVal Email As String, _
'       ByVal Usuario As String, _
'       ByVal Senha As String, _
'       ByVal Perfil As String _
'   )
'
' COMPORTAMENTO
'   - Localiza a planilha/tabela de usuários (DB_USUARIOS/TB_USUARIOS)
'   - Cria uma nova linha (ListRow) no final da tabela
'   - Gera um novo Id_Usuario sequencial (GerarNovoId)
'   - Normaliza Email e Usuario para minúsculas (evita duplicidade por case)
'   - Define:
'       • Ativo = "Sim"
'       • Ultimo_Login = vazio (vbNullString)
'
' OBSERVAÇÕES IMPORTANTES
'   - Este procedimento NÃO valida duplicidade de Usuario/E-mail.
'     Recomendação: antes de inserir, criar uma rotina de validação para:
'       • impedir Usuario repetido
'       • impedir E-mail repetido (se isso for uma regra do sistema)
'   - Segurança: senha em texto puro é vulnerável. Caso futuramente volte a usar
'     hash, este módulo deve ser ajustado para armazenar hash em vez de senha.
'
' FUNÇÃO: GerarNovoId
'-------------------------------------------------------------------------------
' OBJETIVO
'   Gerar um novo Id_Usuario sequencial com base no maior valor existente na
'   coluna "Id_Usuario" da tabela.
'
' ASSINATURA
'   Private Function GerarNovoId(ByVal tbl As ListObject) As Long
'
' COMPORTAMENTO
'   - Varre a coluna Id_Usuario
'   - Identifica o maior Id numérico presente
'   - Retorna (maior + 1)
'
' LIMITAÇÕES
'   - Se registros forem apagados, o Id continua crescendo (o que é desejável
'     na maioria dos casos).
'   - Em cenários multiusuário simultâneos (muitos salvando ao mesmo tempo),
'     pode haver colisão. Para Excel local costuma ser suficiente.
'===============================================================================


Option Explicit

Private Const NOME_PLANILHA As String = "DB_USUARIOS"
Private Const NOME_TABELA As String = "TB_USUARIOS"

Public Sub Usuarios_Adicionar(ByVal Nome As String, ByVal Email As String, ByVal Usuario As String, ByVal Senha As String, ByVal Perfil As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim r As ListRow

    Set ws = ThisWorkbook.Worksheets(NOME_PLANILHA)
    Set tbl = ws.ListObjects(NOME_TABELA)

    Set r = tbl.ListRows.Add

    With r.Range
        .Columns(tbl.ListColumns("Id_Usuario").Index).Value = GerarNovoId(tbl)
        .Columns(tbl.ListColumns("Nome").Index).Value = Nome
        .Columns(tbl.ListColumns("E-mail").Index).Value = LCase$(Email)
        .Columns(tbl.ListColumns("Usuario").Index).Value = LCase$(Usuario)
        .Columns(tbl.ListColumns("Senha").Index).Value = Senha
        .Columns(tbl.ListColumns("Perfil_Acesso").Index).Value = Perfil
        .Columns(tbl.ListColumns("Ativo").Index).Value = "Sim"
        .Columns(tbl.ListColumns("Ultimo_Login").Index).Value = vbNullString
    End With
End Sub

Private Function GerarNovoId(ByVal tbl As ListObject) As Long
    Dim maxId As Long
    Dim c As Range

    maxId = 0
    For Each c In tbl.ListColumns("Id_Usuario").DataBodyRange
        If IsNumeric(c.Value) Then
            If CLng(c.Value) > maxId Then maxId = CLng(c.Value)
        End If
    Next c

    GerarNovoId = maxId + 1
End Function

