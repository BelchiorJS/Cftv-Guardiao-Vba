Attribute VB_Name = "DadosClientes"
Option Explicit

' =============================================================================
' Função:      Get_Clientes
' Descrição:
'   Retorna uma coleção de clientes contendo, para cada item, os dados padronizados de monitoramento operacional.
'
'   Cada cliente é armazenado em um Array na seguinte ordem:
'       (0) Nome do cliente                -> String
'       (1) Quantidade total de câmeras    -> Long
'       (2) Quantidade total de alarmes    -> Long
'
' Retorno:
'   Collection
'       Cada item da coleção é um Array com os dados do cliente.
'
'
' Observações:
'   - Os valores definidos nesta função representam a base operacional fixa.
'   - Alterações devem ser feitas apenas neste módulo, evitando mudanças
'     diretas nos formulários.
' =============================================================================
Public Function Get_Clientes() As Collection

    Dim c As New Collection

    c.Add Array("Acer", 16, 10)
    c.Add Array("Apae", 16, 10)
    c.Add Array("Canopus", 5, 0)
    c.Add Array("Dominalog", 16, 0)
    c.Add Array("PHL", 16, 2)
    c.Add Array("Novo Nordisk", 16, 0)
    c.Add Array("Sabesp", 4, 1)
    c.Add Array("Wilson Sons", 16, 4)

    Set Get_Clientes = c

End Function

