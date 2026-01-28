VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPassagemPlantao 
   Caption         =   "Passagem de Plantão"
   ClientHeight    =   9696.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   16680
   OleObjectBlob   =   "frmPassagemPlantao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPassagemPlantao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Evento executado automaticamente ao abrir o UserForm.
' Responsável por inicializar campos, listas e clientes.
' ==========================================================
Private Sub UserForm_Initialize()
    
    Inicializar_Campos
    Definir_Turno_Automaticamente
    Inicializar_Listas
    Carregar_Lista_De_Clientes
    
End Sub

' Inicializa os campos fixos do formulário:
' - Data
' - Plantão
' - Turno
' - Operador
' ==========================================================
Private Sub Inicializar_Campos()
    
    txtData.Value = Format(Date, "dd/mm/yyyy")
    txtData.Enabled = False
    
    cboPlantao.List = Array("Alpha", "Bravo", "Charlie", "Delta")
    cboTurno.List = Array("Diurno", "Noturno")
    cboNome_Operador.List = Array( _
        "Anderson dos Santos", "Carlos Eduardo", "George Monteiro", _
        "Guilherme Belchior", "Janaina Martins", "Michael Douglas", _
        "Victoria Moreira", "Vinicius Costa" _
    )

End Sub

' Define automaticamente o turno com base no horário atual.
' Diurno: 07:00 às 18:59
' Noturno: 19:00 às 06:59
' ==========================================================
Private Sub Definir_Turno_Automaticamente()

    If Time >= TimeValue("07:00:00") And Time < TimeValue("19:00:00") Then
        cboTurno.Value = "Diurno"
    Else
        cboTurno.Value = "Noturno"
    End If
    
End Sub

' Inicializa as ListBox de câmeras e alarmes.
' ==========================================================
Private Sub Inicializar_Listas()
    Configurar_Lista lstCameras
    Configurar_Lista lstAlarmes
End Sub

' Configura o layout padrão das ListBox.
' Colunas:
' Cliente | Total de Equipamentos| Inoperantes | Operantes(Oculto, somente para o calculo da porcentagem) | Porcentagem
' ==========================================================
Private Sub Configurar_Lista(ByVal lst As MSForms.ListBox)

    With lst
        .ColumnCount = 5
        .ColumnWidths = "80 pt;80 pt;80 pt;0 pt;90 pt"
        .Clear
    End With

End Sub

' Carrega a lista de clientes a partir do módulo de dados.
' Preenche câmeras e alarmes separadamente.
' ==========================================================
Private Sub Carregar_Lista_De_Clientes()

    Dim Clientes As Collection
    Dim Item As Variant
    
    Set Clientes = Get_Clientes()
    
    For Each Item In Clientes
        Adicionar_Linha lstCameras, CStr(Item(0)), CLng(Item(1))
        Adicionar_Linha lstAlarmes, CStr(Item(0)), CLng(Item(2))
    Next Item
    
End Sub

' Adiciona uma linha padrão na ListBox.
' Inicializa com 0 inoperantes e 100% de disponibilidade.
' ==========================================================
Private Sub Adicionar_Linha(ByVal lst As MSForms.ListBox, _
                            ByVal Nome As String, _
                            ByVal Total As Long)
    
    Dim idx As Long
    idx = lst.ListCount
    
    lst.AddItem Nome
    lst.List(idx, 1) = Total
    lst.List(idx, 2) = 0
    lst.List(idx, 3) = Total
    lst.List(idx, 4) = "100%"
    
End Sub

' Evento acionado ao clicar em um cliente da lista de câmeras.
' Preenche o campo de inoperantes correspondente.
' ==========================================================
Private Sub lstCameras_Click()
    
    If lstCameras.ListIndex < 0 Then Exit Sub
    txtCameras_Inoperantes.Value = lstCameras.List(lstCameras.ListIndex, 2)

End Sub

' Evento acionado ao clicar em um cliente da lista de alarmes.
' Preenche o campo de inoperantes correspondente.
' ==========================================================
Private Sub lstAlarmes_Click()

    If lstAlarmes.ListIndex < 0 Then Exit Sub
    txtAlarmes_Inoperantes.Value = lstAlarmes.List(lstAlarmes.ListIndex, 2)

End Sub

' Atualiza o status operacional das câmeras selecionadas.
' ==========================================================
Private Sub btnAtualizarCameras_Click()
    Atualizar_Status lstCameras, txtCameras_Inoperantes
End Sub

' Atualiza o status operacional dos alarmes selecionados.
' ==========================================================
Private Sub btnAtualizarAlarmes_Click()
    Atualizar_Status lstAlarmes, txtAlarmes_Inoperantes
End Sub

' Atualiza os valores de:
' - Inoperantes
' - Operantes
' - Percentual de disponibilidade
' Realiza validações automáticas.
' ==========================================================
Private Sub Atualizar_Status(ByVal lst As MSForms.ListBox, _
                             ByVal txtInop As MSForms.TextBox)

    Dim i As Long, Total As Long, Inop As Long, Op As Long
    Dim Perc As Double
    
    i = lst.ListIndex
    If i < 0 Then
        MsgBox "Selecione um cliente na lista.", vbExclamation
        Exit Sub
    End If

    Total = CLng(lst.List(i, 1))

    If Trim(txtInop.Value) = "" Then txtInop.Value = "0"
    If Not IsNumeric(txtInop.Value) Then
        MsgBox "O valor só pode ser numérico.", vbExclamation
        Exit Sub
    End If

    Inop = CLng(txtInop.Value)
    If Inop < 0 Or Inop > Total Then
        MsgBox "O valor de equipamentos inoperantes não pode exceder o total do cliente.", vbExclamation
        Exit Sub
    End If

    Op = Total - Inop
    Perc = IIf(Total = 0, 0, Op / Total)

    lst.List(i, 2) = Inop
    lst.List(i, 3) = Op
    lst.List(i, 4) = Format(Perc, "0%")

End Sub


