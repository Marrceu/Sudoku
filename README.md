Option Explicit

' ==============================
' Configuração e constantes
' ==============================
Public Const TOP_ROW As Long = 4
Public Const LEFT_COL As Long = 2
Public Const SIZE As Long = 9

' Cores de fonte
Public Const COLOR_PRESET As Long = vbBlack
Public Const COLOR_USER As Long = &HCC6600
Public Const COLOR_OK As Long = &H9900
Public Const COLOR_ERR As Long = &HDC

' Cores de fundo
Public Const BG_OK As Long = &HDCEEDC          ' Verde claro
Public Const BG_ERR As Long = &HDEDDFF         ' Vermelho claro (BBGGRR)

' Estado e cronômetro
Private IsTimerRunning As Boolean
Private StartTime As Date
Private ElapsedBefore As Double
Private NextTick As Date

Private PresetMask(1 To SIZE, 1 To SIZE) As Boolean
Private Solution(1 To SIZE, 1 To SIZE) As Integer

' ==============================
' Interface principal
' ==============================
Public Sub CarregarInterfaceSudoku()
    SetupBoard
    CriarBotoesLaterais
    AdicionarBotaoDica
End Sub

Public Sub AdicionarBotaoCarregarSudoku()
    Dim ws As Worksheet: Set ws = ActiveSheet
    On Error Resume Next
    ws.Shapes("btnCarregarInterfaceSudoku").Delete
    On Error GoTo 0

    Dim shp As Shape
    Set shp = ws.Shapes.AddFormControl(xlButtonControl, 10, 10, 220, 28)
    shp.Name = "btnCarregarInterfaceSudoku"
    shp.TextFrame.Characters.Text = "Carregar Interface Sudoku"
    ws.Buttons(shp.Name).OnAction = "'" & ThisWorkbook.Name & "'!CarregarInterfaceSudoku"
End Sub

Private Sub CriarBotoesLaterais()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim nomes, macros
    nomes = Array( _
        "Novo (Fácil)", _
        "Novo (Médio)", _
        "Novo (Difícil)", _
        "Verificar", _
        "Corrigir", _
        "Resetar", _
        "Iniciar/Pausar Cronômetro" _
    )
    macros = Array( _
        "NovoFacil", _
        "NovoMedio", _
        "NovoDificil", _
        "VerificarSudoku", _
        "CorrigirSudoku", _
        "ResetarJogo", _
        "ToggleCronometro" _
    )

    ' Apagar botões antigos
    Dim i As Long
    For i = 0 To 20
        On Error Resume Next
        ws.Shapes("btnSudoku" & i).Delete
        On Error GoTo 0
    Next i

    ' Posicionar à direita do tabuleiro
    Dim x As Double, y As Double
    x = ws.Cells(TOP_ROW, LEFT_COL + SIZE + 1).Left
    y = ws.Cells(TOP_ROW, LEFT_COL + SIZE + 1).Top

    Dim shp As Shape
    For i = LBound(nomes) To UBound(nomes)
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, x, y + i * 30, 200, 24)
        shp.Name = "btnSudoku" & i
        shp.TextFrame.Characters.Text = nomes(i)
        ws.Buttons(shp.Name).OnAction = "'" & ThisWorkbook.Name & "'!" & macros(i)
    Next i
End Sub

Public Sub AdicionarBotaoDica()
    Dim ws As Worksheet: Set ws = ActiveSheet
    On Error Resume Next
    ws.Shapes("btnDica").Delete
    On Error GoTo 0

    Dim x As Double, y As Double
    x = ws.Cells(TOP_ROW, LEFT_COL + SIZE + 1).Left
    y = ws.Cells(TOP_ROW + 8, LEFT_COL + SIZE + 1).Top + 40

    Dim shp As Shape
    Set shp = ws.Shapes.AddFormControl(xlButtonControl, x, y, 200, 24)
    shp.Name = "btnDica"
    shp.TextFrame.Characters.Text = "Dica"
    ws.Buttons(shp.Name).OnAction = "'" & ThisWorkbook.Name & "'!DarDica"
End Sub

' ==============================
' Tabuleiro: desenho e visual
' ==============================
Public Sub SetupBoard()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Application.ScreenUpdating = False

    ws.Cells.Clear
    ws.Range("A1").Value = "Sudoku"
    SetupCronometroVisual ws
    DrawBoard ws
    ResetCronometro ws

    Application.ScreenUpdating = True
End Sub

Private Sub SetupCronometroVisual(ws As Worksheet)
    With ws.Range("B2:D2")
        .Merge
        .Value = "00:00:00"
        .NumberFormat = "[h]:mm:ss"
        .Interior.Color = RGB(30, 30, 30)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Consolas"
        .Font.SIZE = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Range("A2").Value = "Tempo:"
    ws.Range("E2").Value = "Tempo Final:"
    ws.Range("E2").Font.Bold = True
End Sub

Private Sub DrawBoard(ws As Worksheet)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(TOP_ROW, LEFT_COL), ws.Cells(TOP_ROW + SIZE - 1, LEFT_COL + SIZE - 1))

    With rng
        .Clear
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.SIZE = 14
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        .Borders.Weight = xlThin
        .Locked = False
    End With

    SetSquareCells ws, 28 ' células quadradas

    ' Pinta blocos 3x3
    Dim r As Long, c As Long
    For r = 1 To SIZE
        For c = 1 To SIZE
            With ws.Cells(TOP_ROW + r - 1, LEFT_COL + c - 1)
                .Value = ""
                .Interior.Color = GetBlockColor(r, c)
            End With
        Next c
    Next r

    ' Bordas grossas
    Dim i As Long
    For i = 0 To 3
        With ws.Range(ws.Cells(TOP_ROW + i * 3, LEFT_COL), ws.Cells(TOP_ROW + i * 3, LEFT_COL + 8)).Borders(xlEdgeTop)
            .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(0, 0, 0)
        End With
        With ws.Range(ws.Cells(TOP_ROW, LEFT_COL + i * 3), ws.Cells(TOP_ROW + 8, LEFT_COL + i * 3)).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(0, 0, 0)
        End With
    Next i
    With ws.Range(ws.Cells(TOP_ROW + 8, LEFT_COL), ws.Cells(TOP_ROW + 8, LEFT_COL + 8)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(0, 0, 0)
    End With
    With ws.Range(ws.Cells(TOP_ROW, LEFT_COL + 8), ws.Cells(TOP_ROW + 8, LEFT_COL + 8)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(0, 0, 0)
    End With
End Sub

Private Sub SetSquareCells(ws As Worksheet, targetPts As Double)
    Dim r As Long, c As Long, col As Long

    For r = 0 To SIZE - 1
        ws.Rows(TOP_ROW + r).RowHeight = targetPts
    Next r

    For c = 0 To SIZE - 1
        col = LEFT_COL + c
        ws.Columns(col).ColumnWidth = 3
        Dim tries As Integer: tries = 0
        Do While Abs(ws.Cells(TOP_ROW, col).Width - targetPts) > 0.5 And tries < 20
            If ws.Cells(TOP_ROW, col).Width < targetPts Then
                ws.Columns(col).ColumnWidth = ws.Columns(col).ColumnWidth + 0.2
            Else
                ws.Columns(col).ColumnWidth = ws.Columns(col).ColumnWidth - 0.2
            End If
            tries = tries + 1
        Loop
    Next c
End Sub

Private Function GetBlockColor(r As Long, c As Long) As Long
    Dim blockRow As Long, blockCol As Long
    blockRow = Int((r - 1) / 3)
    blockCol = Int((c - 1) / 3)
    If (blockRow + blockCol) Mod 2 = 0 Then
        GetBlockColor = RGB(240, 240, 255) ' Azul claro
    Else
        GetBlockColor = RGB(255, 255, 220) ' Amarelo claro
    End If
End Function

' ==============================
' Puzzles base por nível
' ==============================
Private Function BaseFacil() As Variant
    BaseFacil = Array( _
        Array(0, 0, 0, 2, 6, 0, 7, 0, 1), _
        Array(6, 8, 0, 0, 7, 0, 0, 9, 0), _
        Array(1, 9, 0, 0, 0, 4, 5, 0, 0), _
        Array(8, 2, 0, 1, 0, 0, 0, 4, 0), _
        Array(0, 0, 4, 6, 0, 2, 9, 0, 0), _
        Array(0, 5, 0, 0, 0, 3, 0, 2, 8), _
        Array(0, 0, 9, 3, 0, 0, 0, 7, 4), _
        Array(0, 4, 0, 0, 5, 0, 0, 3, 6), _
        Array(7, 0, 3, 0, 1, 8, 0, 0, 0))
End Function

Private Function SolFacil() As Variant
    SolFacil = Array( _
        Array(4, 3, 5, 2, 6, 9, 7, 8, 1), _
        Array(6, 8, 2, 5, 7, 1, 4, 9, 3), _
        Array(1, 9, 7, 8, 3, 4, 5, 6, 2), _
        Array(8, 2, 6, 1, 9, 5, 3, 4, 7), _
        Array(3, 7, 4, 6, 8, 2, 9, 1, 5), _
        Array(9, 5, 1, 7, 4, 3, 6, 2, 8), _
        Array(5, 1, 9, 3, 2, 6, 8, 7, 4), _
        Array(2, 4, 8, 9, 5, 7, 1, 3, 6), _
        Array(7, 6, 3, 4, 1, 8, 2, 5, 9))
End Function

Private Function BaseMedio() As Variant
    BaseMedio = Array( _
        Array(5, 3, 0, 0, 7, 0, 0, 0, 0), _
        Array(6, 0, 0, 1, 9, 5, 0, 0, 0), _
        Array(0, 9, 8, 0, 0, 0, 0, 6, 0), _
        Array(8, 0, 0, 0, 6, 0, 0, 0, 3), _
        Array(4, 0, 0, 8, 0, 3, 0, 0, 1), _
        Array(7, 0, 0, 0, 2, 0, 0, 0, 6), _
        Array(0, 6, 0, 0, 0, 0, 2, 8, 0), _
        Array(0, 0, 0, 4, 1, 9, 0, 0, 5), _
        Array(0, 0, 0, 0, 8, 0, 0, 7, 9))
End Function

Private Function SolMedio() As Variant
    SolMedio = Array( _
        Array(5, 3, 4, 6, 7, 8, 9, 1, 2), _
        Array(6, 7, 2, 1, 9, 5, 3, 4, 8), _
        Array(1, 9, 8, 3, 4, 2, 5, 6, 7), _
        Array(8, 5, 9, 7, 6, 1, 4, 2, 3), _
        Array(4, 2, 6, 8, 5, 3, 7, 9, 1), _
        Array(7, 1, 3, 9, 2, 4, 8, 5, 6), _
        Array(9, 6, 1, 5, 3, 7, 2, 8, 4), _
        Array(2, 8, 7, 4, 1, 9, 6, 3, 5), _
        Array(3, 4, 5, 2, 8, 6, 1, 7, 9))
End Function

Private Function BaseDificil() As Variant
    BaseDificil = Array( _
        Array(0, 0, 3, 0, 2, 0, 6, 0, 0), _
        Array(9, 0, 0, 3, 0, 5, 0, 0, 1), _
        Array(0, 0, 1, 8, 0, 6, 4, 0, 0), _
        Array(0, 0, 8, 1, 0, 2, 9, 0, 0), _
        Array(7, 0, 0, 0, 0, 0, 0, 0, 8), _
        Array(0, 0, 6, 7, 0, 8, 2, 0, 0), _
        Array(0, 0, 2, 6, 0, 9, 5, 0, 0), _
        Array(8, 0, 0, 2, 0, 3, 0, 0, 9), _
        Array(0, 0, 5, 0, 1, 0, 3, 0, 0))
End Function

Private Function SolDificil() As Variant
    SolDificil = Array( _
        Array(4, 8, 3, 9, 2, 1, 6, 5, 7), _
        Array(9, 6, 7, 3, 4, 5, 8, 2, 1), _
        Array(2, 5, 1, 8, 7, 6, 4, 9, 3), _
        Array(5, 4, 8, 1, 3, 2, 9, 7, 6), _
        Array(7, 2, 9, 5, 6, 4, 1, 3, 8), _
        Array(1, 3, 6, 7, 9, 8, 2, 4, 5), _
        Array(3, 7, 2, 6, 8, 9, 5, 1, 4), _
        Array(8, 1, 4, 2, 5, 3, 7, 6, 9), _
        Array(6, 9, 5, 4, 1, 7, 3, 8, 2))
End Function

' ==============================
' Randomização e carregamento
' ==============================
Private Sub RandomizePuzzle(ByRef puz As Variant, ByRef sol As Variant)
    Randomize

    Dim t As Long, k As Long
    ' Permutar dígitos 1..9
    Dim map(1 To 9) As Long, used(1 To 9) As Boolean
    For k = 1 To 9
        Do
            t = Int(Rnd * 9) + 1
        Loop While used(t)
        used(t) = True
        map(k) = t
    Next k
    ApplyDigitPermutation puz, map, True
    ApplyDigitPermutation sol, map, False

    ' Trocas internas
    For k = 1 To 10
        SwapRowsInBand puz, sol, Int(Rnd * 3), Int(Rnd * 3), Int(Rnd * 3)
        SwapColsInStack puz, sol, Int(Rnd * 3), Int(Rnd * 3), Int(Rnd * 3)
    Next k

    ' Trocas de bandas/pilhas
    For k = 1 To 3
        SwapRowBands puz, sol, Int(Rnd * 3), Int(Rnd * 3)
        SwapColStacks puz, sol, Int(Rnd * 3), Int(Rnd * 3)
    Next k
End Sub

Private Sub ApplyDigitPermutation(ByRef grid As Variant, ByRef map() As Long, ByVal allowZero As Boolean)
    Dim r As Long, c As Long, v As Long
    For r = 0 To 8
        For c = 0 To 8
            v = grid(r)(c)
            If v = 0 Then
                If allowZero Then
                    ' mantém 0
                End If
            Else
                grid(r)(c) = map(v)
            End If
        Next c
    Next r
End Sub

Private Sub SwapRowsInBand(ByRef puz As Variant, ByRef sol As Variant, ByVal band As Long, ByVal A As Long, ByVal B As Long)
    If A = B Then Exit Sub
    Dim r1 As Long, r2 As Long, tmp As Variant
    r1 = band * 3 + A: r2 = band * 3 + B
    tmp = puz(r1): puz(r1) = puz(r2): puz(r2) = tmp
    tmp = sol(r1): sol(r1) = sol(r2): sol(r2) = tmp
End Sub

Private Sub SwapColsInStack(ByRef puz As Variant, ByRef sol As Variant, ByVal stack As Long, ByVal A As Long, ByVal B As Long)
    If A = B Then Exit Sub
    Dim c1 As Long, c2 As Long, r As Long, t
    c1 = stack * 3 + A: c2 = stack * 3 + B
    For r = 0 To 8
        t = puz(r)(c1): puz(r)(c1) = puz(r)(c2): puz(r)(c2) = t
        t = sol(r)(c1): sol(r)(c1) = sol(r)(c2): sol(r)(c2) = t
    Next r
End Sub

Private Sub SwapRowBands(ByRef puz As Variant, ByRef sol As Variant, ByVal A As Long, ByVal B As Long)
    If A = B Then Exit Sub
    Dim i As Long, tmp As Variant, r1 As Long, r2 As Long
    For i = 0 To 2
        r1 = A * 3 + i: r2 = B * 3 + i
        tmp = puz(r1): puz(r1) = puz(r2): puz(r2) = tmp
        tmp = sol(r1): sol(r1) = sol(r2): sol(r2) = tmp
    Next i
End Sub

Private Sub SwapColStacks(ByRef puz As Variant, ByRef sol As Variant, ByVal A As Long, ByVal B As Long)
    If A = B Then Exit Sub
    Dim r As Long, i As Long, c1 As Long, c2 As Long, t
    For r = 0 To 8
        For i = 0 To 2
            c1 = A * 3 + i: c2 = B * 3 + i
            t = puz(r)(c1): puz(r)(c1) = puz(r)(c2): puz(r)(c2) = t
            t = sol(r)(c1): sol(r)(c1) = sol(r)(c2): sol(r)(c2) = t
        Next i
    Next r
End Sub

Public Sub NovoFacil()
    Dim puz As Variant, sol As Variant
    puz = BaseFacil(): sol = SolFacil()
    RandomizePuzzle puz, sol
    CarregarPuzzle puz, sol
End Sub

Public Sub NovoMedio()
    Dim puz As Variant, sol As Variant
    puz = BaseMedio(): sol = SolMedio()
    RandomizePuzzle puz, sol
    CarregarPuzzle puz, sol
End Sub

Public Sub NovoDificil()
    Dim puz As Variant, sol As Variant
    puz = BaseDificil(): sol = SolDificil()
    RandomizePuzzle puz, sol
    CarregarPuzzle puz, sol
End Sub

Private Sub CarregarPuzzle(puz As Variant, sol As Variant)
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long, c As Long
    Dim cell As Range

    DrawBoard ws

    For r = 1 To SIZE
        For c = 1 To SIZE
            Solution(r, c) = sol(r - 1)(c - 1)
            PresetMask(r, c) = (puz(r - 1)(c - 1) <> 0)

            Set cell = ws.Cells(TOP_ROW + r - 1, LEFT_COL + c - 1)
            If PresetMask(r, c) Then
                cell.Value = puz(r - 1)(c - 1)
                cell.Font.Bold = True
                cell.Font.Color = COLOR_PRESET
                cell.Locked = True
            Else
                cell.Value = ""
                cell.Font.Bold = False
                cell.Font.Color = COLOR_USER
                cell.Locked = False
            End If
            cell.Interior.Color = GetBlockColor(r, c)
        Next c
    Next r

    If Not ws.ProtectContents Then
        ws.Protect UserInterfaceOnly:=True, AllowFormattingCells:=True
    End If

    ResetCronometro ws
    StartCronometro
End Sub

' ==============================
' Ações: verificar / corrigir / resetar / dica
' ==============================
Public Sub VerificarSudoku()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long, c As Long, cell As Range, v As Variant
    For r = 1 To SIZE
        For c = 1 To SIZE
            Set cell = ws.Cells(TOP_ROW + r - 1, LEFT_COL + c - 1)
            If Not PresetMask(r, c) Then
                v = cell.Value
                If Len(v) = 0 Then
                    cell.Font.Color = COLOR_USER
                    cell.Interior.Color = GetBlockColor(r, c)
                ElseIf IsNumeric(v) And CLng(v) >= 1 And CLng(v) <= 9 Then
                    If CLng(v) = Solution(r, c) Then
                        cell.Font.Color = COLOR_OK
                        cell.Interior.Color = BG_OK
                    Else
                        cell.Font.Color = COLOR_ERR
                        cell.Interior.Color = BG_ERR
                    End If
                Else
                    cell.Font.Color = COLOR_ERR
                    cell.Interior.Color = BG_ERR
                End If
            End If
        Next c
    Next r
    FinalizarSeCompleto
End Sub

Public Sub CorrigirSudoku()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long, c As Long, cell As Range
    For r = 1 To SIZE
        For c = 1 To SIZE
            If Not PresetMask(r, c) Then
                Set cell = ws.Cells(TOP_ROW + r - 1, LEFT_COL + c - 1)
                cell.Value = Solution(r, c)
                cell.Font.Color = COLOR_OK
                cell.Interior.Color = BG_OK
            End If
        Next c
    Next r
    PauseCronometro
    ws.Range("E2").Value = ws.Range("B2").Text ' salva tempo final
End Sub

Public Sub ResetarJogo()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long, c As Long, cell As Range
    For r = 1 To SIZE
        For c = 1 To SIZE
            Set cell = ws.Cells(TOP_ROW + r - 1, LEFT_COL + c - 1)
            If Not PresetMask(r, c) Then
                cell.ClearContents
                cell.Font.Bold = False
                cell.Font.Color = COLOR_USER
                cell.Interior.Color = GetBlockColor(r, c)
            Else
                cell.Font.Bold = True
                cell.Font.Color = COLOR_PRESET
                cell.Interior.Color = GetBlockColor(r, c)
            End If
        Next c
    Next r
    ResetCronometro ws
    StartCronometro
End Sub

Public Sub DarDica()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long, c As Long, cell As Range
    Dim candidatos As Collection: Set candidatos = New Collection

    For r = 1 To SIZE
        For c = 1 To SIZE
            Set cell = ws.Cells(TOP_ROW + r - 1, LEFT_COL + c - 1)
            If Not PresetMask(r, c) And Len(cell.Value) = 0 Then
                candidatos.Add Array(r, c)
            End If
        Next c
    Next r

    If candidatos.Count = 0 Then
        MsgBox "Não há mais dicas disponíveis!", vbInformation
        Exit Sub
    End If

    Dim idx As Long: idx = Int(Rnd * candidatos.Count) + 1
    Dim pos: pos = candidatos(idx)
    r = pos(0): c = pos(1)

    Set cell = ws.Cells(TOP_ROW + r - 1, LEFT_COL + c - 1)
    cell.Value = Solution(r, c)
    cell.Font.Color = RGB(0, 0, 255)
    cell.Interior.Color = RGB(220, 220, 255)
End Sub

Public Sub FinalizarSeCompleto()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long, c As Long
    Dim cell As Range
    Dim completo As Boolean: completo = True
    Dim correto As Boolean: correto = True

    For r = 1 To SIZE
        For c = 1 To SIZE
            Set cell = ws.Cells(TOP_ROW + r - 1, LEFT_COL + c - 1)

            If Len(cell.Value) = 0 Then
                completo = False
                Exit For
            End If

            If Not PresetMask(r, c) Then
                If Not IsNumeric(cell.Value) Or CLng(cell.Value) <> Solution(r, c) Then
                    correto = False
                    cell.Interior.Color = BG_ERR
                    cell.Font.Color = COLOR_ERR
                Else
                    cell.Interior.Color = BG_OK
                    cell.Font.Color = COLOR_OK
                End If
            End If
        Next c
        If Not completo Then Exit For
    Next r

    If completo Then
        If correto Then
            PauseCronometro
            ws.Range("E2").Value = ws.Range("B2").Text
            MsgBox "Parabéns! Sudoku completo e correto.", vbInformation, "Sucesso"
        Else
            MsgBox "O tabuleiro está completo, mas contém erros. Células incorretas foram destacadas.", vbExclamation, "Atenção"
        End If
    End If
End Sub

' ==============================
' Cronômetro: controle de tempo
' ==============================
Public Sub ToggleCronometro()
    If IsTimerRunning Then
        PauseCronometro
    Else
        StartCronometro
    End If
End Sub

Public Sub StartCronometro()
    If IsTimerRunning Then Exit Sub
    StartTime = Now
    IsTimerRunning = True
    ScheduleTick
End Sub

Public Sub PauseCronometro()
    If Not IsTimerRunning Then Exit Sub
    ElapsedBefore = ElapsedBefore + (Now - StartTime) * 86400#
    IsTimerRunning = False
    On Error Resume Next
    Application.OnTime EarliestTime:=NextTick, Procedure:="TickCronometro", Schedule:=False
    On Error GoTo 0
End Sub

Private Sub ScheduleTick()
    NextTick = Now + TimeSerial(0, 0, 1)
    Application.OnTime EarliestTime:=NextTick, Procedure:="TickCronometro", Schedule:=True
End Sub

Public Sub TickCronometro()
    If Not IsTimerRunning Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim elapsed As Double
    elapsed = ElapsedBefore + (Now - StartTime) * 86400#
    With ws.Range("B2:D2")
        .NumberFormat = "[h]:mm:ss"
        ws.Range("B2").Value = elapsed / 86400#
    End With
    ScheduleTick
End Sub

Private Sub ResetCronometro(ws As Worksheet)
    IsTimerRunning = False
    ElapsedBefore = 0
    With ws.Range("B2:D2")
        .NumberFormat = "[h]:mm:ss"
        ws.Range("B2").Value = 0
    End With
    ws.Range("E2").Value = ""
    On Error Resume Next
    Application.OnTime EarliestTime:=NextTick, Procedure:="TickCronometro", Schedule:=False
    On Error GoTo 0
End Sub
