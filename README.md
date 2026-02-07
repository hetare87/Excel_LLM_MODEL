# Excel_LLM_MODEL
ExcelでLLMっぽいものを作りました。

```VBA
Option Explicit

' ============================================================
' Excel LLM "GOD" Mode (Ultimate Dynamic Edition)
'
' 作成者: Ask AI
' 概要: LLMの脳内（シナプス結合・Transformer構造）を
'       ダイナミックに可視化し、膨大な語彙から次の一手を予測する
' ============================================================

' ------------------------------------------------------------
' Windows API (Sleep用) - 32bit/64bit両対応
' ------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ------------------------------------------------------------
' 設定・定数
' ------------------------------------------------------------
Private Const SHEET_NAME As String = "LLM_God_Mode"

' LLMパラメータ (前回の編成を踏襲)
Private Const MAX_TOKENS As Long = 14       ' 入力トークン最大表示数
Private Const HIDDEN_NODES As Long = 12     ' 隠れ層（脳内）ノード数
Private Const CANDIDATE_COUNT As Long = 10  ' 最終候補として表示する数

' 描画座標
Private Const LAYER_IN_X As Double = 60     ' 入力層X
Private Const LAYER_OUT_X As Double = 800   ' 出力層X
Private Const BASE_Y As Double = 100        ' 基準Y
Private Const NODE_SIZE As Double = 25      ' ノードサイズ

' 脳マップ
Private Const BRAIN_CX As Double = 450      ' 脳図の中心X
Private Const BRAIN_CY As Double = 300      ' 脳図の中心Y
Private Const BRAIN_R As Double = 140       ' 脳図の半径

' 色定義
Private Const COL_BG As Long = 1315860      ' 背景黒 (RGB(20,20,20))
Private Const COL_TEXT As Long = 65280      ' 文字緑 (RGB(0,255,0))
Private Const COL_FIRE As Long = 65535      ' 発火黄色
Private Const COL_PULSE As Long = 16776960  ' パルス水色 (Cyan)

' ============================================================
' メイン処理
' ============================================================
Public Sub LLM_God_Mode_Ultimate_V2_V2()
    ' 後方互換性用：新しいエントリ名に合わせて実行
    LLM_God_Mode_Ultimate_V2
End Sub

Private Sub LogConsole(ws As Worksheet, msg As String)
    ' ログをスクロールさせる
    ws.Range("A6:A24").Value = ws.Range("A5:A23").Value
    ws.Range("A5").Value = "> " & Format(Now, "hh:mm:ss") & " : " & msg
    ws.Range("A5").Font.color = COL_TEXT
    ws.Range("A6:A25").Font.color = rgb(0, 120, 0)
    DoEvents
    Sleep 20
End Sub

Public Sub LLM_God_Mode_Ultimate_V2()
    Dim ws As Worksheet
    Set ws = InitialSetup()
    
    ' 1. 入力取得 (B2セル) - 絶対に上書きしないロジック
    Dim inputText As String
    inputText = CStr(ws.Range("B2").Value) ' そのまま取得
    
    ' 空欄またはスペースのみの場合のみデフォルト値をセット
    If Trim$(inputText) = "" Then
        inputText = "AIは電気羊の夢を見る"
        ws.Range("B2").Value = inputText
    End If
    
    ' 入力をトリムして使用
    inputText = Trim$(inputText)
    
    Application.ScreenUpdating = True
    
    ' コンソール起動
    LogConsole ws, "SYSTEM BOOT SEQUENCE..."
    Sleep 50
    LogConsole ws, "INITIALIZING NEURAL ENGINE..."
    Sleep 50
    
    ' 2. トークン化
    LogConsole ws, "TOKENIZING INPUT: [" & Left(inputText, 20) & "...]"
    Dim tokens() As String
    tokens = Tokenize(inputText)
    
    Dim T As Long: T = UBound(tokens)
    If T > MAX_TOKENS Then T = MAX_TOKENS
    ReDim Preserve tokens(1 To T)
    
    ' 3. Embedding (マトリックス・レイン高速版)
    LogConsole ws, "GENERATING EMBEDDINGS..."
    DrawMatrixRain_Fast ws, T, 8
    
    ' 4. ネットワーク構築
    LogConsole ws, "BUILDING NEURAL ARCHITECTURE..."
    
    Dim inNodes() As Shape
    Dim brainNodes() As Shape
    Dim outNodes() As Shape
    
    ReDim inNodes(1 To T)
    ReDim brainNodes(1 To HIDDEN_NODES)
    ReDim outNodes(1 To CANDIDATE_COUNT)
    
    ' 入力層描画
    BuildInputLayer ws, tokens, inNodes
    
    ' シナプス脳（円形トポロジー）描画
    BuildBrainTopology ws, brainNodes
    
    ' 出力層（候補単語）の準備
    Dim candidateWords() As String
    candidateWords = GenerateCandidates(CANDIDATE_COUNT)
    BuildOutputLayer ws, candidateWords, outNodes
    
    ' 5. フォワードパス
    LogConsole ws, "TRANSMITTING TO HIDDEN LAYERS..."
    AnimateConnectionPulse ws, inNodes, brainNodes, rgb(0, 255, 255)
    
    LogConsole ws, "MULTI-HEAD ATTENTION FIRING..."
    AnimateBrainFiring ws, brainNodes
    
    LogConsole ws, "DECODING PROBABILITIES..."
    AnimateConnectionPulse ws, brainNodes, outNodes, rgb(255, 100, 100)
    
    LogConsole ws, "CALCULATING SOFTMAX DISTRIBUTION..."
    Dim finalWord As String
    finalWord = AnimateSoftmaxCompetition(ws, outNodes)
    
    ' 7. フィナーレ
    LogConsole ws, "GENERATION COMPLETE."
    FinalEffect ws, finalWord
    
    ' 最後に入力セルを選択（任意）
    ws.Range("B2").Select
    
    ' ここで従来の MsgBox は不要とのことなので削除
End Sub

' ============================================================
' 初期設定・コンソール
' ============================================================
Private Function InitialSetup() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SHEET_NAME
    End If
    
    ' 入力セルの値を退避
    Dim currentVal As Variant
    currentVal = ws.Range("B2").Value
    
    ' 画面リセット
    With ws.Cells
        .Clear
        .Interior.color = COL_BG
        .Font.Name = "Consolas"
        .Font.color = rgb(150, 150, 150)
    End With
    
    ' タイトル
    With ws.Range("A1")
        .Value = "/// EXCEL LLM KERNEL : GOD MODE ///"
        .Font.Bold = True
        .Font.Size = 16
        .Font.color = COL_TEXT
    End With
    
    ws.Range("A2").Value = "PROMPT >"
    
    ' 入力セル(B2) の復元と装飾
    With ws.Range("B2")
        .Interior.color = rgb(40, 40, 40)
        .Font.color = vbWhite
        .Font.Size = 12
        .Borders.LineStyle = xlContinuous
        .Borders.color = rgb(100, 100, 100)
        .Value = currentVal
    End With
    
    ' コンソール枠
    With ws.Range("A5:F25")
        .Interior.color = rgb(10, 10, 10)
        .Borders.LineStyle = xlContinuous
        .Borders.color = rgb(0, 80, 0)
    End With
    ws.Range("A5").Value = "--- SYSTEM LOG ---"
    
    ' 既存図形全削除
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    ' 実行ボタンの配置（初回のみ作成）
    EnsureRunButton ws
    
    Set InitialSetup = ws
End Function

Private Sub EnsureRunButton(ws As Worksheet)
    On Error Resume Next
    Dim btn As Variant
    Dim btnLeft As Double, btnTop As Double
    btnLeft = 20
    btnTop = 550
    
    ' すでにボタンがあれば再配置のみ
    If ws.OptionButtons.count > 0 Or ws.Shapes.count > 0 Then
        Dim sh As Shape
        For Each sh In ws.Shapes
            If sh.Name = "GOD_RUN_BTN" Then
                sh.Left = btnLeft: sh.Top = btnTop
                Exit Sub
            End If
        Next sh
    End If
    
    ' 実行ボタンを新規作成
    Set btn = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False, _
        DisplayAsIcon:=False, Left:=btnLeft, Top:=btnTop, Width:=120, Height:=40).Object
    btn.Name = "GOD_RUN_BTN"
    btn.Caption = "GOD MODE 実行"
    ' ボタンクリック時のイベントへ紐づけ（Applicationのイベントでハンドルする形）
    ' ここでは直接イベントを設定する方法はいくつかありますが、VBA標準ではモジュールにのみイベントコードを置くのが安全です。
    ' 実装を簡略化するため、ボタンに対応するマクロを割り当てます。
    ws.OLEObjects("GOD_RUN_BTN").Object.OnClick = "GOD_RUN_BUTTON_CLICK"
    On Error GoTo 0
End Sub

' ボタンのクリック時処理
Public Sub GOD_RUN_BUTTON_CLICK()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    LLM_God_Mode_Ultimate_V2
End Sub

' ============================================================
' ロジック: トークン化 & 候補生成
' ============================================================
Private Function Tokenize(ByVal text As String) As String()
    Dim res() As String
    Dim i As Long
    ReDim res(1 To Len(text))
    For i = 1 To Len(text)
        res(i) = Mid$(text, i, 1)
    Next i
    Tokenize = res
End Function

' 膨大な語彙からランダムに候補を選出
Private Function GenerateCandidates(count As Long) As String()
    Dim vocabList As Variant
    ' 50個程度の単語プール
    vocabList = Array( _
        "未来", "世界", "希望", "絶望", "愛", "平和", "戦争", "宇宙", "時間", "記憶", _
        "真実", "虚構", "意識", "魂", "機械", "人間", "神", "悪魔", "光", "闇", _
        "創造", "破壊", "進化", "退化", "運命", "自由", "支配", "革命", "秩序", "混沌", _
        "夢", "現実", "幻影", "永遠", "瞬間", "生命", "死", "再生", "知性", "感情", _
        "論理", "直感", "言葉", "沈黙", "始まり", "終わり", "答え", "謎", "彼方", "此処" _
    )
    
    Dim res() As String
    ReDim res(1 To count)
    
    Dim i As Long, idx As Long
    Dim used As Object
    Set used = CreateObject("Scripting.Dictionary")
    
    For i = 1 To count
        Do
            idx = Int(UBound(vocabList) * Rnd)
        Loop While used.Exists(idx)
        used(idx) = True
        res(i) = vocabList(idx)
    Next i
    
    GenerateCandidates = res
End Function

' ============================================================
' 演出: Matrix Rain (高速版)
' ============================================================
Private Sub DrawMatrixRain_Fast(ws As Worksheet, T As Long, D As Long)
    Dim r As Long, c As Long
    Dim startR As Long: startR = 5
    Dim startC As Long: startC = 8
    Dim i As Long
    
    ws.Range(ws.Cells(startR, startC), ws.Cells(startR + D, startC + T)).Font.Size = 9
    
    ' ループ回数を減らして高速化
    For i = 1 To 6
        For r = 1 To D
            For c = 1 To T
                Dim val As Double
                val = Rnd()
                With ws.Cells(startR + r, startC + c)
                    .Value = Format(val, "0")
                    If val > 0.7 Then
                        .Font.color = vbWhite
                        .Font.Bold = True
                    Else
                        .Font.color = rgb(0, 180, 0)
                        .Font.Bold = False
                    End If
                End With
            Next c
        Next r
        DoEvents
        Sleep 5
    Next i
    ws.Range(ws.Cells(startR, startC), ws.Cells(startR + D, startC + T)).Interior.color = rgb(0, 30, 0)
End Sub

' ============================================================
' 演出: ネットワーク構築
' ============================================================
Private Sub BuildInputLayer(ws As Worksheet, tokens() As String, inNodes() As Shape)
    Dim i As Long
    Dim gapY As Double: gapY = 35
    Dim shp As Shape
    
    SafeAddLabel ws, LAYER_IN_X - 20, BASE_Y - 40, "INPUT TOKENS"
    
    For i = 1 To UBound(tokens)
        Set shp = ws.Shapes.AddShape(msoShapeOval, LAYER_IN_X, BASE_Y + (i - 1) * gapY, NODE_SIZE, NODE_SIZE)
        SetupNode shp, rgb(0, 100, 200), tokens(i)
        Set inNodes(i) = shp
        DoEvents
    Next i
End Sub

Private Sub BuildBrainTopology(ws As Worksheet, brainNodes() As Shape)
    Dim i As Long, j As Long
    Dim angle As Double
    Dim shp As Shape
    Dim cx As Double, cy As Double
    
    SafeAddLabel ws, BRAIN_CX - 60, BRAIN_CY - BRAIN_R - 40, "HIDDEN LAYERS (TRANSFORMER)"
    
    ' 円形配置
    For i = 1 To HIDDEN_NODES
        angle = -1.57 + (2 * 3.14159 * (i - 1) / HIDDEN_NODES)
        cx = BRAIN_CX + BRAIN_R * Cos(angle)
        cy = BRAIN_CY + BRAIN_R * Sin(angle)
        
        Set shp = ws.Shapes.AddShape(msoShapeOval, cx - NODE_SIZE / 2, cy - NODE_SIZE / 2, NODE_SIZE, NODE_SIZE)
        SetupNode shp, rgb(60, 60, 60), ""
        Set brainNodes(i) = shp
    Next i
    
    ' 複雑な結合
    For i = 1 To HIDDEN_NODES
        For j = i + 1 To HIDDEN_NODES
            If Rnd() > 0.3 Then
                Dim ln As Shape
                Set ln = ws.Shapes.AddLine(brainNodes(i).Left + NODE_SIZE / 2, brainNodes(i).Top + NODE_SIZE / 2, _
                                           brainNodes(j).Left + NODE_SIZE / 2, brainNodes(j).Top + NODE_SIZE / 2)
                ln.Line.ForeColor.rgb = rgb(80, 80, 80)
                ln.Line.Transparency = 0.7
                ln.ZOrder msoSendToBack
            End If
        Next j
    Next i
End Sub

Private Sub BuildOutputLayer(ws As Worksheet, candidates() As String, outNodes() As Shape)
    Dim i As Long
    Dim gapY As Double: gapY = 40
    Dim shp As Shape
    
    SafeAddLabel ws, LAYER_OUT_X - 20, BASE_Y - 40, "NEXT TOKEN PROBABILITY"
    
    For i = 1 To UBound(candidates)
        Set shp = ws.Shapes.AddShape(msoShapeRectangle, LAYER_OUT_X, BASE_Y + (i - 1) * gapY, 20, 25)
        SetupNode shp, rgb(40, 40, 40), candidates(i)
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        shp.TextFrame2.MarginLeft = 5
        Set outNodes(i) = shp
        DoEvents
    Next i
End Sub

' ============================================================
' 演出: アニメーション群
' ============================================================

Private Sub AnimateConnectionPulse(ws As Worksheet, fromNodes() As Shape, toNodes() As Shape, colorRGB As Long)
    Dim i As Long, j As Long
    Dim u1 As Long, u2 As Long
    u1 = UBound(fromNodes)
    u2 = UBound(toNodes)
    
    Dim linksCount As Long: linksCount = 15
    Dim k As Long
    
    For k = 1 To linksCount
        Dim idxFrom As Long: idxFrom = Int(u1 * Rnd) + 1
        Dim idxTo As Long: idxTo = Int(u2 * Rnd) + 1
        
        On Error Resume Next
        Dim x1 As Double, y1 As Double
        Dim x2 As Double, y2 As Double
        
        x1 = fromNodes(idxFrom).Left + fromNodes(idxFrom).Width
        y1 = fromNodes(idxFrom).Top + fromNodes(idxFrom).Height / 2
        x2 = toNodes(idxTo).Left
        y2 = toNodes(idxTo).Top + toNodes(idxTo).Height / 2
        
        Dim ln As Shape
        Set ln = ws.Shapes.AddLine(x1, y1, x2, y2)
        ln.Line.ForeColor.rgb = colorRGB
        ln.Line.Transparency = 0.4
        ln.Line.weight = 2
        ln.ZOrder msoSendToBack
        
        Dim p As Shape
        Set p = ws.Shapes.AddShape(msoShapeOval, x1 - 4, y1 - 4, 8, 8)
        p.Fill.ForeColor.rgb = vbWhite
        p.Line.Visible = msoFalse
        p.Glow.radius = 8
        p.Glow.color.rgb = colorRGB
        
        Dim step As Long
        Dim dx As Double, dy As Double
        dx = (x2 - x1) / 5
        dy = (y2 - y1) / 5
        
        Dim s As Long
        For s = 1 To 5
            p.Left = p.Left + dx
            p.Top = p.Top + dy
            DoEvents
        Next s
        
        p.Delete
        ln.Delete
        On Error GoTo 0
        
        DoEvents
    Next k
End Sub

Private Sub AnimateBrainFiring(ws As Worksheet, brainNodes() As Shape)
    Dim i As Long, loopCnt As Long
    Dim shp As Shape
    
    For loopCnt = 1 To 15
        i = Int((UBound(brainNodes) * Rnd) + 1)
        Set shp = brainNodes(i)
        
        On Error Resume Next
        Dim origColor As Long
        origColor = shp.Fill.ForeColor.rgb
        
        shp.Fill.ForeColor.rgb = COL_FIRE
        shp.Glow.radius = 18
        shp.Glow.color.rgb = COL_FIRE
        shp.Width = NODE_SIZE * 1.6
        shp.Height = NODE_SIZE * 1.6
        shp.Left = shp.Left - (NODE_SIZE * 0.3)
        shp.Top = shp.Top - (NODE_SIZE * 0.3)
        
        ' Attention連携の薄光
        Dim neighbor As Long
        neighbor = (i Mod UBound(brainNodes)) + 1
        brainNodes(neighbor).Glow.radius = 10
        brainNodes(neighbor).Glow.color.rgb = rgb(255, 100, 0)
        
        DoEvents
        Sleep 15
        
        shp.Width = NODE_SIZE
        shp.Height = NODE_SIZE
        shp.Left = shp.Left + (NODE_SIZE * 0.3)
        shp.Top = shp.Top + (NODE_SIZE * 0.3)
        shp.Fill.ForeColor.rgb = origColor
        shp.Glow.radius = 0
        brainNodes(neighbor).Glow.radius = 0
        On Error GoTo 0
    Next loopCnt
End Sub

' Softmax競争 (バーが伸び縮みして競う)
Private Function AnimateSoftmaxCompetition(ws As Worksheet, outNodes() As Shape) As String
    Dim i As Long, k As Long
    Dim shp As Shape
    Dim maxIdx As Long
    Dim scores() As Double
    Dim n As Long: n = UBound(outNodes)
    ReDim scores(1 To n)
    
    For k = 1 To 15
        Dim maxVal As Double: maxVal = -1
        
        For i = 1 To n
            Set shp = outNodes(i)
            Dim noise As Double
            noise = (Rnd() - 0.5) * 0.3
            scores(i) = scores(i) + Rnd() * 0.2 + noise
            If scores(i) < 0.1 Then scores(i) = 0.1
            
            If k > 10 And i = 1 Then scores(i) = scores(i) + 0.3
            
            If scores(i) > maxVal Then
                maxVal = scores(i)
                maxIdx = i
            End If
            
            Dim barWidth As Double
            barWidth = 50 + scores(i) * 150
            If barWidth > 300 Then barWidth = 300
            shp.Width = barWidth
            
            Dim redComp As Long
            redComp = Int(255 * (scores(i) / 2))
            If redComp > 255 Then redComp = 255
            shp.Fill.ForeColor.rgb = rgb(redComp, 50, 50)
            
            ' 表示テキストを更新（候補語を含むように再設定しておく）
            shp.TextFrame2.TextRange.text = outNodes(i).Name & " " & Format(scores(i) * 10, "0.0") & "%"
            On Error GoTo 0
        Next i
        
        DoEvents
        Sleep 10
    Next k
    
    ' 勝者決定
    On Error Resume Next
    maxIdx = Int(n * Rnd) + 1
    Dim shpWin As Shape
    Set shpWin = outNodes(maxIdx)
    shpWin.Fill.ForeColor.rgb = rgb(0, 255, 0)
    shpWin.Glow.radius = 25
    shpWin.Glow.color.rgb = rgb(0, 255, 0)
    shpWin.Width = 350
    shpWin.TextFrame2.TextRange.Font.Bold = msoTrue
    
    AnimateSoftmaxCompetition = shpWin.TextFrame2.TextRange.text
    On Error GoTo 0
End Function

' 最終エフェクト
Private Sub FinalEffect(ws As Worksheet, word As String)
    On Error Resume Next
    Dim res As Shape
    Set res = ws.Shapes.AddShape(msoShapeRoundedRectangle, BRAIN_CX - 200, BRAIN_CY - 80, 400, 160)
    
    res.Fill.ForeColor.rgb = vbBlack
    res.Line.ForeColor.rgb = COL_TEXT
    res.Line.weight = 4
    res.Fill.Transparency = 0.1
    
    res.TextFrame2.TextRange.text = "PREDICTED TOKEN:" & vbCrLf & word
    res.TextFrame2.TextRange.Font.Size = 36
    res.TextFrame2.TextRange.Font.Name = "Consolas"
    res.TextFrame2.TextRange.Font.Fill.ForeColor.rgb = COL_TEXT
    res.TextFrame2.VerticalAnchor = msoAnchorMiddle
    res.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    
    ' インパクト振動
    Dim i As Long
    Dim baseL As Double, baseT As Double
    baseL = res.Left
    baseT = res.Top
    
    For i = 1 To 10
        res.Left = baseL + (Rnd() * 10 - 5)
        res.Top = baseT + (Rnd() * 10 - 5)
        res.Glow.radius = 10 + i * 2
        res.Glow.color.rgb = COL_TEXT
        DoEvents
        Sleep 10
    Next i
    
    res.Left = baseL
    res.Top = baseT
    On Error GoTo 0
End Sub

' ============================================================
' ヘルパー関数
' ============================================================
Private Sub SetupNode(shp As Shape, colorRGB As Long, txt As String)
    On Error Resume Next
    shp.Fill.ForeColor.rgb = colorRGB
    shp.Line.ForeColor.rgb = rgb(200, 200, 200)
    shp.Line.weight = 1
    
    If Len(txt) > 0 Then
        shp.TextFrame2.TextRange.text = txt
        shp.TextFrame2.TextRange.Font.Size = 10
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.rgb = vbWhite
        shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        shp.Name = txt
    End If
    On Error GoTo 0
End Sub

Private Sub SafeAddLabel(ws As Worksheet, x As Double, y As Double, txt As String)
    On Error Resume Next
    Dim shp As Shape
    Set shp = ws.Shapes.AddLabel(msoTextOrientationHorizontal, x, y, 200, 30)
    shp.TextFrame.TextRange.text = txt
    shp.TextFrame.TextRange.Font.color = COL_TEXT
    shp.TextFrame.TextRange.Font.Size = 9
    shp.TextFrame.TextRange.Font.Bold = msoTrue
    On Error GoTo 0
End Sub

' ============================================================
' 追加: 出力検証の表示を I 列第6行までの値を小数点2桁に整形
' ============================================================
Private Sub UpdateVerificationDisplay(ws As Worksheet, values As Variant)
    ' values は I列の検証データを格納する配列。ここでは仮実装として ws.Cells(row, 9).Value から取得する想定
    Dim row As Long
    For row = 6 To 6 ' I6 のみを対象とする例
        If Not IsEmpty(ws.Cells(row, 9).Value) Then
            Dim v As Variant
            v = ws.Cells(row, 9).Value
            If IsNumeric(v) Then
                ws.Cells(row, 9).NumberFormat = "0.00"
                ws.Cells(row, 9).Value = CDbl(v)
            End If
        End If
    Next row
End Sub



