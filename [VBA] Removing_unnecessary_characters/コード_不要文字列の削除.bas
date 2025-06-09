Sub confirmation()

    rc = MsgBox("処理を実行しますか？", vbOKCancel + vbQuestion, "確認")
    
    'キャンセル⇒処理終了
    If rc = vbCancel Then
        
        MsgBox "処理を終了します"
        
        End

    End If
    
    Call main
    
End Sub


'/**
' * @Sub main
' * @description
' * このVBAプロジェクトの主要な処理を実行します。
' * ユーザーが選択したExcelファイル内の特定のシートを対象に、
' * データの品質チェック（不要な文字列の検出とクリーンアップ）、
' * およびフラグ付け（修正有無、修正内容、修正日）を行います。
' * その後、フラグが立てられたデータを別の「修正済データ一覧」シートに転記します。
' *
' * @param {void}
' * @returns {void}
' *
' * @example
' * 'confirmation' プロシージャから呼び出されます。
' * Sub confirmation()
' * ' ... (確認ダイアログ表示)
' * Call main ' このように呼び出されます
' * End Sub
' */
Sub main()

    Dim wb As Workbook
    Dim arrayWs(5)  As Worksheet
    
    Dim sheetName1 As String, sheetName2 As String, sheetName3 As String, sheetName4 As String, sheetName5 As String
    
    Dim lastCol As Integer   '最終列
    Dim rangeHeader As range 'ヘッダーの範囲
    Dim colNumForOutput As Integer
    
    Dim today As Date
    Dim rangeCorrectionOrNot As range
    
    Dim i As Long, j As Long, k As Long
    Dim arryColNum() As Variant
    
    '今日の日付
    today = Date
    
    Set wb = getFileTofalg
    
    'シート名
    sheetName1 = "企・店コード・CPIDをキーにして削除"
    sheetName2 = "MIDをキーにして削除"
    sheetName3 = "決済用CPIDをキーにして削除"
    sheetName4 = "IPIDをキーにして削除"
    sheetName5 = "決済用CPID・IPIDをキーにして削除"
    
    'ワークシートオブジェクトをセット
    With wb
    
        Set arrayWs(1) = .Worksheets(sheetName1)
        Set arrayWs(2) = .Worksheets(sheetName2)
        Set arrayWs(3) = .Worksheets(sheetName3)
        Set arrayWs(4) = .Worksheets(sheetName4)
        Set arrayWs(5) = .Worksheets(sheetName5)
    
    End With
    
    
    '作業対象のシート数をループ処理
    For i = 1 To 5

        'ヘッダー範囲を取得する関数を呼出
        Set rangeHeader = getRangeOfHeader(arrayWs(i))
    
        '出力先の列を取得
'        Set objColNumForOutput = rangeHeader.Find("修正有無", lookat:=xlWhole)
        
        Set rangeCorrectionOrNot = checkToExistColumnOfCorrectionOrNot(rangeHeader)
        
        '「修正有無」カラムが無い場合
        If rangeCorrectionOrNot Is Nothing Then
            
            'フラグを付与する列を挿入し、列順を取得
            colNumForOutput = insertCellsForFlagsAndGetColNum(arrayWs(i))
            
        '「修正有無」カラムがある場合
        Else
            
            '出力先の列順を取得
            colNumForOutput = rangeCorrectionOrNot.Column

        End If
        
        '検索対象のカラムを取得（配列）
        arryColNum = getColNumAtEverySheet(arrayWs(i), sheetName1, sheetName2, sheetName3, sheetName4, sheetName5)
        
        'フラグを付与する関数を実行
        Call flagUnexpectedCharacters(today, arrayWs(i), arryColNum, colNumForOutput, rangeHeader)
    
        'データ転記
        Call transferData(arrayWs(i), colNumForOutput)
        
    Next i
    
    MsgBox "不要文字列を含むデータに対しフラグを付与しました。"

End Sub


'カラム：修正有無の有無をチェック
Function checkToExistColumnOfCorrectionOrNot(range As range) As Variant

    Dim result As range
    
    Set result = range.Find("修正有無", lookat:=xlWhole)
    Set checkToExistColumnOfCorrectionOrNot = result
    
End Function


'/**
' * @Sub transferData
' * @description
' * このプロシージャは、入力ワークシート (wsInput) から「修正有無」列が「○」となっているデータを抽出し、
' * 現在のワークブック内の「修正済データ一覧」シートに転記します。
' * 転記時には、ヘッダーを基準に対応する列にデータを配置し、
' * 転記されたデータは文字列として書式設定されます。
' *
' * @param {Worksheet} wsInput - データを転記する元のワークシートオブジェクト。
' * @param {Integer} colNumOfCorrectionOrNotInInput - wsInput における「修正有無」列の列番号。
' * @returns {void}
' *
' * @example
' * 'main' プロシージャ内でシートごとに呼び出されます。
' * Call transferData(arrayWs(i), colNumForOutput)
' */
Sub transferData(wsInput As Worksheet, colNumOfCorrectionOrNotInInput As Integer)


    Dim wsTarget As Worksheet
    Dim sourceRange As range
    Dim targetRange As range
    Dim inputHeaders As range
    Dim outputHeaders As range
    
    Dim rangeCorrectionOrNotInOutput As range
    Dim colNumOfCorrectionOrNotInOutput As Integer
    
    
    Dim header As String
    Dim cell As range
    Dim outputCol As Long
    Dim lastRowForOutput As Long
    Dim lastColForOutput As Integer
    Dim i As Long
    Dim rowToOutput As Long
    Dim colToOutput As Integer
    Dim outputData() As Variant

' データ取得 ================================

    ' 転記[先] OUTPUT -----
    
        ' ワークシートオブジェクト
        Set wsOutput = ThisWorkbook.Sheets("修正済データ一覧")

        '以下3つを取得
        ' ①{Long}    lastRowForOutput - 最終行
        ' ②{Integer} lastRowForOutput - 最終列
        ' ③{Range}   outputHeaders    - ヘッダー
        With wsOutput
        
            lastRowForOutput = .Cells(.Rows.Count, 1).End(xlUp).Row           '①最終行
            lastColForOutput = .Cells(1, .Columns.Count).End(xlToLeft).Column '②最終列
            
            '③ヘッダー
            Set outputHeaders = .range(.Cells(1, 1), .Cells(1, lastColForOutput))
            
        End With

        ' 「修正有無」の列順を範囲で取得
        Set rangeCorrectionOrNotInOutput = checkToExistColumnOfCorrectionOrNot(outputHeaders)

        'ヘッダーに「修正有無」が無い場合は処理終了
        If rangeCorrectionOrNotInOutput Is Nothing Then

            MsgBox "シート：修正済みデータ一覧のヘッダーに「修正有無」カラムがありません。" & Chr(13) & _
                   "「修正有無」「修正内容」「修正日」の3つを記載した上で再度実行してください"

            End
            
        End If

        '修正有無カラムの列順
        colNumOfCorrectionOrNotInOutput = rangeCorrectionOrNotInOutput.Column


    ' 転記[元] INPUT -----

        ' ワークシートオブジェクト
        ' ws - 第1引数として取得
        
        '以下3つを取得
        ' ①{Long}    lastRowForInput - 最終行
        ' ②{Integer} lastRowForInput - 最終列
        ' ③{Range}   inputHeaders    - ヘッダー
        With wsInput
        
            lastRowForInput = .Cells(.Rows.Count, 1).End(xlUp).Row           '①最終行
            lastColForInput = .Cells(1, .Columns.Count).End(xlToLeft).Column '②最終列
            
            '③ヘッダー
            Set inputHeaders = .range(.Cells(1, 1), .Cells(1, lastColForInput))

        End With


' データ読み込み ================================
    
    ' 転記[先] OUTPUT -----
    
        With wsOutput
        
            outputData = .range(.Cells(2, 1), .Cells(2, lastColForOutput)).Value
        
        End With
        
    ' 転記[元] INPUT -----
    
        With wsInput
        
            inputData = .range(.Cells(1, 1), .Cells(lastRowForInput, lastColForInput)).Value
        
         End With
        

    rowToOutput = lastRowForOutput

    ' データ転記
    'inputファイルの最終行までループ
    For i = 1 To UBound(inputData, 1)
    
        If inputData(i, colNumOfCorrectionOrNotInInput) <> "○" Then GoTo NextRow

        ' 出力先の最終行を取得
        rowToOutput = rowToOutput + 1

        With wsOutput
        
            .range(.Cells(rowToOutput, 1), .Cells(rowToOutput, lastColForOutput)).NumberFormat = "@"
        
        End With
        
        ' inputシートの各レコード（1行ごと）をループ
        For j = 1 To UBound(inputData, 2)
        
            ' 転記元シートのカラム名を取得
            If inputHeaders.Cells(1, j) <> "" Then
            
                header = inputHeaders.Cells(1, j)

                ' 転記先シートの対応するカラムを検索
                On Error Resume Next
                
                colToOutput = Application.WorksheetFunction.Match(header, outputHeaders, 0)
                
                On Error GoTo 0
    
                ' 対応するカラムが見つかった場合、データを転記
                If colToOutput > 0 Then
                
    '                outputData(rowToOutput, colToOutput) = inputData(i, j)
                    
                    If inputData(i, j) <> "" Then
                    
                        wsOutput.Cells(rowToOutput, colToOutput) = CStr(inputData(i, j))
                    
                    End If
                End If
            End If

        Next j

        'シート名
        wsOutput.Cells(rowToOutput, 1) = wsInput.Name

NextRow:

    Next i

End Sub



'ダイアログでファイルを開く
Function getFileTofalg()

    Dim FileName As Variant
        
        '作業対象のファイルを選択
        FileName = Application.GetOpenFilename(FileFilter:="Excelファイル,*.xls*")
        
        'キャンセルやescでファイルが選択されなかった場合は処理終了
        If FileName = False Then
            
            MsgBox "処理を終了します"
            
            End
        
        End If
        
        'ファイルを開く
        Workbooks.Open FileName
        
        '開いたファイルをセットして返却
        Set getFileTofalg = ActiveWorkbook

End Function



'/**
' * @Sub flagUnexpectedCharacters
' * @description
' * 指定されたワークシートの特定の列を対象に、セルの値に不要な文字（先頭のシングルクォーテーションや、
' * 列の型に合わない文字）が含まれていないかをチェックします。
' *
' * 不要な文字が検出された場合、以下の通りに記録します。
' *   - その行の「修正有無」列に「○」
' *   - 「修正内容」列に検出された不要文字（可視化された形式）
' *   - 「修正日」列に処理実行日
' *
' * また、検出された不要文字は元のセルから削除され、クリーンなデータが文字列として書き戻されます。
' * 処理後、対象列の書式を文字列に設定し、オートフィルターを適用します。
' *
' * @param {Date} today - 処理実行日。修正日の記録に使用されます。
' * @param {Worksheet} ws - 不要文字列のチェックとフラグ付けを行うワークシートオブジェクト。
' * @param {Variant} arryColNum - チェック対象となる列のRangeオブジェクトの配列。`getColNumAtEverySheet` 関数で取得されます。
' * @param {Integer} colNumForOutput - 「修正有無」「修正内容」「修正日」が出力される列の、「修正有無」列の列番号。
' * @param {Range} rangeHeader - ワークシートのヘッダー範囲。
' *
' * @returns {void}
' *
' * @example
' * 'main' プロシージャ内で、処理対象の各シートに対して呼び出されます。
' * Call flagUnexpectedCharacters(today, arrayWs(i), arryColNum, colNumForOutput, rangeHeader)
' */
Sub flagUnexpectedCharacters(today As Date, ws As Worksheet, arryColNum As Variant, colNumForOutput As Integer, rangeHeader As range)

    Dim lastRow As Long
    Dim inputString As String '検索対象の文字列
    Dim nonAlphanumericChars As String '削除対象の文字列
    Dim currentChar As String '検索対象の文字列を1文字ずつ取得した時の文字列
    
    
    '最終行を取得
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    '最終列までループ処理
    For j = 2 To lastRow
    
        '検索対象のカラム（列）をループ
        For k = LBound(arryColNum) To UBound(arryColNum)
        
            '■変数準備 ==========================
            
                '検索対象の列順を格納
                targetCol = arryColNum(k).Column
            
                '非アルファベット・非数字の文字を格納する変数を初期化
                nonAlphanumericChars = ""
                cleanedString = ""
                
                '現在のセルの値を代入(要素は0始まりのため +1)
                inputString = ws.Cells(j, targetCol)


            '■不要文字列チェック & フラグ付与 ==========================
            
                '□チェック①：文字列の先頭にシングルクォーテーションがあるか否か ----------
                    
                    With ws
                    
                        If .Cells(j, targetCol).PrefixCharacter = "'" Then
                        
                            .Cells(j, colNumForOutput) = "○"
                            .Cells(j, colNumForOutput + 1) = "先頭SQ"
                            .Cells(j, colNumForOutput + 2) = today
                        
                        End If
                    End With
                
                
                '□チェック②：文字列に想定外の文字列が含まれているか否か ----------
                
                    '***** 不要文字列チェック *****
                    
                        'セル内の文字列の文字数だけループ処理
                        For l = 1 To Len(inputString)
                        
                            '各セルから1文字ずつ取得
                            currentChar = Mid(inputString, l, 1)
                            
                            '[1] ヘッダーが「決済用CPID」「IPID」の場合：数字のみ
                            If ws.Cells(1, targetCol) = "決済用CPID" Or ws.Cells(1, targetCol) = "IPID" Then
                            
                                '条件：【整数】
                                If currentChar Like "[0-9]" Then
                                    
                                    cleanedString = cleanedString & currentChar

                                Else
                                
                                    nonAlphanumericChars = nonAlphanumericChars & currentChar
                                    
                                End If
                            
                            '[2] ヘッダーが「決済用CPID」「IPID」【以外】の場合
                            Else
                                
                                '条件：【整数】 or 【文字列】
                                If currentChar Like "[A-Za-z0-9]" Then
                                
                                    cleanedString = cleanedString & currentChar
                                
                                Else
                                
                                    nonAlphanumericChars = nonAlphanumericChars & currentChar
                                    
                                End If
                            End If
                            
                            
                            'オーバーフローになる
                            '原因は、シングルクォーテーション除外後のセルの書式が日付に自動変換されるため
'                            ws.Cells(j, targetCol) = Replace(ws.Cells(j, targetCol), nonAlphanumericChars, "")
                        
                        Next l
                        
                        
                        ' 非英数字を除外した文字列をセルに設定（文字列として）
                        ws.Cells(j, targetCol).Clear
                        ws.Cells(j, targetCol).NumberFormat = "@" ' セルの書式を文字列に設定
                        ws.Cells(j, targetCol).Value = cleanedString

                    
                    '***** フラグ付与 *****
                    
                        '削除文字列が含まれていた場合
                        If nonAlphanumericChars <> "" Then
                            
                            '削除文字列の整備
                            '「'」⇒「SQ」（Single Quotation)・・・「'」単独ではセル上で見えないため可視化
                            nonAlphanumericChars = Replace(nonAlphanumericChars, "'", "SQ")
                            
                            '「 」「　」⇒「空白」・・・「 」「　」はセル上で見えないため可視化
                            nonAlphanumericChars = Replace(nonAlphanumericChars, " ", "空白")
                            nonAlphanumericChars = Replace(nonAlphanumericChars, "　", "空白")
                            nonAlphanumericChars = Replace(nonAlphanumericChars, "    ", "空白") 'この空白が認識されない
                            
                                    
                            'output列に出力
                            With ws
                            
                                .Cells(j, colNumForOutput) = "○"
                                .Cells(j, colNumForOutput + 1) = nonAlphanumericChars
                                .Cells(j, colNumForOutput + 2) = today
                                
                            End With
                        End If
                
        Next k
    Next j
    
    '文字列へ変換（標準のまま先頭のシングルクォーテーションを削除すると、0落ちや指数表記（E+）になるため）
    Call formatToString(ws, arryColNum, lastRow)
    
    'フィルターを設定
    Call autoFilter(ws)


End Sub


'作業中のシートを判別し、該当シートにおける作業対象列の列順を取得
Function getColNumAtEverySheet(ws As Worksheet, sheetName1 As String, sheetName2 As String, sheetName3 As String, sheetName4 As String, sheetName5 As String) As Variant

    Dim sheetName As String
    
    
    sheetName = ws.Name
    
    Select Case sheetName
    
        'シート：企・店コード・CPIDをキーにして削除
        Case sheetName1
            
            '要素数 = 3
            ReDim arrayCol(2)
            
            arrayCol(0) = "企業コード"
            arrayCol(1) = "店舗コード"
            arrayCol(2) = "決済用CPID"
    
        'シート：決済用CPID・IPIDをキーにして削除
        Case sheetName5
        
            '要素数 = 2
            ReDim arrayCol(1)
            
            arrayCol(0) = "決済用CPID"
            arrayCol(1) = "IPID"

        'シート：sheetName2 - MIDをキーにして削除
        'シート：sheetName3 - 決済用CPIDをキーにして削除
        'シート：sheetName4 - IPIDをキーにして削除
        Case Else
            
            '要素数 = 1
            ReDim arrayCol(0)
            
            'シート：MIDをキーにして削除
            If InStr(ws.Name, "MID") > 0 Then
                arrayCol(0) = "マーチャントID"
             
            'シート：決済用CPIDをキーにして削除
            ElseIf InStr(ws.Name, "決済用CPID") > 0 Then
                arrayCol(0) = "決済用CPID"

            'シート：IPIDをキーにして削除
            ElseIf InStr(ws.Name, "IPID") > 0 Then
                arrayCol(0) = "IPID"

            End If

    End Select
    
    arrayColNum = findColNum(ws, arrayCol)

    '列順を格納する配列を返す
    getColNumAtEverySheet = arrayColNum

End Function


'作業対象列の列順を取得
Function findColNum(ws As Worksheet, arrayCol As Variant) As Variant
    
    '要素数を配列に設定
    ReDim arryColNum(UBound(arrayCol))
    
    'ヘッダー範囲を取得する関数を呼出
    Set rangeHeader = getRangeOfHeader(ws)
    
    '作業対象の列数分をループ処理
    For i = LBound(arrayCol) To UBound(arrayCol)

        Set arryColNum(i) = rangeHeader.Find(arrayCol(i), lookat:=xlWhole)

    Next i

    findColNum = arryColNum

End Function


'ヘッダーの範囲を取得
Function getRangeOfHeader(ws As Worksheet) As range

    '最終列
    lastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    'ヘッダーの範囲
    With ws
    
        Set rangeHeader = .range(.Cells(1, 1), .Cells(1, lastCol))

    End With

    Set getRangeOfHeader = rangeHeader
    
End Function



'検索対象範囲を文字列に変換する
Sub formatToString(ws As Worksheet, arrayColNum As Variant, lastRow As Long)

    Dim rangeForFormatingToString As range

        '作業対象の列を範囲で取得
        With ws
        
            Set rangeForFormatingToString = .range(.Cells(1, LBound(arrayColNum) + 1), .Cells(lastRow, UBound(arrayColNum) + 1))
        
        End With
        
        '文字列に変換
        rangeForFormatingToString.NumberFormat = "@"

End Sub



'フラグ付与列を作成
Function insertCellsForFlagsAndGetColNum(ws As Worksheet)

    Dim lastRowForGetLastCol As Long
    Dim lastCol As Long
    Dim currentRow As Long
    Dim currentCol As Long
    Dim rightmostCell As range
    Dim tempCell As range


    '最終行を取得
    lastRowForGetLastCol = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' 初期化
    Set rightmostCell = ws.Cells(1, 1)
    
    ' 各行をループして最右端のセルを特定
    For currentRow = 1 To lastRowForGetLastCol
    
        lastCol = ws.Cells(currentRow, ws.Columns.Count).End(xlToLeft).Column
        
        Set tempCell = ws.Cells(currentRow, lastCol)
        
        ' 現在の最右端のセルより右にある場合、更新
        If tempCell.Column > rightmostCell.Column Then

            Set rightmostCell = tempCell

        End If

    Next currentRow

    'フラグを挿入
    ws.Cells(1, rightmostCell.Column + 1) = "修正有無"
    ws.Cells(1, rightmostCell.Column + 2) = "修正内容"
    ws.Cells(1, rightmostCell.Column + 3) = "修正日"

    '「修正有無」の位置を返却
    insertCellsForFlagsAndGetColNum = rightmostCell.Column + 1

End Function



'フィルターを設定する
Sub autoFilter(ws As Worksheet)

    With ws

        'フィルターが設定されている場合
        If .AutoFilterMode Then
        
            'かつ「いずれかのカラムで絞り込みが実施されている」場合
            If .FilterMode Then
    
                'フィルタをクリアして全て表示（フィルタそのものは設定したまま）
                .ShowAllData
    
            End If
            
            With .range("A1").CurrentRegion
            
                .autoFilter '解除
                .autoFilter '設定
           
            End With
        
        'フィルターが設定されていない場合
        Else
            
            'フィルターを設定
            .range("A1").CurrentRegion.autoFilter
            
        End If
    End With
End Sub
