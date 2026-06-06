' =====================================================================
' 1. Option ステートメント
' =====================================================================
Option Explicit On   ' すべての変数宣言を必須にする
Option Strict On     ' 暗黙的な型変換を禁止し、厳密な型チェックを行う
Option Compare Text  ' 文字列の比較を大文字・小文字を区別せずに行う

' =====================================================================
' 2. Imports ステートメント
' =====================================================================
Imports System
' ★作成したVbNetCodeTemplate名前空間をインポートすることで、
' クラス名（OrderManagerなど）をそのまま短く書けるようになります。
'Imports VbNetCodeTemplate

' =====================================================================
' 3. 実行メインモジュール
' =====================================================================
Module Module1

    Sub Main()
        Console.WriteLine("=== VbNetCodeTemplate 利用サンプルの開始 ===")
        Console.WriteLine()

        ' -----------------------------------------------------------------
        ' パターン①：Shared（共有）メソッドの呼び出し
        ' -----------------------------------------------------------------
        ' Sharedメソッドは New（インスタンス化）が不要！クラス名から直接呼び出せます。
        Dim orderCode As String = "ORD-20260606"
        Dim isValid As Boolean = OrderManager.IsValidOrderFormat(orderCode)

        Console.WriteLine($"【Sharedメソッド呼出】")
        Console.WriteLine($"コード '{orderCode}' のフォーマット検証結果: {isValid}")
        Console.WriteLine()


        ' -----------------------------------------------------------------
        ' パターン②：クラスのインスタンス化 と メソッド・プロパティの利用
        ' -----------------------------------------------------------------
        ' 1. コンストラクタ（New）に変数を渡してインスタンスを作成（ヒープ領域へ確保）
        Dim dummyConnectionString As String = "Server=myServerAddress;Database=myDataBase;"
        Dim manager As New OrderManager(dummyConnectionString)

        Console.WriteLine($"【インスタンスメソッド呼出】")

        ' 2. Try-Catchによる安全な処理の実行
        Try
            ' 引数2つのオーバーロードメソッドを呼び出す
            ' （内部でPriceAmount構造体の生成や、例外処理のTry-Catchが走ります）
            manager.ProcessOrder(1001, "太郎")

            ' プロパティのゲッター（Get）を呼び出して値を出力
            Console.WriteLine($"注文処理が正常に終了しました。")
            Console.WriteLine($"注文ID: {manager.OrderId}")
            Console.WriteLine($"現在のステータス: {manager.Status}")

            ' 3. プロパティ経由で取得した構造体（値型）のデータ出力
            Dim price As PriceAmount = manager.TotalPrice
            Console.WriteLine($"合計金額: {price.Amount} {price.Currency}")

        Catch ex As ApplicationException
            ' OrderManager内部で発生し、ラップされて再スローされた業務例外をキャッチ
            Console.WriteLine($"【エラー検知】業務処理でエラーが発生しました。")
            Console.WriteLine($"メッセージ: {ex.Message}")

            ' InnerException（包まれている元の生の例外）が存在する場合はその情報も出力
            If ex.InnerException IsNot Nothing Then
                Console.WriteLine($"根本原因: {ex.InnerException.Message}")
            End If
        End Try

        Console.WriteLine()
        Console.WriteLine("=== サンプルの終了（Enterキーを押して終了） ===")
        Console.ReadLine()
    End Sub

End Module