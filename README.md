# VB.NET（C#） コーディング規約

Microsoft公式の設計ガイドライン（dotnet/docs）に基づき、C#およびVB.NETにおける命名規則のデファクトスタンダードとコードテンプレートを記載

---

## 1. 各構成要素

### ① 前提定義（Option / Imports / Namespace）
* **Option:** ファイルの最上部でコンパイルルールを厳格化します。特に `Option Strict On` は暗黙の型変換によるバグを防ぐため、実務では**必須のデファクトスタンダード**です。
* **Imports / Namespace:** 外部ライブラリを呼び出すための宣言と、自分自身のコードが所属する論理的なフォルダ（階層）を定義します。

### ② 列挙型（Enum）の定義
* **役割:** ステータスコードや区分、区分値など、関連する一連の固定的な整数値を分かりやすい「名前」でまとめて管理する仕組みです。
* **実務でのポイント:** コード内に `0` や `1` といったマジックナンバー（生の数値）が散らばるのを防ぎ、プログラムの可読性とバグの抑制に大きく貢献します。

### ③ 構造体（Structure）の定義と「メモリ格納場所」の違い
* **役割:** 複数の小規模な関連データを1つのオブジェクトとしてパッケージ化する「値型（Value Type）」の仕組みです。
* **クラスとのメモリ格納場所の違い（スタックとヒープ）:**
  クラス（参照型）と構造体（値型）は、どちらもPCの「メモリ（RAM）」に配置されますが、使われる内部領域が決定的に異なります。
  * **構造体（値型） ➡ 主に「スタック領域」に格納:**
    机の端にデータを積み上げていくような、非常に高速なエリアです。構造体が定義されたメソッド（処理）が**終了した瞬間に、跡形もなく自動で一瞬でメモリから消滅（解放）**します。データの出し入れも片付けも超高速なのが特徴です。
  * **クラス（参照型） ➡ 主に「ヒープ領域」に格納:**
    広い倉庫のようなエリアです。メソッドが終了しても自動では消えず、後から **GC（ガベージコレクション）** という専属のお掃除ロボットが定期的に巡回して、使われなくなったデータを回収・解放します。

### ④ 定数（Const）の定義
* **役割:** プログラムの中で**絶対に変わらない固定値**（消費税率やリトライ上限など）に名前をつけて管理する仕組みです。
* **実務でのポイント:** 値は「コンパイル時」に完全固定されます。コードの保守性を高めるために、クラスの先頭付近に集約して定義します。

### ⑤ フィールド（クラス内部用の変数）の定義
* **役割:** クラスの**「内部」だけで使い回す生のデータ箱**です。
* **実務でのポイント:** 外部から直接触らせないよう、アクセス修飾子は必ず **`Private`** にします。また、コンストラクタで初期化した後に値を書き換えられたくない変数（接続文字列など）には **`ReadOnly`** を付与し、実行時の安全性を担保します。

### ⑥ プロパティ（クラス外部向けの安全な窓口）
* **役割:** クラスの**「外」に見せる公式なデータ窓口**です。原則 `Public` で公開します。

* **プロパティ（ゲッター・セッター）の代入と取得の呼び出し方:**
  プロパティの実態は関数（メソッド）ですが、使う側は**「普通の変数」と全く同じ感覚で記述する**ことで、裏側の仕組みが自動的に呼び出されます。
  * **代入時（セッターの呼び出し）：** イコール（`=`）の左側にプロパティ名を書くと、裏側の `Set` ブロックが実行されます。
    ```vb
    manager.OrderId = 1 ' 裏側で Set(ByVal value As Integer) が走り、1 が代入される
    ```
  * **取得時（ゲッターの呼び出し）：** 変数への代入や条件式など、値を取り出す形でプロパティ名を書くと、裏側の `Get` ブロックが実行されます。
    ```vb
    Dim x As Integer = manager.OrderId ' 裏側で Get が走り、現在の値が返される
    ```

* **標準プロパティ（Get/Set明示）:** 値の出し入れ時にチェック処理（バリデーション）や連動ロジックを挟みたい場合に使用します。
* **自動実装プロパティ（1行記述）:** 特別なロジックが不要な場合に使用します。1行書くだけでコンパイラが裏側で「隠しフィールド」と「Get/Setの仕組み」を自動生成してくれます。最初はバリデーションが不要でも、将来の仕様変更が発生した際、**呼び出し側のコードを1ミリも書き換えることなくクラス内部だけで安全に修正できる保険**となるため、値の公開にはフィールドではなく必ずプロパティを使うのが鉄則です。

### ⑦ メソッド と オーバーロード（実際の処理と多重定義）
* **メソッド:** クラスが実行する「処理」を記述する場所です。
* **オーバーロード:** **「名前は同じだけど、引数の数や型が違うメソッドを複数定義できる仕組み」**です。使う側は1つのメソッド名だけを覚えればよく、渡す引数の数や型によってプログラムが自動で最適なパターン（処理）を判断して実行してくれます。

### ⑧ 共有メソッド（Shared）
* **役割:** クラスを `New`（インスタンス化）しなくても、**クラス名から直接呼び出せる共通関数（ユーティリティ）**です。C#の `static` メソッドに該当します。
* **呼び出し方の違い（実務のイメージ）:**
  * **通常のメソッド（Newが必要）：** クラスが持つ固有のデータ（接続文字列など）を使って処理する「オーダーメイドの職人」のイメージ。必ずインスタンスを生成して呼び出します。
    ```vb
    Dim mgr As New OrderManager("Server=...")
    mgr.ProcessOrder(1001)
    ```
  * **Sharedメソッド（Newが不要）：** インスタンスのデータに依存せず、外から渡された引数（データ）だけで処理が完結する「街の自動販売機や計算機」のイメージ。`New` を完全にすっ飛ばして、**クラス名から直接1行で呼び出します。**
    ```vb
    Dim isValid As Boolean = OrderManager.IsValidOrderFormat("ORD-20260606")
    ```
* **実務での使いどころ:** 「状態（データ）を持たず、渡されたデータを加工して返すだけ」の処理（文字列のバリデーション、日付の共通変換、誰が呼んでも同じ結果になるマスタデータの全件取得など）をプロジェクト全体の共通部品（便利関数群）として定義する際によく使用します。

### ⑨ 例外処理（Try-Catch-Finally）
* **役割:** 実行中に予期せぬエラー（データベースの接続切断やファイルの不在など）が発生した際に、プログラムが強制終了するのを防ぎ、安全にエラー処理を行うための仕組みです。
* **実務でのポイント（集中管理とラップ）:**
  各メソッド内でむやみにエラーを「もみ消す（握りつぶす）」のは致命的なバグの元になります。実務の例外処理の標準は以下の通りです。
  * **特定の例外を検知（Catch）:** `SqlException` など原因が明確なエラーを捕捉し、必要に応じてリトライや独自の回復処理を行います。
  * **独自の例外でラップして再スロー（Throw）：** 低レイヤの生の例外をそのまま上位に投げず、業務的な意味を持つ自作の例外（例：`ApplicationException`）に、元の例外を「インナーエクセプション（`InnerException`）」として包んで（ラップして）上に投げ直します。これにより、上位層でエラーを一元管理しやすくなり、かつデバッグ時には発生元の一番深いエラー原因（スタックトレース）を完全に追跡できます。
  * **後片付け（Finally）：** エラーの発生の有無に関わらず、ファイルやデータベース接続などの外部リソースを確実に閉じる（解放する）ために `Finally` ブロックを使用します。

---

## 2. VB.NET（.NET Framework 4.7.2）公式準拠のコードテンプレート

> 💡 **プロジェクト名に関する注意点**
> プロジェクト名を `VbNetCodeTemplate` に設定する場合、Visual Studioのプロジェクトプロパティを開き、**「ルート名前空間」の欄を完全に空欄（クリア）**にしてください。これにより、コード側に記述された `Namespace VbNetCodeTemplate` が名前空間の自動二重化を起こすことなく、正しく適用されます。

```vb
' =====================================================================
' 1. Option ステートメント
' =====================================================================
Option Explicit On   ' すべての変数宣言を必須にする
Option Strict On     ' 暗慢的な型変換を禁止し、厳密な型チェックを行う
Option Compare Text  ' 文字列の比較を大文字・小文字を区別せずに行う

' =====================================================================
' 2. Imports ステートメント
' =====================================================================
Imports System
Imports System.Data.SqlClient

' =====================================================================
' 3. Namespace ステートメント
' =====================================================================
Namespace VbNetCodeTemplate
    
    ''' <summary>
    ''' 注文状態を表す列挙型（PascalCase）
    ''' </summary>
    Public Enum OrderStatus
        NotStarted = 0
        Processing = 1
        Completed = 2
        Failed = 9
    End Enum

    ''' <summary>
    ''' 簡易的な金額データを表す不変の構造体（PascalCase）
    ''' 主にスタック領域に格納され、軽量・高速に処理される
    ''' </summary>
    Public Structure PriceAmount
        Public ReadOnly Amount As Decimal
        Public ReadOnly Currency As String

        Public Sub New(ByVal amount As Decimal, ByVal currency As String)
            Me.Amount = amount
            Me.Currency = currency
        End Sub
    End Structure

    ''' <summary>
    ''' 注文処理を管理するクラス
    ''' 主にヒープ領域に格納され、GC（ガベージコレクション）によって管理される
    ''' </summary>
    Public Class OrderManager
        
        ' -----------------------------------------------------------------
        ' ■ 定数（Const）の定義
        ' -----------------------------------------------------------------
        Public Const MaxRetryCount As Integer = 3
        Public Const DefaultCustomerName As String = "Unknown"
        
        ' -----------------------------------------------------------------
        ' ■ フィールド（Private / クラス内部用）の定義
        ' -----------------------------------------------------------------
        Private ReadOnly _connectionString As String ' 実行時にNewした瞬間だけ値が決まる
        Private _orderId As Integer                  ' 通常のプライベート変数（_camelCase）
        
        ' -----------------------------------------------------------------
        ' ■ コンストラクタ（初期化処理）
        ' -----------------------------------------------------------------
        Public Sub New(ByVal connectionString As String)
            Me._connectionString = connectionString
            Me._orderId = 0
        End Sub
        
        ' -----------------------------------------------------------------
        ' ■ プロパティ（Public / クラス外部公開用窓口）の定義
        ' -----------------------------------------------------------------
        
        ''' <summary>
        ''' 標準的なプロパティ構文（Get / Set を明示的に書く形）
        ''' </summary>
        Public Property OrderId As Integer
            Get
                Return Me._orderId
            End Get
            Set(ByVal value As Integer)
                If value < 0 Then
                    Throw New ArgumentException("注文IDは0以上である必要があります。")
                End If
                Me._orderId = value
            End Set
        End Property
        
        ''' <summary>
        ''' 現在の注文ステータス（自動実装プロパティ）
        ''' </summary>
        Public Property Status As OrderStatus
        
        ''' <summary>
        ''' 注文の合計金額（構造体型プロパティ）
        ''' </summary>
        Public Property TotalPrice As PriceAmount
        
        ' 【参考】上記の「自動実装プロパティ（Status）」を定義した際、
        ' コンパイラが裏側で自動生成してくれているコードのイメージ（コメントアウト）
        ' -----------------------------------------------------------------
        ' Private _status As OrderStatus ' 自動生成される隠しフィールド
        '
        ' Public Property Status As OrderStatus
        '     Get
        '         Return Me._status
        '     End Get
        '     Set(ByVal value As OrderStatus)
        '         Me._status = value
        '     End Set
        ' End Property
        ' -----------------------------------------------------------------
        
        ' -----------------------------------------------------------------
        ' ■ メソッドの定義（通常メソッド & オーバーロード）
        ' -----------------------------------------------------------------
        ' 【参考】オーバーロードする場合、いずれのメソッドにもOverloadsを付ける必要がある
        
        ''' <summary>
        ''' パターン1：注文IDだけで処理する（引数1つ）
        ''' </summary>
        Public Overloads Sub ProcessOrder(ByVal orderId As Integer)
            Me.ProcessOrder(orderId, DefaultCustomerName)
        End Sub
        
        ''' <summary>
        ''' パターン2：顧客名も指定して処理する（引数2つ）
        ''' </summary>
        Public Overloads Sub ProcessOrder(ByVal orderId As Integer, ByVal customerName As String)
            Me.OrderId = orderId
            Me.Status = OrderStatus.Processing 
            
            ' 構造体をNewしてプロパティにセットする実例（スタックに割り当て）
            Me.TotalPrice = New PriceAmount(1500, "JPY")
            
            ' .NET標準のDateTime構造体を使用する実例
            Dim currentTimestamp As DateTime = DateTime.Now
            
            ' -------------------------------------------------------------
            ' ■ 例外処理（Try-Catch-Finally）の実装パターン
            ' -------------------------------------------------------------
            Try
                ' データベース処理を実行
                ExecuteDatabaseQuery(customerName)
                Me.Status = OrderStatus.Completed
                
            Catch ex As SqlException
                ' パターンA：特定の技術例外を検知した場合の処理
                Me.Status = OrderStatus.Failed
                
                ' 実務の鉄則：元の例外（ex）を内部に包んで（ラップして）、上位層へ再スローする
                Throw New ApplicationException("データベース接続エラーのため、注文処理に失敗しました。", ex)
                
            Catch ex As Exception
                ' パターンB：その他の予期せぬ汎用例外を検知した場合
                Me.Status = OrderStatus.Failed
                
                ' NLogやSerilogなどのロギングライブラリでエラーを記録する実務イメージ
                ' Logger.Error("予期せぬシステムエラーが発生しました。", ex)
                
                Throw New ApplicationException("システム内部エラーが発生したため、処理を中断しました。", ex)
                
            Finally
                ' パターンC：エラーの有無に関わらず、必ず最後に通る後片付け処理
                ' （コネクションの切断や、開いたファイルのクローズなど）
            End Try
        End Sub
        
        Private Sub ExecuteDatabaseQuery(ByVal customerName As String)
            ' ダミー処理（実際のSQL接続ロジックが入る）
        End Sub
        
        ' -----------------------------------------------------------------
        ' ■ 共有メソッド（Shared / インスタンス化不要の共通関数）
        ' -----------------------------------------------------------------
        Public Shared Function IsValidOrderFormat(ByVal orderCode As String) As Boolean
            If String.IsNullOrEmpty(orderCode) Then Return False
            Return orderCode.StartsWith("ORD-")
        End Function
        
    End Class
    
End Namespace
```

---

## 3. C# と VB.NET の命名規則デファクトスタンダード比較

| 対象となる要素 | C# の標準ルール | VB.NET の標準ルール | 具体的なコード例 |
| :--- | :--- | :--- | :--- |
| **クラス / 構造体 / Enum** | `PascalCase` | `PascalCase` | `OrderManager` / `PriceAmount` / `OrderStatus` / `DateTime` |
| **メソッド** | `PascalCase` | `PascalCase` | `ProcessOrder()` |
| **プロパティ** | `PascalCase` | `PascalCase` | `OrderId` / `Status` |
| **定数（Const）** | `PascalCase` | `PascalCase` | `MaxRetryCount` |
| **プライベートフィールド**| **`_camelCase`** | **`_camelCase`** | `_connectionString` / `_orderId` |
| **ローカル変数 / 引数** | `camelCase` | `camelCase` | `customerName` / `orderId` / `currentTimestamp` |

### 🔗 情報元（Microsoft公式ガイドラインへのリンク）

上記の命名規則は、以下のMicrosoft公式ドキュメントで推奨されている記述方法に基づいています。

* [C# のコーディング規則（識別子名・スタイル標準）](https://learn.microsoft.com/ja-jp/dotnet/csharp/fundamentals/coding-style/identifier-names)
  * ※「Camel 形式」の項目に、private または internal フィールドには `_camelCase` を使用する旨が明記されています。
* [.NET クラスデザインガイドライン（大文字と小文字の規則）](https://learn.microsoft.com/ja-jp/dotnet/standard/design-guidelines/capitalization-conventions)
  * ※PascalCase / camelCase の明確な定義と使い分けの基本となる根拠ページです。
* [.NET クラスデザインガイドライン（命名規則トップ）](https://learn.microsoft.com/ja-jp/dotnet/standard/design-guidelines/naming-guidelines)
  * ※タイプ、メソッド、プロパティなどの総合的な命名規則が網羅されています。
* [System.DateTime（DateTime.cs）](https://github.com/dotnet/dotnet/blob/b0f34d51fccc69fd334253924abd8d6853fad7aa/src/runtime/src/libraries/System.Private.CoreLib/src/System/DateTime.cs)

<br>

---

# .editorconfig 設定

## 1. .editorconfig の概要
`.editorconfig` は、コードのインデント、改行コード、コードスタイル、および名前付け規則（命名規則）をプロジェクト単位で一元管理し、開発者間で統一するための設定ファイルです。Visual Studio はこのファイルを自動的に読み込み、リアルタイムでの警告表示や、クイックアクション（`Ctrl` + `.`）によるコードの自動修正機能を提供します。

---

## 2. .editorconfig 共通・主要プロパティ一覧

### ① 基本的なコードスタイル（全言語共通）
ファイルの物理的なフォーマットを規定するプロパティです。

* **`root`** (`true` / `false`): 
  このファイルがあるフォルダを最上位（ルート）として宣言します。`true` に設定すると、これより上の階層にある設定ファイルを探しに行くのを停止します。
* **`indent_style`** (`space` / `tab`): 
  インデントにスペースを使用するか、タブ文字を使用するかを指定します。
* **`indent_size`** (整数): 
  インデント1段分に相当するスペースの数を指定します（`indent_style = space` の際に有効）。
* **`tab_width`** (整数): 
  タブ文字1つ分に相当する幅をスペースの数で指定します。
* **`end_of_line`** (`crlf` / `lf` / `cr`): 
  行末の改行コードを指定します。Windows環境における .NET 開発では `crlf` が標準的です。
* **`insert_final_newline`** (`true` / `false`): 
  ファイルの末尾に必ず空行（改行）を1行挿入するかどうかを制御します。

### ② .NET 共通のコードスタイル（言語規則）
クラスメンバーへのアクセス方法や、修飾子の付与基準など、コードの記述スタイルを制御します。
*(※ 各設定値の後ろに `:warning` や `:suggestion` を付与することで、違反時の重要度を制御できます)*

* **`dotnet_style_qualification_for_field`** (`true` / `false`): 
  フィールドへのアクセス時に `Me.` (VB.NET) や `this.` (C#) の記述を必須とするか、省略するかを設定します。
* **`dotnet_style_qualification_for_property`** (`true` / `false`): 
  プロパティへのアクセス時に `Me.` や `this.` の記述を必須とするか、省略するかを設定します。
* **`dotnet_style_qualification_for_method`** (`true` / `false`): 
  メソッド呼び出し時に `Me.` や `this.` の記述を必須とするか、省略するかを設定します。
* **`dotnet_style_readonly_field`** (`true` / `false`): 
  コンストラクタ以外で変更されないフィールドに対し、`ReadOnly`（C#は `readonly`）修飾子の付与を推奨するかどうかを制御します。
* **`dotnet_style_require_accessibility_modifiers`** (`always` / `never` / `for_non_interface_members`): 
  `Private` や `Public` などのアクセス修飾子を明示的に記述させるかどうかを設定します。

### ③ 命名規則（Naming Styles）に関するカスタムプロパティ
独自のルールを組み立てるための仕組みであり、以下の3つの要素を組み合わせて定義します。

* **`dotnet_naming_rule.[ルール名].[symbols / style / severity]`**: 
  特定のシンボル（対象）に、どのスタイルと重要度を紐づけるかを定義するプロパティです。
* **`dotnet_naming_symbols.[シンボル名].[applicable_kinds / applicable_accessibilities / required_modifiers]`**: 
  ルールを適用したい要素の種類（`class`, `property`, `field` など）やアクセス権（`private`, `public` など）を絞り込むプロパティです。
* **`dotnet_naming_style.[スタイル名].[capitalization / required_prefix / required_suffix]`**: 
  大文字小文字のケース（`pascal_case`, `camel_case`）や、必須とする接頭辞・接尾辞を定義するプロパティです。

---

## 3. サンプル設定コード（VB.NET 向けカスタム設定）と各行解説
[cite_start][cite: 1][cite_start]提示された特定のコードルールに準拠した `.editorconfig` の実設定例です [cite: 1]。各設定ブロックおよび行が持つ意味の詳細をコメントとあわせて記載しています。

```ini
# =====================================================================
# 最上位の設定ファイルであることを宣言
# =====================================================================
root = true 

# =====================================================================
# すべてのファイルに共通するインデント・改行の定義
# =====================================================================
[*]
indent_size = 4 
indent_style = space 
tab_width = 4 
end_of_line = crlf 
insert_final_newline = false 

# =====================================================================
# VB.NET (*.vb) 専用の命名規則（コーディング規約）定義
# =====================================================================
[*.vb] 

#### ① 名前付けルールの割り当て (Rules) ####
# 「どのシンボル（対象）」に「どの表現形式」を適用し、「どの重要度」で警告するかを紐づけます。
# ※ 判定は上から順に行われるため、最も条件が厳しいプライベートフィールドを最上部に配置します。

# ルール1: プライベートフィールドは「_camelCase」にする
dotnet_naming_rule.vb_private_fields_should_be_underscore_camel.symbols = vb_private_field_symbols 
dotnet_naming_rule.vb_private_fields_should_be_underscore_camel.style = underscore_camel_case_style 
dotnet_naming_rule.vb_private_fields_should_be_underscore_camel.severity = warning 

# ルール2: クラス、構造体、列挙型、プロパティ、メソッド、定数は「PascalCase」にする
dotnet_naming_rule.vb_types_and_members_should_be_pascal_case.symbols = vb_pascal_case_symbols 
dotnet_naming_rule.vb_types_and_members_should_be_pascal_case.style = pascal_case_style 
dotnet_naming_rule.vb_types_and_members_should_be_pascal_case.severity = warning 

# ルール3: ローカル変数およびメソッドの引数は「camelCase」にする
dotnet_naming_rule.vb_locals_and_parameters_should_be_camel_case.symbols = vb_camel_case_symbols 
dotnet_naming_rule.vb_locals_and_parameters_should_be_camel_case.style = camel_case_style 
dotnet_naming_rule.vb_locals_and_parameters_should_be_camel_case.severity = warning 


#### ② 対象となる要素の定義 (Symbols) ####
# ルールが適用されるコード上の具体的な「対象（シンボル）」を定義します。

# 対象1: アクセス修飾子が private または friend のフィールド（変数）
dotnet_naming_symbols.vb_private_field_symbols.applicable_kinds = field 
dotnet_naming_symbols.vb_private_field_symbols.applicable_accessibilities = private, friend 

# 対象2: クラス、構造体、Enum、プロパティ、メソッド、または「定数(Const)」であるフィールド
dotnet_naming_symbols.vb_pascal_case_symbols.applicable_kinds = class, struct, enum, property, method, field 
dotnet_naming_symbols.vb_pascal_case_symbols.applicable_accessibilities = * 
dotnet_naming_symbols.vb_pascal_case_symbols.required_modifiers = const 

# 対象3: ローカル変数（メソッド内部の変数）および パラメーター（引数）
dotnet_naming_symbols.vb_camel_case_symbols.applicable_kinds = local, parameter 


#### ③ スタイルの具体的な定義 (Styles) ####
# 変数名などの「表現形式（ケース）」の具体的な形を定義します。

# スタイル1: 先頭にアンダースコア（_）を必須とし、続く文字をキャメルケース（小文字始まり）にする
dotnet_naming_style.underscore_camel_case_style.capitalization = camel_case 
dotnet_naming_style.underscore_camel_case_style.required_prefix = _ 

# スタイル2: パスカルケース（大文字始まり、単語の区切りも大文字）にする
dotnet_naming_style.pascal_case_style.capitalization = pascal_case 

# スタイル3: キャメルケース（小文字始まり、単語の区切りは大文字）にする（接頭辞なし）
dotnet_naming_style.camel_case_style.capitalization = camel_case 
```

### この設定ファイルを導入する効果
* **自動検知と警告:** 開発者が `Private test As String` や `Dim Test As Integer` と記述した際、即座に Visual Studio 画面上に緑色の波線で警告（`severity = warning`）を表示します。
* **一括クイック修正:** 違反箇所にカーソルを合わせ、`Ctrl` + `.` を押下することで、それぞれ `_test` や `test` へと規約に沿った適切な命名にワンクリックで自動変換を行うことが可能です。

---

## 4. 参考リソース・公式ドキュメント
さらに詳細な仕様や追加の設定項目については、以下の公式ドキュメントを参照してください。

* [Microsoft Learn: Visual Studio で移植可能なカスタムのエディター設定を作成する](https://learn.microsoft.com/ja-jp/visualstudio/ide/create-portable-custom-editor-options?view=visualstudio)
  * Visual Studio における EditorConfig の基本的な使い方、ファイル生成、適用優先順位に関する公式チュートリアルです。
* [Microsoft Learn: .NET のコード分析スタイル規則のドキュメント](https://learn.microsoft.com/ja-jp/dotnet/fundamentals/code-analysis/style-rules/?view=visualstudio)
  * 今回定義した命名規則の他にも、`Me.` や `this.` の制限、波かっこ、パターンマッチングといった .NET 固有のすべてのコードスタイル規則プロパティが網羅されています。
* [EditorConfig 公式サイト: サンプルファイル解説（英語）](https://editorconfig.org/#example-file)
  * プログラミング言語を問わない、EditorConfig 標準のファイルフォーマット（インデントや改行コード、文字コード設定など）の基本仕様と具体例が掲載されています。
```
