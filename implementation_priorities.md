# Excel メタデータ抽出ツール リファクタリング進捗

## 0. 現状の処理構成と課題

### 0.1 現状のファイル構成
```
ExcelMetaExtractor/
├── main.py                      # メインのエントリーポイント
├── excel_metadata_extractor.py  # 主要な抽出ロジック
└── openai_helper.py            # OpenAI APIヘルパー
```

### 0.2 主要な機能と実装状況
1. **チャートの抽出**
   - 棒グラフ、折れ線グラフ、円グラフの検出
   - データ系列とカテゴリの抽出
   - チャートタイトルと軸ラベルの取得

2. **図形要素の抽出**
   - 画像の位置情報と内容の取得
   - SmartArtの構造解析
   - 図形の種類と属性の特定

3. **領域検出**
   - テーブル領域の特定
   - テキストブロックの抽出
   - フォームコントロールの検出

### 0.3 現状の課題
1. **コードの複雑性**
   - `excel_metadata_extractor.py`に全ての処理が集中
   - 機能間の依存関係が複雑
   - エラーハンドリングが不統一

2. **保守性の問題**
   - テストコードの不足
   - ドキュメントの不足
   - 型ヒントの部分的な使用

3. **拡張性の制限**
   - 新機能追加時の影響範囲が大きい
   - 機能の再利用が困難
   - 設定の柔軟性が低い

4. **パフォーマンスの課題**
   - メモリ使用量の最適化が必要
   - 大規模ファイルでの処理速度低下
   - リソース解放のタイミング制御

### 0.4 コードの具体的な問題点

1. **excel_metadata_extractor.py の問題**
   ```python
   class ExcelMetadataExtractor:
       def extract_all_metadata(self):
           # 1000行以上の巨大なメソッド
           # 様々な処理が混在
           pass

       def detect_regions(self):
           # 複雑な条件分岐
           # エラーハンドリングの不統一
           pass

       def extract_drawing_info(self):
           # 画像、図形、SmartArtの処理が混在
           # メモリ管理が不適切
           pass
   ```

2. **エラーハンドリングの問題**
   ```python
   try:
       # 大きすぎるtryブロック
       # 例外の種類による処理分けがない
   except Exception as e:
       # 一般的すぎる例外捕捉
       print(f"Error: {str(e)}")
   ```

3. **型ヒントの不統一**
   ```python
   def process_chart(self, chart):  # 型ヒントなし
       pass

   def extract_text(self, cell: Any) -> str:  # Anyの過剰使用
       pass

   def get_metadata(self, sheet: Worksheet) -> Dict[str, Any]:  # 戻り値の型が不明確
       pass
   ```

4. **設定の硬直化**
   ```python
   class ExcelMetadataExtractor:
       def __init__(self):
           self.MAX_CELLS = 1000  # ハードコードされた定数
           self.SUPPORTED_TYPES = ['xlsx', 'xlsm']  # 拡張が困難
   ```

5. **テストの問題**
   - 単体テストが不足
   - テストケースのカバレッジが低い
   - テストデータの準備が不十分

### 0.5 リファクタリングによる改善点

1. **コードの分割**
   - 機能ごとに適切なクラスに分割
   - 責務の明確な分離
   - インターフェースの整理

2. **エラーハンドリング**
   ```python
   try:
       # 具体的な処理
   except FileNotFoundError as e:
       # ファイル関連のエラー処理
   except ValueError as e:
       # 値のエラー処理
   except ExtractionError as e:
       # 抽出処理特有のエラー処理
   ```

3. **型ヒントの改善**
   ```python
   def process_chart(self, chart: Chart) -> ChartData:
       pass

   def extract_text(self, cell: Cell) -> str:
       pass

   def get_metadata(self, sheet: Worksheet) -> MetadataDict:
       pass
   ```

4. **設定の柔軟化**
   ```python
   class ExcelMetadataExtractor:
       def __init__(self, config: Config):
           self.max_cells = config.max_cells
           self.supported_types = config.supported_types
   ```

## 1. 実装済みの内容

### 1.1 チャートデータ抽出機能の改善
- ✅ 基底クラス（`BaseExtractor`）の実装
  - エラーハンドリングの共通化
  - ロギング機能の統一
  - 型ヒントの導入

- ✅ チャートデータモデル（`ChartData`）の実装
  - データクラスを使用した明確な構造化
  - 必要な属性の整理
  - 辞書形式への変換機能

- ✅ チャート抽出機能（`ChartExtractor`）の改善
  - メソッドの責務を明確に分離
  - エラーハンドリングの強化
  - デバッグ情報の充実化
  - カテゴリ抽出ロジックの改善

- ✅ テストの整備
  - 基本的なチャートデータ抽出のテスト実装
  - テストデータの作成と検証
  - エッジケースの考慮

## 2. 今後の実装計画

### フェーズ2: 図形要素の分離
- 🔲 モデルの作成
  ```
  models/
  ├── drawing.py
  │   ├── DrawingData（基底クラス）
  │   ├── ImageData
  │   ├── ShapeData
  │   └── SmartArtData
  ```
- 🔲 エクストラクターの実装
  ```
  extractors/
  └── drawing_extractor.py
      ├── DrawingExtractor（基底クラス）
      ├── ImageExtractor
      ├── ShapeExtractor
      └── SmartArtExtractor
  ```

### フェーズ3: 領域検出機能の分離
- 🔲 モデルの作成
  ```
  models/
  └── region.py
      ├── RegionData（基底クラス）
      ├── TableRegion
      ├── TextRegion
      └── FormControlRegion
  ```
- 🔲 エクストラクターの実装
  ```
  extractors/
  └── region_extractor.py
      ├── RegionExtractor（基底クラス）
      ├── TableExtractor
      ├── TextExtractor
      └── FormControlExtractor
  ```

### フェーズ4: 共通機能の整理
- 🔲 ユーティリティの作成
  ```
  utils/
  ├── excel.py
  ├── validation.py
  └── conversion.py
  ```
- 🔲 設定管理の導入
  ```
  config/
  ├── settings.py
  └── logging.py
  ```

### フェーズ5: エラーハンドリングの改善
- 🔲 カスタム例外の導入
  ```
  exceptions/
  ├── base.py
  ├── extraction.py
  └── validation.py
  ```
- 🔲 エラーメッセージの整理
  ```
  constants/
  ├── messages.py
  └── codes.py
  ```

### フェーズ6: テストの拡充
- 🔲 単体テストの追加
  - 各エクストラクターのテスト
  - ユーティリティのテスト
  - エラーケースのテスト

- 🔲 統合テストの追加
  - 複合的なデータ抽出のテスト
  - エッジケースのテスト

### フェーズ7: ドキュメントの整備
- 🔲 APIドキュメント
  - 各クラス・メソッドの使用方法
  - パラメータの説明
  - 戻り値の形式

- 🔲 開発者ドキュメント
  - アーキテクチャの説明
  - 貢献ガイドライン
  - テスト方法の説明

## 3. 技術的な考慮事項

### 3.1 依存関係
- openpyxl: Excelファイルの操作
- dataclasses: データモデルの定義
- typing: 型ヒント
- logging: ログ出力

### 3.2 品質基準
- コードカバレッジ: 80%以上
- 型ヒント: 全てのパブリックAPIに必須
- ドキュメント: 全ての公開クラス・メソッドに必須
- エラーハンドリング: 全ての外部操作に必須

### 3.3 パフォーマンス目標
- 大規模ファイル（100MB以上）でも処理可能
- メモリ使用量の最適化
- 処理速度の向上

## 4. 次のアクション

1. フェーズ2の詳細設計書の作成
2. 図形要素の抽出に関する要件の整理
3. テストケースの設計
4. レビュー基準の策定 