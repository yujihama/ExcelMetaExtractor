# Excel Metadata Extractor

高度なExcelメタデータ抽出と解析のためのStreamlitベースのWebアプリケーション。複雑なスプレッドシート構造とXMLベースの図形情報を正確に解析します。

## 主な機能

### メタデータ解析
- Excelファイルからの包括的なメタデータ抽出
- シート構造、セル情報、フォーマット設定の詳細分析
- 数式、参照関係、依存関係の解析
- コメント、ノート、変更履歴の抽出

### 図形・グラフ解析
- XMLベースの図形情報の詳細抽出（drawing_extractor.py）
- VML要素の解析と構造化（vml_processor.py）
- チャート構造とデータ系列の分析（chart_processor.py）
- 図形の位置、サイズ、スタイル情報の抽出

### 領域分析機能
- スプレッドシートの論理的な領域検出（region_detector.py）
- セル領域の意味解析と分類（region_analyzer.py）
- データ構造の自動認識
- テーブル領域の特定と解析

### AI統合機能
- OpenAI gptによるメタデータの意味解析
- 図形とチャートの自動分類と説明生成
- データ構造の自動検出と推奨事項の提示
- 日本語コンテンツの高精度な解析と理解

### インターフェース
- Streamlitによる直感的なWebインターフェース
- リアルタイムのデータ可視化と探索機能
- インタラクティブな解析結果の表示
- カスタマイズ可能なレポート生成

## 他のExcelライブラリとの機能比較

本ツールは、既存のPythonライブラリと比較して、より包括的で高度な機能を提供します。

### ライブラリ/ツールの特徴

#### openpyxl
- Excel 2010以降のxlsx形式に特化
- セル単位の詳細な操作が可能
- メモリ効率が良好
- 環境制約：
  - 特定のOS制約なし
  - Excel本体のインストール不要
  - .xlsm形式のマクロ付きファイルは読み取りのみ対応

#### pandas
- データ分析に特化
- 大規模データの処理が得意
- データフレーム形式での操作
- 環境制約：
  - 特定のOS制約なし
  - Excel本体のインストール不要
  - openpyxlまたはxlrdなどのバックエンドライブラリが必要
  - メモリ制約：大規模データの場合はRAMの制限に注意

#### xlwings
- Excelアプリケーションの制御が可能
- マクロやVBAとの連携可能
- Windowsでの使用に最適
- 環境制約：
  - Windows環境推奨（Macでは機能制限あり）
  - Excel本体のインストールが必要
  - COM interfaceへのアクセス権限が必要
  - 32bit/64bit版Excelとの互換性に注意

#### Excel Metadata Extractor（当ツール）
- AIを活用した高度な解析機能
- 日本語コンテンツの精密な解析
- 包括的なメタデータ抽出
- 直感的なWebインターフェース
- リアルタイムの可視化機能
- 環境制約：
  - クロスプラットフォーム対応（Windows/Mac/Linux）
  - Excel本体のインストール不要
  - インターネット接続必須（AI機能使用時）
  - Python 3.11以上が必要
  - メモリ要件：最低4GB推奨

### 機能別の対応関係

#### 1. 基本メタデータ

##### ファイルプロパティ情報
| 機能 | openpyxl | pandas | xlwings | Excel Metadata Extractor |
|------|----------|---------|----------|------------------------|
| 作成者情報 | ✅ | ❌ | ✅ | ✅ |
| 作成日時 | ✅ | ❌ | ✅ | ✅ |
| ファイルサイズ | ✅ | ❌ | ✅ | ✅ |
| バージョン情報 | ✅ | ❌ | ✅ | ✅ |
| カスタムプロパティ | ✅ | ❌ | ✅ | ✅⭐ |

##### シート構造
| 機能 | openpyxl | pandas | xlwings | Excel Metadata Extractor |
|------|----------|---------|----------|------------------------|
| シート名一覧 | ✅ | ✅ | ✅ | ✅ |
| 表示/非表示状態 | ✅ | ❌ | ✅ | ✅ |
| シート間リンク | ✅ | ❌ | ✅ | ✅⭐ |
| シート保護状態 | ✅ | ❌ | ✅ | ✅ |
| 印刷設定 | ✅ | ❌ | ✅ | ✅ |
| シート間の依存関係分析 | ❌ | ❌ | ❌ | ✅⭐ |

##### セル情報
| 機能 | openpyxl | pandas | xlwings | Excel Metadata Extractor |
|------|----------|---------|----------|------------------------|
| セル値とデータ型 | ✅ | ✅ | ✅ | ✅ |
| 書式設定 | ✅ | ❌ | ✅ | ✅⭐ |
| 数式 | ✅ | ❌ | ✅ | ✅⭐ |
| 条件付き書式 | ✅ | ❌ | ✅ | ✅⭐ |
| セル結合 | ✅ | ❌ | ✅ | ✅ |
| 入力規則 | ✅ | ❌ | ✅ | ✅⭐ |
| 数式の依存関係分析 | ❌ | ❌ | ❌ | ✅⭐ |

#### 2. 図形・グラフ関連情報

##### 図形要素
| 機能 | openpyxl | pandas | xlwings | Excel Metadata Extractor |
|------|----------|---------|----------|------------------------|
| 図形の種類 | ✅ | ❌ | ✅ | ✅⭐ |
| 座標情報 | ✅ | ❌ | ✅ | ✅⭐ |
| サイズ・回転 | ✅ | ❌ | ✅ | ✅ |
| 書式設定 | ✅ | ❌ | ✅ | ✅⭐ |
| VML要素 | ✅ | ❌ | ❌ | ✅⭐ |
| 図形の意味解析 | ❌ | ❌ | ❌ | ✅⭐ |

##### チャート情報
| 機能 | openpyxl | pandas | xlwings | Excel Metadata Extractor |
|------|----------|---------|----------|------------------------|
| チャート種類 | ✅ | ❌ | ✅ | ✅⭐ |
| データ系列 | ✅ | ❌ | ✅ | ✅⭐ |
| 軸設定 | ✅ | ❌ | ✅ | ✅ |
| グラフ書式 | ✅ | ❌ | ✅ | ✅ |
| チャートの意味解析 | ❌ | ❌ | ❌ | ✅⭐ |

#### 3. 領域分析情報

##### テーブル領域
| 機能 | openpyxl | pandas | xlwings | Excel Metadata Extractor |
|------|----------|---------|----------|------------------------|
| テーブル範囲 | ✅ | 🔸 | ✅ | ✅⭐ |
| ヘッダー識別 | ✅ | 🔸 | ✅ | ✅⭐ |
| データ型分布 | ❌ | ✅ | ❌ | ✅⭐ |
| 集計行/列 | ❌ | ✅ | ✅ | ✅⭐ |
| 論理的領域の自動検出 | ❌ | ❌ | ❌ | ✅⭐ |

##### データ構造
| 機能 | openpyxl | pandas | xlwings | Excel Metadata Extractor |
|------|----------|---------|----------|------------------------|
| ピボットテーブル | ❌ | 🔸 | ✅ | ✅⭐ |
| クロス集計 | ❌ | ✅ | ✅ | ✅⭐ |
| 階層関係 | ❌ | 🔸 | ❌ | ✅⭐ |
| データ構造の自動認識 | ❌ | ❌ | ❌ | ✅⭐ |

#### 4. AI解析による高度な情報
| 機能 | openpyxl | pandas | xlwings | Excel Metadata Extractor |
|------|----------|---------|----------|------------------------|
| 意味解析 | ❌ | ❌ | ❌ | ✅⭐ |
| 構造解析 | ❌ | ❌ | ❌ | ✅⭐ |
| 日本語解析 | ❌ | ❌ | ❌ | ✅⭐ |
| データ品質評価 | ❌ | ❌ | ❌ | ✅⭐ |
| 改善提案生成 | ❌ | ❌ | ❌ | ✅⭐ |

#### 5. レポート生成情報
| 機能 | openpyxl | pandas | xlwings | Excel Metadata Extractor |
|------|----------|---------|----------|------------------------|
| 分析レポート | ❌ | 🔸 | ❌ | ✅⭐ |
| データ可視化 | ❌ | ✅ | ❌ | ✅⭐ |
| インタラクティブ表示 | ❌ | 🔸 | ❌ | ✅⭐ |
| カスタマイズ可能なレポート | ❌ | ❌ | ❌ | ✅⭐ |

凡例：
- ✅：基本的な対応
- ✅⭐：高度な対応（AI活用または詳細な解析機能付き）
- 🔸：部分的に対応
- ❌：非対応

当ツールの主な強み：
1. **AI統合による高度な解析**
   - メタデータの意味理解
   - 構造の自動認識
   - 改善提案の生成

2. **包括的な解析機能**
   - 基本的なExcel機能の完全サポート
   - 高度な図形・チャート解析
   - 論理的な領域検出

3. **日本語処理の優位性**
   - 日本語コンテンツの精密な解析
   - ビジネス文書の文脈理解
   - 日本語特有の表記ゆれ対応

4. **使いやすさ**
   - Webベースの直感的なインターフェース
   - リアルタイムの可視化機能
   - カスタマイズ可能なレポート生成

5. **拡張性**
   - モジュール式のアーキテクチャ
   - APIによる機能拡張
   - カスタム分析の追加が容易

## システム構成

### コアモジュール
- `excel_metadata_extractor.py`: メインの抽出エンジン
- `cell_processor.py`: セルデータの詳細処理
- `chart_processor.py`: チャートデータの解析
- `drawing_extractor.py`: 図形要素の抽出
- `vml_processor.py`: VML形式の解析
- `region_analyzer.py` & `region_detector.py`: 領域検出・分析
- `openai_helper.py`: gpt連携機能
- `logger.py`: 詳細なログ管理

### 技術スタック
- Python 3.11
- Streamlit: Webインターフェース
- OpenPyXL: Excel処理
- Pandas: データ操作
- OpenAI API: メタデータ分析
- ElementTree: XML解析
- Matplotlib: データ可視化

## セットアップ

1. 環境変数の設定
`.env.example`ファイルを`.env`にコピーし、必要な環境変数を設定します：

### OpenAI APIを使用する場合
```env
OPENAI_API_TYPE=openai
OPENAI_API_KEY=your-openai-api-key
OPENAI_MODEL_NAME=gpt-4
```

### Azure OpenAI APIを使用する場合
```env
OPENAI_API_TYPE=azure
AZURE_OPENAI_API_KEY=your-azure-openai-api-key
AZURE_OPENAI_API_VERSION=2024-02-15-preview
AZURE_OPENAI_ENDPOINT=https://your-resource-name.openai.azure.com
AZURE_OPENAI_DEPLOYMENT_NAME=your-model-deployment-name
```

2. 依存パッケージのインストール
```bash
pip install -r requirements.txt
```

3. アプリケーションの起動
```bash
streamlit run main.py
```

## 使用方法

1. Webブラウザで`http://localhost:5000`にアクセス
2. Excelファイルをアップロード
3. 解析オプションを選択：
   - 基本メタデータ解析
   - 図形・チャート解析
   - 領域検出と分析
   - AI支援解析
4. 解析結果の表示：
   - ファイルプロパティ
   - シート構造
   - セル情報
   - 図形・チャートデータ
   - 検出された領域情報
   - AIによる分析レポート
5. 結果のエクスポートとレポート生成

## 環境設定

### 必要な環境変数
- `OPENAI_API_KEY`: OpenAI APIキー（AI解析機能に必須）

### システム要件
- Python 3.11以上
- インターネット接続（AI機能使用時）
- ディスク容量: 最低500MB（アプリケーションと依存関係用）

## パフォーマンス注意事項

- 大規模なExcelファイル（100MB以上）の処理には追加の処理時間が必要
- 複雑なチャートや多数の図形要素を含むファイルは解析に時間がかかる場合あり
- AI解析機能の使用には有効なOpenAI APIキーと十分なAPIクレジットが必要
- メモリ使用量は処理するファイルのサイズに応じて増加

## 注意事項

- 機密情報を含むExcelファイルの処理時は適切なセキュリティ対策を実施してください
- AI解析機能の使用にはAPI利用料金が発生します
- 大量のファイル処理時はシステムリソースの使用状況に注意してください
- 処理可能なファイルサイズには環境に応じた制限があります