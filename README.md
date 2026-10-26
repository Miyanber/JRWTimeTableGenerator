# JRWTimeTableGenerator

JR西日本の公式サイト（[JRおでかけネット](https://www.jr-odekake.net/railroad/#eki)）で公開されている駅のWeb時刻表から、かつて駅で配布されていた**ポケット時刻表（小型時刻表）風のシート**を Excel 上に自動生成する Office アドインです。

> [!IMPORTANT]
> 本ツールは個人が作成した**非公式ツール**であり、西日本旅客鉄道株式会社（JR西日本）とは一切関係ありません。
> 時刻表データの著作権は各権利者に帰属します。生成物は個人的な利用の範囲でご利用ください。

## 主な機能

- 平日用・土休日用の Web 時刻表（HTML）を読み込み、1枚のシートに左右並べて出力
- 列車種別（新快速・快速など）と行先を、記号・略称・色で表示
- 1時間あたりの最大本数に応じて、列数と列幅を自動調整
- 凡例（種別・行先の説明）を自動生成
- 特急列車の除外オプション
- 選択した HTML ファイルの「駅名 / 方面 / 平日・土休日」の自動判定と、取り違えの警告

## 動作環境

| 項目 | 要件 |
|---|---|
| Excel | Microsoft 365 / Excel 2016 以降（Windows デスクトップ版、Excel on the web） |
| 開発環境 | Visual Studio 2022（**「Office/SharePoint 開発」ワークロード**が必要） |
| フレームワーク | .NET Framework 4.8（Web プロジェクト） |
| ネットワーク | 実行時に Office.js（CDN）、Google Fonts、データ取得用の Google Apps Script へアクセスします |

## リポジトリ構成

```text
JRWTimeTableGenerator.sln
├─ JRWTimeTableGenerator/            # アドイン プロジェクト（マニフェスト）
│  └─ JRWTimeTableGeneratorManifest/
│     └─ JRWTimeTableGenerator.xml   # Office アドイン マニフェスト
└─ JRWTimeTableGeneratorWeb/         # アドイン本体（作業ウィンドウの Web アプリ）
   ├─ Home.html / Home.css / Home.js # 作業ウィンドウ UI とメインロジック
   ├─ Functions/                     # リボン コマンド用の関数ファイル
   ├─ Content/  Scripts/  Images/    # CSS・jQuery・アイコン
   └─ Web.config
```

## セットアップ（開発・デバッグ）

1. リポジトリをクローンします。
   ```powershell
   git clone https://github.com/Miyanber/JRWTimeTableGenerator.git
   ```
2. Visual Studio 2022 で `JRWTimeTableGenerator.sln` を開きます。
3. 初回は NuGet パッケージを復元します（ソリューションを右クリック →「NuGet パッケージの復元」）。
4. スタートアップ プロジェクトを **`JRWTimeTableGenerator`**（マニフェスト側）にして `F5` を押します。
   - IIS Express（`https://localhost:44346/`）で Web プロジェクトが起動し、アドインがサイドロードされた状態で Excel が開きます。
   - 初回は IIS Express の開発用証明書を信頼するよう求められることがあります。
5. Excel の「ホーム」タブにある **「Show Taskpane」** ボタンで作業ウィンドウを開きます。

### 自分でホスティングする場合

マニフェスト内の `~remoteAppUrl` は、Visual Studio が実行時に自動で置き換えるプレースホルダーです。手動でサイドロードしたり、他のサーバーで公開したりする場合は、`~remoteAppUrl` を **HTTPS で公開した `JRWTimeTableGeneratorWeb/` の URL** に書き換えてください。手動サイドロードの手順は [Microsoft のドキュメント](https://learn.microsoft.com/ja-jp/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing) を参照してください。

## 使い方

1. テンプレートのワークシートをブックに用意し、シート名を **`base`** にします（アドインはこのシートを複製して出力します）。
2. JRおでかけネットで目的の駅・方面の Web 時刻表を開き、**平日用**と**土休日用**をそれぞれ「名前を付けて保存」で **HTML 形式**で保存します。
3. 作業ウィンドウの「平日用」「土・休日用」をクリックし、保存した HTML を選択します。
   - 駅名・方面・平日/土休日が表示されます。曜日区分が合っていない場合は赤字で警告されます。
4. 必要に応じて「特急列車を除外」を切り替え、**「作成」** をクリックします。
5. 先頭に `駅名-方面` という名前のシートが生成されます。印刷するときは拡大縮小を **25%** にしてください。

## データソースについて

種別・行先の表記ルール（略称・記号・色など）は、作者が管理する Google スプレッドシートを Google Apps Script 経由で JSON として取得しています（[Home.js](JRWTimeTableGeneratorWeb/Home.js) の `DESTINATION_URL` / `TRAIN_TYPE_URL`）。

- 公開されているのは読み取り専用のエンドポイントです。
- エンドポイントが停止または変更されると、アドインは「列車情報を読み込み中」から先へ進まなくなります。フォークして使う場合は、同じ形式（`{ "allData": [[...], ...] }`）の JSON を返すエンドポイントを用意し、URL を差し替えてください。

## 既知の制限

- 祝日データは [Home.js](JRWTimeTableGeneratorWeb/Home.js) にハードコードされており、**2023年まで**しかありません。それ以降の日付では、祝日を平日と判定します（警告表示にのみ影響します）。
- テンプレートの `base` ワークシートは現在リポジトリに含まれていません。作業ウィンドウのダウンロードリンクも未設定です。
- 「左上端のセル」オプションは未実装です。
- JRおでかけネットのページ構造が変わると、時刻表を解析できなくなる可能性があります。

## セキュリティ

脆弱性を見つけた場合は [SECURITY.md](SECURITY.md) を参照してください。

## ライセンス

[MIT License](LICENSE) © 2022 Miyanber

同梱しているサードパーティ ライブラリは、いずれも MIT License です。

| ライブラリ | バージョン | ライセンス |
|---|---|---|
| [jQuery](https://jquery.com/) | 3.6.0 | MIT |
| [Office UI Fabric Core](https://github.com/OfficeDev/office-ui-fabric-core) | 9.6.0 | MIT |
| [Microsoft.CodeDom.Providers.DotNetCompilerPlatform](https://github.com/aspnet/RoslynCodeDomProvider) | 2.0.1 | MIT（NuGet で取得、同梱なし） |

Office.js は Microsoft の CDN（`appsforoffice.microsoft.com`）から読み込んでおり、リポジトリには含まれていません。

## 作者

Miyanber — [Twitter (X)](https://twitter.com/seiyanber3818)