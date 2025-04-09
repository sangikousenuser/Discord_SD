# Word Save and PDF Add-in

このWord向け拡張機能（アドイン）は、ドキュメントの保存とPDF変換を1クリックで行うことができます。

## 機能

- Wordのタブバーにボタンを表示
- ボタンを押すとWordファイルを保存
- PDFを自動的に作成
- ファイルパスをクリップボードにコピー（環境によりサポート）

## 技術的な詳細

この拡張機能は以下の技術を使用しています：

- Office JavaScript API (Office JS)
- Adobe PDF Services SDK（PDF変換用）
- HTML/CSS/JavaScript

## 制限事項

Office JSの技術的な制約により、以下の制限があります：

1. ドキュメント保存時にはユーザーに保存ダイアログが表示されます（自動保存はAPIの制限により不可）
2. PDF変換には Adobe PDF Services の認証情報が必要です
3. クリップボード機能はブラウザやOfficeのバージョンによって動作が異なる場合があります

## インストール方法

### 開発環境での実行

1. 必要なパッケージをインストール：
   ```
   npm install
   ```

2. 開発サーバーを起動：
   ```
   npm run dev
   ```

3. Office Add-inをサイドロードする：
   - Word for Webを開く
   - 「挿入」タブ > 「アドイン」 > 「マイアドイン」 > 「マニフェストを管理」
   - 「参照」をクリックし、manifest.xmlファイルを選択

### 本番環境へのデプロイ

1. ビルドを実行：
   ```
   npm run build
   ```

2. 生成されたdistフォルダの内容をウェブサーバーにアップロード

3. manifest.xmlファイル内のURLを本番環境のURLに更新

4. 更新したmanifest.xmlファイルを使用してアドインをインストール

## カスタマイズ

PDF変換機能を有効にするには：

1. Adobe PDF Services APIの認証情報を取得（https://developer.adobe.com/document-services/）
2. 認証情報ファイルをプロジェクトに追加
3. taskpane.jsファイル内のcreatePDF関数のコメントを解除し、適切なパスを設定

## トラブルシューティング

- **アドインが表示されない**: マニフェストファイルが正しくサイドロードされているか確認
- **保存ボタンが動作しない**: Office JSのバージョンとの互換性を確認
- **PDF変換が失敗する**: Adobe PDF Services APIの認証情報が正しいか確認

## 今後の改善点

- より直感的なユーザーインターフェース
- 保存場所の設定オプション
- PDF変換オプションのカスタマイズ
- 複数のドキュメント形式のサポート
