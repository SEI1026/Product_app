# GitHub 自動アップデート機能セットアップガイド

## 1. GitHubリポジトリの作成

### 1.1 リポジトリ作成
1. GitHub（https://github.com）にログイン
2. 「New repository」をクリック
3. リポジトリ名: `product-register-tool`（または任意の名前）
4. 説明: `商品登録入力ツール - ECサイト向け商品データ管理アプリケーション`
5. Public または Private を選択
6. 「Add a README file」をチェック
7. 「Create repository」をクリック

### 1.2 リポジトリ構造
```
product-register-tool/
├── README.md                          # プロジェクト説明
├── version.json                       # バージョン情報（自動更新用）
├── releases/                          # リリースファイル格納
│   └── v2.1.0/
│       └── ProductRegisterTool-v2.1.0.zip
└── docs/                              # ドキュメント
    ├── INSTALLATION.md
    └── CHANGELOG.md
```

## 2. version.json の配置

### 2.1 ファイルの配置
- `version.json` をリポジトリのルートディレクトリに配置
- このファイルはアプリケーションが定期的にチェックします

### 2.2 アクセスURL
リポジトリが `https://github.com/username/product-register-tool` の場合：
```
https://raw.githubusercontent.com/username/product-register-tool/main/version.json
```

## 3. GitHub Releasesの使用

### 3.1 リリースの作成手順
1. GitHubのリポジトリページで「Releases」タブをクリック
2. 「Create a new release」をクリック
3. タグバージョン: `v2.1.0`（version.jsonと一致させる）
4. リリースタイトル: `商品登録入力ツール v2.1.0`
5. 説明にchangelogを記載
6. `ProductRegisterTool-v2.1.0.zip` をアップロード
7. 「Publish release」をクリック

### 3.2 ダウンロードURL
GitHubが自動生成するURL形式：
```
https://github.com/username/product-register-tool/releases/download/v2.1.0/ProductRegisterTool-v2.1.0.zip
```

## 4. 設定ファイルの更新

### 4.1 version_checker.py の URL 更新
```python
# あなたのGitHubリポジトリに合わせて変更
VERSION_CHECK_URL = "https://raw.githubusercontent.com/YOUR_USERNAME/product-register-tool/main/version.json"
```

### 4.2 version.json の download_url 更新
```json
{
  "version": "2.1.0",
  "download_url": "https://github.com/YOUR_USERNAME/product-register-tool/releases/download/v2.1.0/ProductRegisterTool-v2.1.0.zip"
}
```

## 5. 自動アップデート機能

### 5.1 アプリケーション起動時
- 自動的にGitHubの `version.json` をチェック
- 新しいバージョンがある場合、ユーザーに更新を提案

### 5.2 手動チェック
- ヘルプメニュー → 「更新の確認」
- いつでも手動で最新バージョンをチェック可能

### 5.3 更新プロセス
1. GitHubから更新ファイルをダウンロード
2. 現在のEXEファイルを `.new` として保存
3. 再起動時に古いファイルを新しいファイルで置換
4. 更新完了

## 6. リリースワークフロー（推奨）

### 6.1 新バージョンリリース手順
1. **version.json の更新**
   - バージョン番号を更新
   - changelogを追加
   - download_urlを新しいバージョンに更新

2. **EXEビルド**
   ```bat
   build_complete.bat
   ```

3. **GitHub Releaseの作成**
   - 新しいタグを作成
   - ZIPファイルをアップロード
   - リリースノートを記載

4. **version.json をリポジトリにコミット**
   ```bash
   git add version.json
   git commit -m "Update version to v2.1.0"
   git push
   ```

### 6.2 自動化（GitHub Actions - オプション）
```yaml
# .github/workflows/release.yml
name: Auto Release
on:
  push:
    tags:
      - 'v*'
jobs:
  release:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v2
      - name: Build EXE
        run: build_complete.bat
      - name: Create Release
        uses: actions/create-release@v1
        with:
          tag_name: ${{ github.ref }}
          release_name: Release ${{ github.ref }}
          draft: false
          prerelease: false
      - name: Upload Release Asset
        uses: actions/upload-release-asset@v1
        with:
          upload_url: ${{ steps.create_release.outputs.upload_url }}
          asset_path: ./ProductRegisterTool-v2.1.0.zip
          asset_name: ProductRegisterTool-v2.1.0.zip
          asset_content_type: application/zip
```

## 7. セキュリティ考慮事項

### 7.1 署名付きEXEファイル（推奨）
- コードサイニング証明書を使用してEXEファイルに署名
- Windows Defenderの警告を減らすことができます

### 7.2 HTTPS通信
- すべての通信はHTTPS経由で行われます
- GitHubのSSL証明書により通信が保護されます

### 7.3 検証機能
- ダウンロードファイルのハッシュ値チェック（オプション）
- バージョン番号の整合性確認

## 8. トラブルシューティング

### 8.1 よくある問題
- **404エラー**: リポジトリのURL確認
- **ダウンロード失敗**: ファイアウォール設定確認
- **更新失敗**: 管理者権限で実行

### 8.2 ログの確認
アプリケーションログで詳細なエラー情報を確認可能：
```
Documents/ProductAppUserData/logs/
```

---

## 株式会社大宝家具 向けカスタマイズ

この自動アップデート機能により以下のメリットがあります：

1. **自動配布**: 新機能・バグ修正を自動でユーザーに配布
2. **バージョン管理**: 複数バージョンの並行管理が容易
3. **使用統計**: GitHubでダウンロード数を確認可能
4. **セキュリティ**: HTTPS通信とGitHubの信頼性
5. **コスト削減**: GitHub Releasesは無料で使用可能