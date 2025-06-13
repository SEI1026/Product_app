# GitHub CLI セットアップガイド

## 完全自動ビルド＆リリースのための準備

### 1. GitHub CLI のインストール

#### 方法1: 公式サイトからダウンロード
1. https://cli.github.com/ にアクセス
2. "Download for Windows" をクリック
3. ダウンロードしたインストーラを実行
4. インストール完了後、コマンドプロンプトを**再起動**

#### 方法2: winget (Windows 11/10)
```cmd
winget install --id GitHub.cli
```

#### 方法3: Chocolatey
```cmd
choco install gh
```

### 2. GitHub認証の設定

#### 2.1 認証開始
```cmd
gh auth login
```

#### 2.2 認証手順
```
? What account do you want to log into?
> GitHub.com

? What is your preferred protocol for Git operations?
> HTTPS

? Authenticate Git with your GitHub credentials?
> Yes

? How would you like to authenticate GitHub CLI?
> Login with a web browser
```

#### 2.3 ブラウザ認証
1. ワンタイムコードが表示される（例：ABCD-1234）
2. Enter キーを押すとブラウザが開く
3. GitHubにログインしてコードを入力
4. 認証完了

### 3. リポジトリアクセス権限の確認

#### 3.1 認証状態確認
```cmd
gh auth status
```

期待する出力：
```
✓ Logged in to github.com as [ユーザー名] (keyring)
✓ Git operations for github.com configured to use https protocol.
✓ Token: *******************
```

#### 3.2 リポジトリアクセステスト
```cmd
gh repo view SEI1026/Product_app
```

### 4. Git の設定（初回のみ）

#### 4.1 ユーザー情報設定
```cmd
git config --global user.name "あなたの名前"
git config --global user.email "your-email@example.com"
```

#### 4.2 リポジトリの初期化（新規の場合）
プロジェクトフォルダで以下を実行：
```cmd
git init
git remote add origin https://github.com/SEI1026/Product_app.git
git branch -M main
git add .
git commit -m "Initial commit"
git push -u origin main
```

### 5. 自動ビルド＆リリース実行

準備完了後、以下のコマンド一つで全て自動実行：

```cmd
build_and_release.bat
```

### 6. 実行内容

`build_and_release.bat` は以下を自動実行：

1. ✅ **環境チェック** - Python, PyInstaller, GitHub CLI
2. ✅ **依存関係インストール** - 必要なパッケージ
3. ✅ **ビルド実行** - PyInstallerでEXE作成
4. ✅ **ZIP作成** - 配布用パッケージ作成
5. ✅ **version.json更新** - 自動アップデート用URL更新
6. ✅ **GitHub Release作成** - タグ付きリリース作成
7. ✅ **ファイルアップロード** - ZIPファイル自動アップロード
8. ✅ **Git コミット** - version.json変更をコミット&プッシュ

### 7. トラブルシューティング

#### 7.1 GitHub CLI が見つからない
```
ERROR: GitHub CLI (gh) is not installed.
```
→ GitHub CLIをインストール後、コマンドプロンプトを再起動

#### 7.2 認証エラー
```
ERROR: GitHub authentication required.
```
→ `gh auth login` を実行して認証

#### 7.3 リポジトリアクセスエラー
```
ERROR: Failed to create GitHub release.
```
→ リポジトリの権限を確認（Writeアクセスが必要）

#### 7.4 Git コミットエラー
```
WARNING: Failed to push version.json changes.
```
→ Git設定を確認（user.name, user.email）

### 8. バージョンアップ手順

#### 8.1 新バージョンリリース
1. `src/utils/version_checker.py` の `CURRENT_VERSION` を更新
2. `build_and_release.bat` の `APP_VERSION` を更新
3. `version.json` の changelog を更新
4. `build_and_release.bat` を実行

#### 8.2 完全自動化
```cmd
# バージョン更新後、このコマンド1つで完了
build_and_release.bat
```

### 9. セキュリティ注意事項

#### 9.1 アクセストークン
- GitHub CLIは安全にトークンを管理
- トークンの手動管理は不要

#### 9.2 リポジトリ権限
- Releases作成にはWrite権限が必要
- 組織のリポジトリの場合は管理者に確認

#### 9.3 二段階認証
- GitHubで二段階認証を有効にしていても正常動作
- ブラウザ認証で対応

---

## 完了後の確認

✅ GitHub CLI インストール済み  
✅ GitHub認証完了  
✅ リポジトリアクセス確認済み  
✅ Git設定完了  

これで `build_and_release.bat` を実行するだけで、ビルドからGitHubリリースまで全自動で完了します！

---

**株式会社大宝家具**  
商品登録入力ツール 完全自動化システム