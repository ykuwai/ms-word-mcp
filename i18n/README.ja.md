<h1 align="center">
  <img src="https://raw.githubusercontent.com/ykuwai/ms-word-mcp/main/assets/word-mcp-logo-letter.svg" alt="Word MCP (ms-word-mcp)" width="400">
</h1>

<p align="center">
  <a href="https://github.com/ykuwai/ms-word-mcp/blob/main/README.md">English (en)</a> · <strong>日本語 (ja)</strong>
</p>

<p align="center">
  <a href="https://pypi.org/project/ms-word-mcp/"><img src="https://img.shields.io/pypi/v/ms-word-mcp.svg" alt="PyPI"></a>
  <a href="https://www.python.org/"><img src="https://img.shields.io/badge/Python-3.10%2B-blue.svg" alt="Python"></a>
  <a href="https://github.com/ykuwai/ms-word-mcp/blob/main/LICENSE"><img src="https://img.shields.io/badge/License-MIT-green.svg" alt="License"></a>
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20macOS-0078d4.svg" alt="Platform">
</p>

<p align="center">
  <strong>Windows と macOS の Microsoft Word をリアルタイムに操作——<br>あなたの隣で同じ文書を編集する、<!--tools-->101<!--/tools--> ツールの MCP サーバー</strong>
</p>

---

AI エージェントが、起動中の Microsoft Word を操作できるようにする MCP（Model Context Protocol）サーバーです。Windows では COM、macOS では Apple Events を使います。python-docx のようにファイルを書き換えるライブラリとは違い、開いている Word の中で動くので、AI が別の場所を書いている間も同じ文書で入力を続けられます。

> **状況:** 初期リリース（0.x）です。ツール名や引数は今後変わることがあります。

## ✨ 主な特長

- **隣で一緒に作業** — 書き込みはすべて Range を通し、カーソルや選択範囲には触れません。ウィンドウは変更箇所までスクロールします
- **書き込むたびに読み戻す** — 結果には Word が今持っている内容を返し、反映されなかったものは `warnings` に挙げます
- **あらゆる文字体系** — 欧文、東アジア文字（日本語・中国語・韓国語）、複雑な文字（アラビア語・ヘブライ語など）にそれぞれフォントを指定できます。文書グリッド、ルビ、傍点、縦書き、全角の段落番号にも対応しています
- **スタイルが基本** — 文字と段落のプロパティはすべてスタイルにも設定できるので、文書の見た目を 1 か所で定義できます
- **数式** — LaTeX または UnicodeMath から Word 自身の数式エディターで組み立て、右端の (1)、(2)… はフィールドなので自動で振り直されます
- **ページを見る** — `word_get_page_image` で、Word がレイアウトしたページの画像をモデルに渡せます。画面のキャプチャはしません
- **既定で安全** — 文書を開く間はマクロを無効にし、文書の外に手を伸ばすフィールドは許可しない限り拒否します
- **Windows と macOS** — ツール、引数、応答は同じです。Word for Mac でできないことは、ツールがそう答えます

## 📋 必要なもの

- Windows 10 / 11、または macOS
- Microsoft Word（Microsoft 365 または 2021 以降）
- [uv](https://docs.astral.sh/uv/getting-started/installation/)

macOS では、最初のツール呼び出しでシステムのオートメーションの確認が出るので、Word を一度許可してください。違いは [macOS 対応](#-macos-対応) を参照してください。

## 🚀 はじめかた

標準の設定です。Claude Desktop、Cursor、`.mcp.json` など多くの MCP クライアントで使えます。

```json
{
  "mcpServers": {
    "word": {
      "command": "uvx",
      "args": ["ms-word-mcp"]
    }
  }
}
```

### Claude Code

**ユーザー単位**（すべてのプロジェクトで使用）
```bash
claude mcp add --scope user word uvx ms-word-mcp
```

**プロジェクト単位**（`.mcp.json` に保存してチームで共有）
```bash
claude mcp add --scope project word uvx ms-word-mcp
```

### Cursor

[![Install in Cursor](https://cursor.com/deeplink/mcp-install-dark.svg)](https://cursor.com/install-mcp?name=ms-word-mcp&config=eyJ0eXBlIjoic3RkaW8iLCJjb21tYW5kIjoidXZ4IiwiYXJncyI6WyJtcy13b3JkLW1jcCJdfQ==)

または、上の標準の設定を `~/.cursor/mcp.json` に追加します。

### Claude Desktop

上の標準の設定を `claude_desktop_config.json` に書きます。場所は Windows が `%APPDATA%\Claude\`、macOS が `~/Library/Application Support/Claude/` です。

### Codex

`~/.codex/config.toml` に書きます。

```toml
[mcp_servers.ms-word-mcp]
command = "uvx"
args = ["ms-word-mcp"]
```

### VS Code

```bash
code --add-mcp '{"name":"word","command":"uvx","args":["ms-word-mcp"]}'
```

### ソースから

```bash
git clone https://github.com/ykuwai/ms-word-mcp.git
cd ms-word-mcp
uv sync
```

```json
{
  "mcpServers": {
    "word": {
      "command": "uv",
      "args": ["--directory", "/path/to/ms-word-mcp", "run", "ms-word-mcp"]
    }
  }
}
```

## 🛠️ ツールの種類

文書、読み取り、書き込み、書式、スタイル、箇条書き、表、画像・キャプション・相互参照、ページ設定、数式、浮動の図形と画像、脚注と文末脚注、コメント、スペルチェックと文書チェック、ハイパーリンクとブックマーク、フィールドと目次、表示、ページの画像、元に戻す。

全ツールの一覧は [TOOLS.md](https://github.com/ykuwai/ms-word-mcp/blob/main/TOOLS.md) にあります。コメントにはまだ対応していません。

## 💡 プロンプトの例

やりたいことを普通の言葉で伝えるだけで、コードは要りません。

**シンプルに** — テーマだけ

> *「週休 3 日制への移行を知らせる 1 ページの社内通知を書いて」*

**見た目も指定** — テーマとスタイル

> *「社内図書室を新設する 2 ページの企画書を書いて。本文は明朝、見出しは紺色で、費用の表も入れて」*

**今ある文書を直す**

> *「report.docx を開いて。見出しは太字にしただけ、箇条書きは手打ちなので、本物の見出しと箇条書きに直して目次も付けて」*

**和文の組版**

> *「このお知らせを A4 の縦書き、40 字 × 30 行で組んで、難しい名前にはルビを振って」*

**数式**

> *「2 段落目のあとに正規分布の密度関数を番号付きの数式で入れて、本文からその番号で参照して」*

**効く言葉:**

| 観点 | 言葉の例 | 効果 |
|---|---|---|
| **構成** | `本物の見出し` · `目次` · `章番号` | Word のスタイルとフィールドを使うので、ナビゲーションやページ番号が機能します |
| **見た目** | `本文は明朝` · `見出しは紺色` · `和文は游明朝` | 見た目をスタイルに、文書全体で一度だけ定義します |
| **ページ** | `A4` · `2 段組` · `40 字 × 30 行` · `下にページ番号` | 用紙、余白、段組、文書グリッドを設定します |
| **表** | `費用の表` · `見出しを結合` · `1 行目に網掛け` | すべてのセルを一度に作ってから書式を整えます |
| **画像と図形** | `図に番号を付けて` · `文字を回り込ませて` · `文字の背面に` | キャプションと相互参照を付け、画像や図形を浮動にします |
| **確認** | `各ページの見た目を確かめて` | モデルがページ画像を見てレイアウトを直します |

## 🔍 機能の詳細

### 🎯 対象の文書を決める

`word_activate_document` で、以降の呼び出しの対象を決めます。どのウィンドウが手前にあっても変わりません。この選択はサーバー全体で 1 つで、Claude Desktop などではすべての会話が 1 つのサーバーを通るので、複数の会話が同時に動くときは各呼び出しに `document` を渡してください。その呼び出しだけに効き、フルパスかファイル名を受け取ります。開いていない文書を指定するとエラーになり、別の文書に書くことはありません。

### ✍️ 隣で一緒に作業する

書き込みは Range を通すので、カーソルと選択範囲はそのままです。書き込むたびに、その文書のウィンドウが変更箇所までスクロールします。同じ文書に 2 つ目のウィンドウを開くと（表示 > 新しいウィンドウ）、後のウィンドウが追いかけ、最初のウィンドウはそのままです。どのウィンドウも動かしたくないときは `WORD_FOLLOW_EDITS=off` にします。

Windows では、対象が Word のアクティブな文書のとき、1 回の呼び出しが Word の「元に戻す」1 回分になります。そうでないときは結果の `undo` にその旨が返ります。

### ✅ 信頼できる応答

書き込むたびに Word の内容を読み戻します。`warnings` がある結果は一部だけ反映されたもので、何が反映されなかったかが書かれています。`error` の結果は、特に書かれていない限り何も変更していません。

メンバーと定数はすべて Microsoft のドキュメント（[docs/reference](https://github.com/ykuwai/ms-word-mcp/tree/main/docs/reference)）と照合し、ドキュメントに書かれていないことは実機の Word で測定しています（[docs/PROBES.md](https://github.com/ykuwai/ms-word-mcp/blob/main/docs/PROBES.md)）。

### 🌏 あらゆる文字体系

- **すべての文字に 3 つのフォント** — 欧文、東アジア文字、複雑な文字を別々に設定します
- **和文の組版** — グリッド行単位の間隔、字数単位のインデント、禁則、ぶら下げ、傍点、縦書き、`word_add_ruby` によるルビ
- **段落番号** — 全角数字、①、一、ア、가、א、α、а など

### 🎨 スタイル

`word_set_style` で文字と段落のどのプロパティもスタイルに設定できます。`word_format_text` と `word_set_paragraph_format` は同じプロパティを例外の箇所に使います。組み込みスタイルの英語名（"Normal"、"Heading 1"、"TOC 1" など）は、どの言語の Word でも使えます。

### 🧮 数式

`word_insert_equation` は LaTeX または UnicodeMath を受け取り、Word 自身の数式エディターで組み立てます。`numbered` を付けると、番号は右端にフィールドとして入るので、数式を挿入すると以降の番号が振り直されます。`word_insert_inline_equation` は段落の文中に数式を入れます（文の先頭・末尾、またはプレースホルダーの位置）。`word_get_equation` と `word_set_equation` は段落の中の何番目か（`nth`）で指定します。`chapter` を付けると、番号付きの数式とキャプション（`word_insert_caption`）が章ごとに (2.1)、図 2.1 のように振られます。章番号は直前の番号付き見出しから取ります。

### 🖼️ 図形と浮動の画像

`word_insert_shape`、`word_insert_text_box`、`word_insert_line`、`word_insert_floating_picture` で、文字の折り返し、位置、前後関係、塗り、線、テキストを指定して配置します。`word_set_shape` であとから変更でき、画像のトリミングや薄くする処理、浮動と行内の切り替えもできます。`header_footer` を指定するとヘッダーやフッターに配置でき、ロゴや罫線として全ページに表示されます（Mac では浮動の画像は [VBA ブリッジ](#vba-ブリッジ)を通します）。

### 📄 ページを見る

`word_get_page_image` は、Word がレイアウトしたとおりのページを返します。テキストを読み戻しても分からないレイアウトの崩れを、モデルが確かめられます。Windows では Word から直接ページを取得し、macOS では Word が一時的に保存した PDF から描きます。どちらも画面はキャプチャしません。

## ⚙️ 詳細設定

| 変数 | 既定値 | 内容 |
|---|---|---|
| `WORD_TOOL_MODE` | `catalog` | `catalog` は残りのツールに届く 3 ツールだけを一覧に出し、`full` は全ツールを出します。 |
| `WORD_FOLLOW_EDITS` | `on` | 変更のたびに文書のウィンドウをスクロールします。 |
| `WORD_ALLOW_ACTIVE_FIELDS` | `off` | 文書の外に手を伸ばすフィールドを許可します（INCLUDETEXT、INCLUDEPICTURE、INCLUDE、LINK、IMPORT、RD、EMBED、DATABASE、DDE、DDEAUTO、MACROBUTTON、PRINT、ASK、FILLIN）。オフのときは `word_insert_field` が拒否し、`word_update_fields` は文書にあるものをそのままにします。オンにすると Word 自身のダイアログで確認されることがあり、ツールは答えるまで待ちます。 |
| `WORD_ALLOW_NETWORK_PATHS` | `off` | ネットワークパス（`\\server\share\...`）のファイルを許可します。オフのときは拒否します。そこへ保存すると文書がそのサーバーに送られ、ファイルの有無を確かめるだけでも資格情報を使って接続するためです。 |
| `WORD_VBA_BRIDGE` | `on` | サーバーの VBA ブリッジを Word に読み込みます。段落の読み取りと検索・置換を 1 回の呼び出しで行うマクロと、Mac ではコメントの返信と解決済みの状態を扱うマクロです。オフのときは従来どおり（より遅く）読み、Mac ではコメントへの返信と解決を拒否します。[VBA ブリッジ](#vba-ブリッジ)を参照してください。 |
| `WORD_MAC_FILES_DIR` | 文書のフォルダー | macOS: ページ画像の元になる PDF を Word が一時的に書き出す場所。既定は文書のフォルダーで、未保存の文書では Word 自身のフォルダーです。 |

サーバーの `env` に設定します。

```json
{
  "mcpServers": {
    "word": {
      "command": "uvx",
      "args": ["ms-word-mcp"],
      "env": { "WORD_FOLLOW_EDITS": "off" }
    }
  }
}
```

### ツール一覧: カタログモード

クライアントは、一覧にあるツールの説明をリクエストのたびに送ります。このサーバーのツールの説明は合わせて約 18 万文字あります。既定では代わりに 3 ツール（約 2,500 文字）だけを一覧に出し、モデルはそれを通して残りのツールを使います。

| ツール | 内容 |
|---|---|
| `word_describe` | 名前なしで呼ぶと全ツールを 1 行ずつ一覧にします。名前を渡すと、その説明と引数を返します。 |
| `word_run` | ツールを 1 つ実行します: `{"tool": "word_insert_paragraphs", "params": {...}}` |
| `word_run_steps` | 複数のツールを 1 回の呼び出しで順に実行します。前のステップの結果を使えます。`"$prev.key"` は直前のステップの `key`、`"$2.inserted.0.index"` は 2 番目のステップの値です。最初に失敗したステップで止まり、理由を返します。 |

ツール自体は同じです。カタログモードで失われるのは、クライアント側のツール単位の制御です。名前によるツールの承認やブロックは、各ツールではなく上の 3 ツールに対して効きます。それが必要なクライアントでは `WORD_TOOL_MODE=full` にしてください。

### Word が応答できないとき

呼び出しは、頼まれた順に 1 つずつ Word に届きます。Word でダイアログやメニューが開いていると、呼び出しは閉じるのを最大 15 秒待ち、そのあと Word がビジーだと答えます。実行が Windows で 45 秒、macOS で 69 秒を超えた呼び出しは打ち切り、サーバーは次の呼び出しに進みます。大量の検索と置換など時間のかかるツールは、時間内に止めて `complete: false` と警告を返すので、もう一度呼び出せます。

Word が起動していないときは、`word_create_document` と `word_open_document` が起動します。他のツールは起動していないと答えます。

### VBA ブリッジ

COM や Apple Events で段落を読むと、プロパティごとに 1 回 Word に問い合わせるので、200 段落に 8〜27 秒かかっていました。サーバーには Word のテンプレート `ms-word-mcp-bridge.dotm` が含まれ、そのマクロが Word の中で 1 回の呼び出しで読みます。マクロのソースは [`WordMcpBridge.bas`](../src/ms_word_mcp/backend/vba/WordMcpBridge.bas) で、テンプレートがそこから作られたことをテストで確かめています。

| `word_get_paragraphs`、200 段落 | ブリッジなし | ブリッジあり |
|---|---|---|
| Windows | 8.1 秒、`detail` 付きで 11.7 秒 | 0.28 秒、`detail` 付きで 0.30 秒 |
| macOS | 6.9 秒、`detail` 付きで 26.9 秒 | 0.45 秒、`detail` 付きで 0.53 秒 |

`word_find_replace` もブリッジを通ります。3,000 件の一致を 1 回の呼び出しで置換し、Windows で 2.4 秒、macOS で 7.4 秒です（以前は 1 回で 740 件と 1,249 件まででした）。macOS ではワイルドカード検索もできるようになり、置換全体が「元に戻す」1 回分になります。

答えは同じで、Mac の `detail` は Windows と同じ内容になります（太字が混在する段落は混在、テーマの色と自動の色はなし）。Mac ではさらに、Word for Mac のスクリプトでは届かないもの（コメントの返信と解決済みの状態、表の列の追加、スタイルの新規作成、段落の方向とタブ位置、複雑な文字のサイズと下線の色、図形の代替テキストと文字の上下の配置）もブリッジで扱います。

テンプレートは、必要になった最初の呼び出しで、サーバーが使う Word のグローバルテンプレートになり、サーバーが止まるときに取り除かれるので、Word は元のままです。Windows ではパッケージから `AddIns.Add` で追加します。Word for Mac のスクリプトでは追加できないので、テンプレートを Word 自身のフォルダーにコピーして一瞬開き（テンプレートから新しい文書ができます）、その文書のマクロがテンプレートを追加してから文書を閉じます。ポリシーでマクロが無効になっているなどでブリッジが使えないときは、ツールは従来どおりに読み、`word_get_app_info` の `vba_bridge` が理由を答えます。

### 🔒 セキュリティ

サーバーはあなたのコンピューター上であなたの権限で動き、stdio で 1 つのローカルクライアントと通信します。ポートは開かず、サーバー自身がネットワークにアクセスすることもありません。

これらのツールを使う AI は、Word で開いているどの文書でも読み書きでき、Word が開けるファイルなら何でも開け、あなたのアカウントで読み書きできる場所にファイルを保存したり埋め込んだりできます。文書や Web ページ、他のツールの応答に含まれる文章が AI を誘導しようとすることもあります。他の人から受け取った文書を扱うときは AI の動作を確かめ、必要に応じてツールの呼び出しを承認してください。

サーバーが行う対策:

- 文書を開くとき、テンプレートから作るときはマクロを無効にします。そうしないと Word が実行してしまうためです。文書はマクロが無効のまま開いた状態になります。
- 実行するマクロは、パッケージにあるブリッジのテンプレートの自前のものだけで、ツールが頼むこと（段落の読み取り、検索と置換、コメント、Mac ではスクリプトで届かない書式の設定）しかしません。`WORD_VBA_BRIDGE=off` で使わないようにできます。
- 文書の外に手を伸ばすフィールドは、`WORD_ALLOW_ACTIVE_FIELDS` を設定しない限り拒否し、触れません。
- ネットワークパスは、`WORD_ALLOW_NETWORK_PATHS` を設定しない限り拒否します。
- 新しいリンクの宛先は http、https、mailto、tel、または文書からの相対パスに限ります。文書に元からあるリンクはそのままです。

行わないこと: これらのツールで開いた文書には Word の保護ビューが適用されません。自分でも開くような文書だけを開いてください。

脆弱性の報告は [SECURITY.md](https://github.com/ykuwai/ms-word-mcp/blob/main/SECURITY.md) を参照してください。

## 🍎 macOS 対応

ツール、引数、応答は同じです。違いは Word for Mac のスクリプトで届く範囲で、届かないものはツールがそう答えます。

Word for Mac のスクリプトでは届かないため、[VBA ブリッジ](#vba-ブリッジ)を通してだけ使えるもの（ブリッジがないときはツールがそう答えます）:

- `word_insert_table_columns`、スタイルの新規作成、スタイルと番号付けのリンク
- 段落の方向（右から左）とタブ位置
- 複雑な文字のサイズと、下線の色
- スタイルの `quick_style` と `priority`
- 浮動の図形の代替テキストと、図形内の文字の上下の配置
- 図形のグループ化と解除（`word_group_shapes`、`word_ungroup_shape`）と、`word_list_shapes` のグループの `members`
- ヘッダーやフッターへの浮動の画像
- コメントの返信と解決済みの状態、ワイルドカード検索、文書の最後の数式の後への文章の追加

細かな違い:

- 変更ごとに「元に戻す」1 回分になります。Word for Mac のスクリプトではまとめられないためです。ツールの結果にその旨が返ります。[VBA ブリッジ](#vba-ブリッジ)を通した置換は 1 回分です。
- [VBA ブリッジ](#vba-ブリッジ)の初回の読み込みではテンプレートから作った文書が 1 秒ほど開き、Word 自身のフォルダーにテンプレートとロックファイルが残ります。
- Word for Mac はサンドボックス内で動くので、確認なしに書き込めるのは `~/Library/Containers/com.microsoft.Word/Data/Documents` の中だけです。それ以外への保存では Word がアクセスの許可を求めることがあります。
- Word は自身のフォルダーの外にあるファイルごとに確認するので、`word_insert_picture` は Word のフォルダーに一時的にコピーを置いて読み込みます。初回は、サーバーを動かしているアプリが Word のデータを使ってよいか macOS が確認することがあります。
- `word_get_page_image` では Word が文書の隣に PDF を一時的に保存します。未保存の文書では Word 自身のフォルダーに置き、macOS が同じ確認をすることがあります。
- Word for Mac はアクティブな文書にしか表を作れないので、表を挿入するときはその文書を一時的に手前に出し、元の文書に戻します。

## 🤝 コントリビュート

プルリクエストを歓迎します。`develop` をベースに、1 トピックずつ、テスト付きでお願いします。

```bash
uv run pytest -q                    # Word は不要
uv run python scripts/live_check.py # 実機の Word の作業用文書で全ツールを実行
```

live check は文書を作り、ツールを実行して各ステップの PASS か FAIL を表示し、PDF を書き出して文書を閉じます。他の文書には触れません。

Windows では、クライアントがサーバーを動かしている間は `uv sync` が `.venv\Scripts\ms-word-mcp.exe` を置き換えられません。テストは `uv run --no-sync` で実行するか、先にサーバーを止めてください。

## 📄 ライセンス

MIT

## 🙏 クレジット

- [ppt-mcp](https://github.com/ykuwai/ppt-mcp) — このサーバーの設計のもと
- [MCP Python SDK](https://github.com/modelcontextprotocol/python-sdk) — MCP サーバーのフレームワーク
- [pywin32](https://github.com/mhammond/pywin32) — Windows の COM 操作
- [appscript](https://github.com/hhas/appscript) — macOS の Apple Event ブリッジ
- [Model Context Protocol](https://modelcontextprotocol.io/) — Anthropic による
