# 36協定届 作成ツール（36kyotei-tool）

ブラウザだけで **時間外労働・休日労働に関する協定届（36協定届）** を作成し、**Word（.docx）形式でダウンロード**できる無料ツールです。
インストール不要・サーバー不要。入力したデータは**どこにも送信されず、すべてブラウザ内で完結**します。

> 様式第9号（一般条項）と様式第9号の2（特別条項）の両方に対応しています。

[**▶ オンラインで使う**](#-オンラインデモ) ・ [English](#english)

---

## ✨ 特長

- 📄 **Word出力** — 様式第9号 / 様式第9号の2 を `.docx` で生成（そのまま印刷・提出可）
- 🗂 **ZIP一括ダウンロード** — 2つの様式をまとめて取得
- 🔒 **完全ブラウザ完結** — データは一切外部送信されない（労務情報を安心して扱える）
- 💾 **保存 / 読み込み** — 入力内容を JSON で保存し、後から再編集
- 🧭 **上限ガイド付き** — 月45時間 / 年360時間、特別条項の上限（月100時間未満・年720時間以下・複数月平均80時間以下・年6回以内）を画面に表示
- 🌙 **ダークモード対応**
- ⚙️ **依存ゼロのビルドレス構成** — HTML / CSS / JS の3ファイルのみ

## 🖥 使い方

1. `index.html` をブラウザで開く（またはオンラインデモにアクセス）
2. 事業場情報・業務・労働時間などを入力
3. 「📄 様式第9号」または「📄 様式第9号の2（特別条項）」をクリックして Word を生成
4. 必要に応じて「⬇ ZIP一括」でまとめてダウンロード

入力途中の内容は「💾 保存」で JSON に書き出し、「📂 読み込み」で復元できます。

## 🚀 ローカルで動かす

クローンして `index.html` を開くだけです。

```bash
git clone https://github.com/lucymak-lab/36kyotei-tool.git
cd 36kyotei-tool
# そのまま index.html をブラウザで開く
```

ローカルサーバーで開きたい場合：

```bash
python3 -m http.server 8000
# http://localhost:8000 を開く
```

## 🌐 オンラインデモ

> 公開後にURLをここへ記載します（GitHub Pages 予定）。

## 🛠 技術構成

- 純粋な HTML / CSS / バニラ JavaScript（フレームワーク・ビルド不要）
- [docx](https://github.com/dolanmiu/docx) v7.8.2 — Word(.docx) 生成
- [JSZip](https://github.com/Stuk/jszip) v3.10.1 — ZIP一括ダウンロード

外部ライブラリは `vendor/` に**同梱**しています。CDNに依存せず、**ネット接続なし（完全オフライン）でも動作**します。これによりサプライチェーン経由でのコード改ざんリスクを排除し、入力した労務データが外部に漏れる経路を一切持たない構成になっています。

## ⚠️ 免責事項

本ツールは36協定届の作成を補助するものであり、内容の法的正確性・最新の法令適合を保証するものではありません。
実際の届出にあたっては、**最新の様式・記載要領を厚生労働省／所轄の労働基準監督署で必ずご確認ください**。本ツールの利用により生じたいかなる損害についても、作者は責任を負いません。

## 🤝 コントリビュート

不具合報告・改善提案は Issue / Pull Request で歓迎します。

## 📄 ライセンス

[MIT License](LICENSE)

---

## English

### 36 Agreement Form Maker (Japanese labor law "36協定届")

A free, browser-only tool to fill in and export Japan's **"36 Agreement" overtime/holiday-work notification forms** as **Word (.docx)** files.
No install, no server — all data stays in your browser and is **never sent anywhere**.

Supports both Form No. 9 (general clause) and Form No. 9-2 (special clause).

**Features:** Word export, bundled ZIP download, fully client-side (privacy-safe), JSON save/load, statutory-limit guidance, dark mode, zero build step (3 files only).

**Run locally:** clone the repo and open `index.html` in a browser.

**Disclaimer:** This tool assists in preparing the forms but does not guarantee legal accuracy. Always verify the latest official forms and instructions with the Japanese Ministry of Health, Labour and Welfare / your local Labour Standards Inspection Office.

**License:** [MIT](LICENSE)
