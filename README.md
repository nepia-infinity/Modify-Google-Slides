# 概要
- Pythonを使って、Google Slidesを操作することを学ぶためのスクリプト
- GCPの設定に苦戦

## Google Slides API 利用時の注意点

Google Slides API を利用する際は、以下の点に注意してください。

1.  **OAuth 2.0 クライアント ID:** デスクトップ アプリケーション用を使用 (ウェブアプリ用は `400` エラーの原因)。
2.  **認証情報:** ダウンロードしたファイルは `credentials.json` にリネーム。
3.  **ファイルパス:** スクリプト内で `credentials.json` をフルパス指定。
4.  **機密情報 (最重要):** `credentials.json` と `token.json` は `.gitignore` に追加し、Git にコミットしない
5.  **環境変数**や**シークレット管理ツール**で安全に管理

**公式ドキュメント:** Google API の仕様変更に注意し、公式ドキュメントも参照してください:

*   [Quick Start](https://developers.google.com/slides/api/quickstart/python?hl=ja)
*   [GitHubの関連リポジトリ](https://github.com/googleworkspace/python-samples/blob/main/slides/quickstart/quickstart.py)

## ブログなどの関連リンクなど
*   [サンプルのGoogle Slides](https://docs.google.com/presentation/d/1GButmtqvj5LT8TzLexFRnA1-5tDsW5ft-S7E4_Sy8AE/edit?usp=sharing)
*   [note　スライドにスプレッドシートの内容を挿入するスクリプト](https://note.com/nepia_infinity/n/nfe7a2f763655)

![image](https://github.com/user-attachments/assets/1c70ce77-e8b7-49d3-812e-00fb03772345)
  
| Slide Page | Slide ID | Element | Element ID | Element Text |
|---------|----------|-----------|------------|--------------|
| 1 | gcb9a0b074_1_0 | 1 | gcb9a0b074_1_1 | {title}  ・{people1} ・{people2} ・{people3} ・{people4} |
| 2 | ge965474a9_3_282 | 1 | ge965474a9_3_301 | 2015 年 8 月 |
| 2 | ge965474a9_3_282 | 2 | ge965474a9_3_304 | アプリ内でテキストを翻訳 |
| 2 | ge965474a9_3_282 | 3 | ge965474a9_3_303 | 2015 年 10 月 |
| 2 | ge965474a9_3_282 | 4 | ge965474a9_3_283 | マイルストーン |
| 2 | ge965474a9_3_282 | 5 | ge965474a9_3_284 |  |
| 2 | ge965474a9_3_282 | 6 | ge965474a9_3_285 |  |
| 2 | ge965474a9_3_282 | 7 | ge965474a9_3_299 | 2014 年 10 月 |
| 2 | ge965474a9_3_282 | 8 | ge965474a9_3_300 | Chrome 拡張機能でウェブページを翻訳 |
| 2 | ge965474a9_3_282 | 9 | ge965474a9_3_302 | Android 搭載時計で会話を翻訳 |
| 2 | ge965474a9_3_282 | 10 | ge965474a9_3_305 | 2015 年 11 月 |
| 2 | ge965474a9_3_282 | 11 | ge965474a9_3_306 | カメラアイコンのタップで英語やドイツ語のテキストをアラビア語に翻訳 |
| 2 | ge965474a9_3_282 | 12 | ge965474a9_3_307 |  |
| 2 | ge965474a9_3_282 | 13 | ge965474a9_3_308 |  |
| 2 | ge965474a9_3_282 | 14 | ge965474a9_3_309 |  |
| 3 | SLIDES_API1490328969_0 | 1 | SLIDES_API1490328969_1 | 廊下  ・スネ夫 ・のび太 ・ジャイアン ・ドラえもん |
| 4 | SLIDES_API1490328969_4 | 1 | SLIDES_API1490328969_5 | 窓  ・のび太 ・ジャイアン ・しずか ・スネ夫 |
| 5 | SLIDES_API1490328969_8 | 1 | SLIDES_API1490328969_9 | 床  ・ジャイアン ・しずか ・ドラえもん ・のび太 |
| 6 | SLIDES_API1490328969_12 | 1 | SLIDES_API1490328969_13 | デスク  ・しずか ・ドラえもん ・スネ夫 ・ジャイアン |
| 7 | SLIDES_API1490328969_16 | 1 | SLIDES_API1490328969_17 | 掃除機  ・ドラえもん ・スネ夫 ・のび太 ・しずか |

- SLIDES_API1490328969_0　この部分はGASで加筆修正した部分

![image](https://github.com/user-attachments/assets/909caded-236a-48a0-b76b-f51be4d91f93)

## コマンドなど
```
cd C:\Users\nepia\OneDrive\デスクトップ\Google_Slides
```
カレントディレクトリに移動

```
python -m venv venv
```
仮想環境を作成。1回しか使わない

```
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
```
Visual Studio Codeのコマンドでエラーが出ないようにする

```
.\venv\Scripts\Activate.ps1
```
仮想環境を有効にする

```
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
```
仮想環境に最新のライブラリをインストールする

```
tree /F | clip
```
- コマンドプロンプトでフォルダ構造などを取得するコマンド
- 仮想環境を出力したら14万文字あった....。
