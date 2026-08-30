# WinEvtView

WinEvtView は、Windows の System / Application イベントログをすぐに確認するための、
PowerShell + WinForms 製の簡易イベントビューアです。

「コンピュータの管理」やイベントビューアを探して開き、ログを選び、フィルタを設定する手間を省き、
起動した時点で「直近に何が起きているか」が見える状態にすることを目的としています。

イベントビューアの完全な代替ではなく、障害調査の初動で状況を短時間で把握するための軽量ツールです。

## 主な用途

- PC の調子がおかしいときに、直近のエラーや警告をまず確認する
- 特定のイベント ID（例: 41, 1000, 6008）が出ていないかを素早く調べる
- 確認した内容を CSV に出して、記録・共有する
- さらに詳しく追う必要があれば、そのまま本家イベントビューアを開く

## 対応環境

- Windows 11（Windows 10 でも動作します）
- Windows PowerShell 5.1
- 追加のモジュール、外部ライブラリ、ネットワーク接続は不要です
- 管理者権限は不要です（System / Application ログは通常のユーザーで参照できます）

## 主な機能

- 対象ログの切り替え：System + Application / System のみ / Application のみ
- 期間の切り替え：直近 1 時間 / 直近 24 時間 / 直近 7 日
- レベル：Critical / Error / Warning を表示（1 回の検索で最大 500 件）
- イベント ID による絞り込み（半角数字。空欄ならすべての ID が対象）
- 一覧で選んだイベントの詳細（メッセージ全文）を下部に表示
- CSV 出力（デスクトップに `WinEvtView_yyyyMMdd_HHmmss.csv` として保存）
- 本家イベントビューア（eventvwr.msc）の起動

起動すると「System + Application / 直近 1 時間 / Critical・Error・Warning」で自動的に検索します。
対象ログや期間を変更すると、その場で再検索します。イベント ID を入力したときは
「検索」ボタン（または Enter キー）で検索します。

条件に一致するイベントが無い場合は、エラーではなく「0 件」として表示されます。

## 実行方法

1. `winevtview.ps1` を任意のフォルダに置きます。
2. PowerShell を開き、次のコマンドを実行します。

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\winevtview.ps1
```

エクスプローラーからファイルを右クリックして「PowerShell で実行」しても起動できます。

## 注意事項

- 本ツールはイベントログの**閲覧のみ**を行います。イベントログの変更・削除・クリアは行いません。
- 対象は System / Application ログです。Security ログ、リモート PC のログ取得、常駐監視には対応していません。
