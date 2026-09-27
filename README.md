# Excel Diff

Excel / CSVを2つ選ぶだけで差分を比較できる、ブラウザ完結の無料ツールです。

- XLSX / XLS / CSV比較
- セル変更、行追加・削除
- 数式変更、基本書式変更
- 空白・大文字小文字の無視
- 比較対象外列、主キー指定
- 差分だけ表示
- 差分Excel / HTML / PDF（印刷）出力
- ファイルアップロードなし（ブラウザ内処理）

## GitHub Pages

リポジトリの **Settings → Pages** で `Deploy from a branch` を選び、`main / root` を指定すると公開できます。

## 技術

静的HTML + JavaScript。Excel解析には SheetJS (`xlsx`) をCDNから利用しています。
