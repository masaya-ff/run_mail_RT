# ①メール受信ルーティン

「メール受信ルーティン.xlsx」に登録したフォルダの .bat を順に実行する入口。全体の共通ルールは親フォルダの CLAUDE.md を参照。

## ファイル

- `run_mail_routine.py` … 本体
- `run_mail_routine.bat` … 手動実行用（UTF-8、`chcp 65001`）
- `メール受信ルーティン.xlsx` … 実行対象の一覧

```
python run_mail_routine.py
```

## 動き

- Excel 1枚目のシート、3行目から：B列「〇」（○◯も可）の行が対象、C列=名前、D列=フォルダ。
- D列フォルダ内の `*.bat` を名前順にすべて `cmd /c` で実行（cwd はそのフォルダ）。
- 各 bat 末尾の `pause` で止まらないよう標準入力に改行を渡している。
- 実行した行の E列に「最終実行日時」を書く。Excel が開いていれば閉じるまで待つ（n で記録しない）。
- 登録先は各案件フォルダの save_mail.py（メール本文PDF保存）と、最後に ★AI_日常業務 直下の update_mail_todo.bat（TODO更新）。

## 注意

- この Excel の B〜D列の並びと3行目開始は update_mail_todo.py（`load_folders`）も前提にしている。列構成を変えるときは両方直す。
- 登録フォルダにある .bat はすべて実行される。登録フォルダに余計な .bat を置かない。
