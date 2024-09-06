# replace-line-in-VBAProject-module
パスワードで保護されたVBAProjectのDeclare を Declare PtrSafe へ置き換える。

## 事前準備

VBAを使用してVBAコードを置換するには
　・「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェック
を入れるておく必要があります。

## 手順

1. [ファイル]→[その他]→[オプション]をクリックする。
1. [トラスト センター]の「トラスト センターの設定」をクリックする。
1. [マクロの設定]をクリックする。
1. 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェックを入れて「OK」をクリック
1. 「OK」をクリック

## 実行

```
cscript .\Make.vbs
```

自動作成されるbatファイルを実行する。

## xlsmをxstmへ変換

```
cscript .\SaveAsXltm.vbs

dir /b /s | findstr ".xlsm$" > temp.txt
for /f %%f in (temp.txt) do del "%%f"
del temp.txt
```

## 参考

[【EXCEL】VBA-マクロのパスワード解除方法【コード有】](https://nkmrdai.com/vba-password-unrocked/)

[Office の 32 ビット バージョンと 64 ビット バージョン間の互換性](https://learn.microsoft.com/ja-jp/office/client-developer/shared/compatibility-between-the-32-bit-and-64-bit-versions-of-office)