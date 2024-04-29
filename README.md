# Outlook_bcc
Outlookでメール送信時にBCCを追加して送信する。
Office2021で実装したものです。

# 証明書を発行
以下にアクセスして実行し、証明書を発行する。<br>
証明書の名前はお好きな名前を入力してください。
```
C:\Program Files\Microsoft Office\root\Office16
SELFCERT.EXE
```

# Outlookのマクロ検出レベルの変更
トラストセンター　→　マクロの設定<br>
「すべてのマクロに対して警告を表示する」

# 作成した電子証明書を付与
開発タブのVisual Basicをクリックする。<br>
ツール → デジタル署名 → 選択をクリックして自分が作成した署名を選んでOKをクリックする。

# コードの入力
ThisOutlookSessionにコードを入力する。<br>
BCCで送信したいメールアドレスはコレクションに設定してください。<br>
保存して一度Outlookを終了する。

# Outlookの起動と送信
Outlook起動時は必ずこれが表示されます。<br>
「マクロを有効にする」をクリックしないとプログラムが実行されません。<br>
※「すべてのマクロを有効にする」にすると表示されなくなりますが、悪意のあるコードが意図しないときに実行された際に防げなくなりますので推奨しません。
![1](https://github.com/Sonoda-a/Outlook_bcc/assets/13926852/0123deac-9c03-4725-9007-53ec5131bd31)

メールを送信してBCCのアドレスで受信できているかを確認してください。

# プログラムコードのロック
現在の状態だと、Visual Basicをクリックするとプログラムコードが見えてしまいます。<br>
見られたくない場合は保護することができます。<br>
<ol>
    <li>ThisOutlookSessionを右クリック → Project1のプロパティ → 保護</li>
    <li>プロジェクトを表示用にロックするにチェックを付ける。</li>
    <li>パスワードを入力する。</li>
    <li>Outlookを再起動する。</li>
</ol>
