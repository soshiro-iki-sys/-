# reference/ — 過去資料への参照

過去の優れた営業資料を **すぐ参照できるように** しておく場所です。
実体をコピーするのではなく、**symlink（近道）** を置くのが基本です。

## 過去資料への近道を作る

```bash
# 例：過去のAddness資料フォルダへの symlink を作る
ln -s "/path/to/ADDNESS-Sequoia-Materials-2026-05-02" reference/ADDNESS-Sequoia-Materials-2026-05-02
```

これで `reference/ADDNESS-Sequoia-Materials-2026-05-02` から過去資料を即参照できます。

## 使い方のお作法
- 参照は **構成・言い回し・型を学ぶため**。中身の丸写しはしない。
- symlink の実体（過去資料そのもの）はこのリポジトリには含めない。

> 過去資料の置き場所（パス）を教えてもらえれば、私（Claude）が symlink を作成します。
