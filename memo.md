## 用途

- VBAソースコードをGitで管理する
- VBAソースコードのドキュメントをDoxygenで出力する

## 基本的な使い方

### 初めにすること

1. `git clone https://github.com/lie/ariawase.git project-name`
2. `cd project-name`
3. `/bin/` 配下に `*.xlsm` ファイルを置く

### VBEでソースコードを編集する場合

1. VBEで `/bin/*.xlsm` 内のVBAソースコードを作成、編集、保存
2. `cscript //nologo vbac.wsf decombine # vbacでエクスポート` 
3. `git add src/*`
4. `git commit -m "commit message"`
5. 1～4の繰り返し

### エクスポートされたVBAソースファイルを他のエディタで編集する場合

1. `/src/*.xlsm/` 内のVBAソースコードを作成、編集、保存
2. `cscript //nologo vbac.wsf combine # vbacでインポート` 
3. `git add src/*`
4. `git commit -m "commit message"`
5. 1～4の繰り返し

## vbacについて

### コマンドオプションの説明

| コマンド | 説明 |
| :------- | :--- |
| `cscript //nologo vbac.wsf decombine` | `/bin/hoge.xlsm` から `/src/hoge.xlsm/*` へエクスポート |
| `cscript //nologo vbac.wsf combine`   | `/src/hoge.xlsm/*` から `/bin/hoge.xlsm` へインポート   |

### 注意

- `decombine` を実行すると、`/src/` 内のソースファイルがすべて上書きされる。また、 `.xlsm` ファイル内に存在しないファイルはすべて削除される。
- `combine` を実行すると、`.xlsm` ファイル内のソースファイルが全て上書きされる。また、 `/src/` 内に存在しないソースファイルは全て削除される。
- そのため、ソースファイルの編集は、`.xlsm` ファイル内か、`/src/` ディレクトリ内のどちらかのみで行うことが望ましい。

## 参考

- [いげ太の日記: Ariawase v0.6.0 解説（vbac 編）](http://igeta-diary.blogspot.com/2014/03/what-is-vbac.html)