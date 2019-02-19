## ToDo

- 別のリポジトリ [`vba-mng`](https://github.com/lie/vba-mng) に移行
	- [`ariawase`](https://github.com/vbaidiot/ariawase) と [`doxygen-vb-filter`](https://github.com/sevoku/doxygen-vb-filter) を Fork して、submodule にする
	- `vbfilter` のファイルは、submodule からコピー
- `Makefile` を追加
	- `init.sh` と同じ機能を `make init <project_name>` で実行できるようにする
	- `make rm <project_name>` でプロジェクトを削除する機能を追加
	- `make doc <project_name>` でDoxygen文書を生成する機能を追加
- `vbfilter` について
	- [`vbfilter`のWiki](https://github.com/sevoku/doxygen-vb-filter/wiki/Documenting-your-sources)を参考に、使い方をまとめる
	- 別の `INPUT_FILTER` も検討してみる