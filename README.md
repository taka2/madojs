#概要
Windows用デスクトップJSライブラリ。<br/>
JScript/Windows Scripting Hostにruby風の味付けをほどこし、
使いやすさの向上を目指すプロジェクト。

<p>
<a href = "http://code.google.com/p/wshf/">wshf</a>を参考にしています。
</p>

#クイックスタート
  * https://github.com/taka2/madojs/zipball/gh-pages からアーカイブされた最新版をダウンロードします。
  * ダウンロードしたzipファイルを解凍し、examples/file/filecopy.wsfをダブルクリックします。
  * filecopy.jsに記述されたプログラムによって、filecopy.jsがcopy_of_filecopy.jsにコピーされます。

#機能一覧
  * ファイルI/O
  * フォルダ操作、フォルダ内の再帰的なファイル検索
  * ドライブ情報の取得
  * プログラムの実行
  * HTTP GET/POST
  * FTP GET/PUT
  * クリップボード
  * サービスの開始と終了
  * zipファイルの展開と圧縮
  * レジストリへのアクセス
  * イベントログの書き込み
  * 特殊フォルダパスの取得
  * Excelへのアクセス
  * ADOを使ったOracle/MSAccess/ODBCデータソースへのアクセス
  * String/Array/Date組み込みクラスの拡張

#インストール
##1.zipアーカイブをダウンロードする
https://github.com/taka2/madojs/zipball/gh-pages
からダウンロードする。

##2.リポジトリから最新版をダウンロードする
以下のコマンドを実行してください。

   git clone git://github.com/taka2/madojs.git madojs

#実行方法
実行には三つのファイルが必要です。

* mado.js (ライブラリ本体)
* filecopy.js (プログラム本体)
* filecopy.wsf (実行可能ファイル; HTMLに相当)

filecopy.wsfをダブルクリックすると、filecopy.jsに記述したプログラムが実行されます。

#ドキュメント
<p>
<a href = "http://taka2.github.com/madojs/">API リファレンス</a>
</p>