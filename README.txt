=概要=
Windows用デスクトップJSライブラリ。<br/>
JScript/Windows Scripting Hostにruby風の味付けをほどこし、
使いやすさの向上を目指すプロジェクト。

<p>
<a href = "http://code.google.com/p/wshf/">wshf</a>を参考にしています。
</p>

=クイックスタート=
  * http://code.google.com/p/madojs/downloads/list からアーカイブされた最新版をダウンロードします。
  * ダウンロードしたzipファイルを解凍し、test.wsfをダブルクリックします。
  * test.jsに記述されたプログラムによって、mado.jsがcopy_of_mado.jsにコピーされます。

=機能一覧=
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

=インストール=
==1.zipアーカイブをダウンロードする==
http://code.google.com/p/madojs/downloads/list
からダウンロードする。

==2.リポジトリから最新版をダウンロードする==
以下のコマンドを実行してください。

{{{
svn checkout http://madojs.googlecode.com/svn/trunk/ madojs-read-only
}}}

zipアーカイブは古い場合があります。
最新版は、リポジトリから取得してください。

=実行方法=
実行には三つのファイルが必要です。

  * mado.js (ライブラリ本体)
  * test.js (プログラム本体)
  * test.wsf (実行可能ファイル; HTMLに相当)

test.wsfをダブルクリックすると、test.jsに記述したプログラムが実行されます。

=ドキュメント=
<p>
<a href = "http://madojs.googlecode.com/svn/trunk/docs/index.html">API リファレンス</a>
</p>