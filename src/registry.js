/**
 * インスタンス化しません。
 * @class レジストリにアクセスする機能を提供するクラス
<pre class = "code">
使用例：
// IP Messengerのユーザ名設定を取得
print(Registry.regRead("HKCU\\Software\\HSTools\\IPMsg\\NickNameStr"));

// "NickNameStra"に"hoge"を設定
Registry.regWrite("HKCU\\Software\\HSTools\\IPMsg\\NickNameStra", "hoge", Registry.REG_TYPE_REG_SZ);

// "NickNameStra"を削除
Registry.regDelete("HKCU\\Software\\HSTools\\IPMsg\\NickNameStra");
</pre>
 */
var Registry = {};

/**
 * レジストリのデータ型：REG_SZ
 */
Registry.REG_TYPE_REG_SZ = "REG_SZ";

/**
 * レジストリのデータ型：REG_DWORD
 */
Registry.REG_TYPE_REG_DWORD = "REG_DWORD";

/**
 * レジストリのデータ型：REG_BINARY
 */
Registry.REG_TYPE_REG_BINARY = "REG_BINARY";

/**
 * レジストリのデータ型：REG_EXPAND_SZ
 */
Registry.REG_TYPE_REG_EXPAND_SZ = "REG_EXPAND_SZ";

/**
 * レジストリ内のキー名または値名の値を返します。
 * @param {String} name 読み取るキーまたは値の名前です。
 * @return {Object} レジストリの値
 */
Registry.regRead = function(name) {
  return Const.WSHELL.RegRead(name);
};

/**
 * 新しいキーの作成、新しい値名の既存キーへの追加 (および値の設定)、既存の値名の値変更などを行います。
 * @param {String} name 作成、追加、変更するキー名、値名、または値を示す文字列値です。
 * @param {String} value 作成するキーの名前、既存のキーに追加する値の名前、または既存の値名に設定する値です。
 * @param {String} type (オプション)レジストリに保存する値のデータ型を指定します。
 */
Registry.regWrite = function(name, value, type) {
  var regType = type || Registry.REG_TYPE_REG_SZ;
  Const.WSHELL.RegWrite(name, value, regType);
};

/**
 * レジストリから指定されたキーまたは値を削除します。
 * @param {String} name レジストリ内で削除するキーまたは値の名前を示す文字列値です。
 */
Registry.regDelete = function(name) {
  Const.WSHELL.RegDelete(name);
};
