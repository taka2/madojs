/** 
 * ドライブの存在チェックを行った上で、ドライブオブジェクトを作成します。
 * @class ドライブの操作を行うためのクラスです。
<pre class = "code">
使用例：
// Cドライブの空き容量を表示
var d = new Drive("c:");
print(d.getFreeSpace());
</pre>
 * @param  {String} path ディレクトリのパスを文字列で指定します。
 * @throws ディレクトリが存在しない場合にスローされます。
 */
var Drive = function(path) {
  if(!Drive.exist(path)) {
    throw new Error("Drive " + path + " not found.");
  }
  this.path = path;
  this.driveObj = Const.FSO.GetDrive(path);
};

/**
 * ドライブタイプの定数：不明(0)
 */
Drive.DRIVE_TYPE_UNKNOWN = 0;

/**
 * ドライブタイプの定数：リムーバブルディスク(1)
 */
Drive.DRIVE_TYPE_REMOVABLE_DISK = 1;

/**
 * ドライブタイプの定数：ハードディスク(2)
 */
Drive.DRIVE_TYPE_HARD_DISK = 2;

/**
 * ドライブタイプの定数：ネットワークドライブ(3)
 */
Drive.DRIVE_TYPE_NETWORK_DRIVE = 3;

/**
 * ドライブタイプの定数：CD-ROM(4)
 */
Drive.DRIVE_TYPE_CDROM = 4;

/**
 * ドライブタイプの定数：RAMディスク(5)
 */
Drive.DRIVE_TYPE_RAM_DISK = 5;

/**
 * ドライブの一覧を取得します。
 * @return {Array} Driveオブジェクトの配列
 */
Drive.getDrives = function() {
  var result = [];
  var e = new Enumerator(Const.FSO.Drives);
  for (; !e.atEnd(); e.moveNext())
  {
    result.push(new Drive(e.item().Path));
  }

  return result;
};

/**
 * pathで与えられたドライブが存在するかどうかをチェックします。
 * @param {String} path ドライブのパスを文字列で指定します。
 * @return {Boolean} 存在する場合はtrue、存在しない場合はfalseを返します。
 */
Drive.exist = function(path) {
  return Const.FSO.DriveExists(path);
};

// Prototypes of Drive
Drive.prototype = {
  /**
   * このドライブのAvailableSpaceを取得します。
   * @return {Number} このドライブのAvailableSpace
   */
  getAvailableSpace: function() {
    return this.driveObj.AvailableSpace;
  },
  /**
   * このドライブのDriveLetterを取得します。
   * @return {String} このドライブのDriveLetter
   */
  getDriveLetter: function() {
    return this.driveObj.DriveLetter;
  },
  /**
   * このドライブのDriveTypeを取得します。
   * @return {Number} このドライブのDriveType
   */
  getDriveType: function() {
    return this.driveObj.DriveType;
  },
  /**
   * このドライブのFreeSpaceを取得します。
   * @return {Number} このドライブのFreeSpace
   */
  getFreeSpace: function() {
    return this.driveObj.FreeSpace;
  },
  /**
   * このドライブのFileSystemを取得します。
   * @return {String} このドライブのFileSystem
   */
  getFileSystem: function() {
    return this.driveObj.FileSystem;
  },
  /**
   * このドライブが使用可能かどうかを取得します。
   * @return {String} このドライブが使用可能な場合true、そうでない場合はfalseを返します。
   */
  isReady: function() {
    return this.driveObj.IsReady;
  },
  /**
   * このドライブのSerialNumberを取得します。
   * @return {String} このドライブのSerialNumber
   */
  getSerialNumber: function() {
    return this.driveObj.IsReady;
  },
  /**
   * このドライブのShareNameを取得します。
   * @return {String} このドライブのShareName
   */
  getShareName: function() {
    return this.driveObj.ShareName;
  },
  /**
   * このドライブのTotalSizeを取得します。
   * @return {String} このドライブのTotalSize
   */
  getTotalSize: function() {
    return this.driveObj.TotalSize;
  },
  /**
   * このドライブのVolumeNameを取得します。
   * @return {String} このドライブのVolumeName
   */
  getVolumeName: function() {
    return this.driveObj.VolumeName;
  }
};
