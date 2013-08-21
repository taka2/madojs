// 固定長ファイルからExcelファイルを作成する。
// 0-1byte, 1-4byte, 4-5byte
Excel.convertFromFixed(ARGV[0], [[0, Excel.COLUMN_TYPE_TEXT], [1, Excel.COLUMN_TYPE_TEXT], [4, Excel.COLUMN_TYPE_TEXT]]);

print("done.");
