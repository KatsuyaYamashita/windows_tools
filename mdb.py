# -*- coding: utf-8 -*-

from win32com import client as w32com
import sys

class Mdb:
    """Microsoft ACCESS(mdb)用モジュール
       本モジュールはMicrosoft Jet Driverを使用している。

       Office 2010用のJetドライバは以下で配布しているので適宜Jet Driverをインストールすること。
       http://www.microsoft.com/en-us/download/details.aspx?id=13255

       また、Officeのバイナリが32bit版か64bit版かにより実行するPythonのバージョンを
       32bitないし64bitに合わせる必要がある。
    """

    def __init__(self, mdb_name):
        """初期化関数
        Args:
          mdb_name: 扱うmdbファイル名(パス付)
        """

#        __CONNECTION_STRING = "Driver={Microsoft Access Driver (*.mdb)};DBQ=%s;"
        __CONNECTION_STRING = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=%s;"

        # OLEDB.4.0はAccess2007以降非推奨となった
        # http://msdn.microsoft.com/ja-jp/library/office/ff965871%28v=office.14%29.aspx#DataProgrammingWithAccess2010_using32vs64ace
#        __CONNECTION_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s;"
        self._connection_string = __CONNECTION_STRING % (mdb_name)
        print(self._connection_string)

    def open(self):
        """mdbをオープンする"""

        self._connection = w32com.Dispatch("ADODB.Connection")
        self._connection.Open(self._connection_string)
        self._catalog = w32com.Dispatch("ADOX.Catalog")
        self._catalog.ActiveConnection = self._connection

    def tables(self):
        """mdb内のテーブルをすべて取得してリストにして返す
        Return:
          mdb内のテーブルのリスト
        """

        ts = []
        for table in self._catalog.Tables:
            ts.append(table.Name)
        
        return ts

    def fields(self, table):
        """指定したテーブル内のフィールド名をリストにして返す
        Args:
          table: フィールドを取得したいテーブル名
        Return:
          フィールドのリスト
        """

        rs = w32com.Dispatch("ADODB.RecordSet")

        t = []
        rs.Open(table, self._connection)
        for f in rs.Fields:
            t.append(f.Name)
        rs.Close()

        return t

    def query(self, sql):
        """オープンしたmdbに対してqueryを投げる
        Args:
          sql: query(SQL)
        Return:
          取得したレコード(二重リスト)。以下のように取得される。
          [
          [field1, field2, field3], # レコード1
          [field1, field2, field3], # レコード2
          [field1, field2, field3]  # レコード3
          ] #レコードリスト
        """
        rs = w32com.Dispatch("ADODB.RecordSet")

        records = []
        rs.Open(sql, self._connection)
        while not rs.EOF:
            r = []
            for f in rs.Fields:
                r.append(f.Value)
            rs.MoveNext()
            records.append(r)
        rs.Close()

        return records

    def close(self):
        """mdbをクローズする"""
        self._connection.Close()

"""
テスト用メインルーチン
    mdb.py mdb_file table_name

Args:
    mdb_file: テストするmdbファイル名(パス付)
    table_name: テストするテーブル名

"""
if __name__ == '__main__':
    for n in sys.argv:
        print(n)
    
    mdb = Mdb(sys.argv[1])
    mdb.open()
    ts = mdb.tables()

    for t in ts:
        print("table: %s" % t)

    fs = mdb.fields(sys.argv[2])
    for f in fs:
        print("fields: %s" % f)

    rs = mdb.query("SELECT * FROM %s" % sys.argv[2])
    print("Record Count: %d" % len(rs))

    for r in rs:
        for f in r:
            print(f, end=",")
        print()

    mdb.close()
