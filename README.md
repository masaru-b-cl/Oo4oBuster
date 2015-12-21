oo4o Buster
=====

oo4oを操作するコードをなるべく変更せずに、ODP.NETへ置き換えるためのサポートを行います。

## 必須要件

- ODP.NET 11g or higher

## 使い方

```vb
Using db = BusterDb.OpenDatabase(dbName, userId, password)
  Dim dynaset = db.CreateDynaset(sql)

  Console.WriteLine("count = {0}", dynaset.RecordCount)

  While Not dynaset.EOF
    Console.WriteLine("name = {0}", dynaset("name").Value)
    Console.WriteLine("age = {0}", dynaset("age").Value)

    dynaset.MoveNext()
  Next

  dynaset.Close()
  db.Close()
End Using
```

## ライセンス

[zlib/libpng](http://opensource.org/licenses/Zlib)
