oo4o Buster
=====

oo4o�𑀍삷��R�[�h���Ȃ�ׂ��ύX�����ɁAODP.NET�֒u�������邽�߂̃T�|�[�g���s���܂��B

## �K�{�v��

- ODP.NET 11g or higher

## �g����

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

## ���C�Z���X

[zlib/libpng](http://opensource.org/licenses/Zlib)
