# Excelで佐々木希を描く with C# 

Qiitaのとある記事に触発されて作りました｡(丸パクリです)  
[Excelで佐々木希を描く with Node.js](http://qiita.com/Algebra_nobu/items/33781129460eb0338b1b "Title")  
~~Excel操作は[ClosedXML](https://closedxml.codeplex.com "Title")を使いました｡  
完了までにかなり時間がかかります｡(SaveAsに時間がかかってるっぽい)  
倍率は40%くらいにすると程よい感じになります｡  
Bitmapで取得すると何故か反転した状態で取得されるので､苦肉の策として､元画像を反転､回転させました｡~~

ClosedXMLでは保存が遅すぎる､ズームレベルが指定できないなど､不満があったため､[EPPlus](http://epplus.codeplex.com/ "Title")を使用しました｡  
Bitmapで取得すると何故か反転した状態で取得されるので､RotateFlipで調整しました｡  

画像をアップするのはどうかと思ったので削除しました｡
