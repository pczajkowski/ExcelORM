# ExcelORM

They say that if you want have something done right better do it yourself. This is usually not true for software development, but recently I’ve encountered a peculiar issue with my favorite Excel ORM library ([Npoi.Mapper](https://github.com/donnytian/Npoi.Mapper)). For all the users, apart from myself, it was reading only first 170 rows. Why 170, and why only for them, I don’t know. But as I work with Excel files at work constantly and always wanted to play with such thing I’ve decided to create my own, very crude, Excel ORM.

It’s only doing simple reading and writing to and from C# objects. But it’s enough for me. I’ve not developed it from scratch, it’s using [ClosedXML](https://github.com/ClosedXML/ClosedXML) to manipulate Excel files. So, I’ve used the same type of license to honor it.

Check test cases to get an idea of how you should use this library. I may add some examples here later on.

There’s one limitation in ClosedXML where it produces Excel files incompatible either with [Aspose.Cells](https://products.aspose.com/cells/) or with a tool I’m using which uses Aspose.Cells for reading Excel files. Hence, I’ve added IExcelConverter to allow me to workaround it in my tools. You’ll probably not have to use it though.

It currently supports properties of types as supported by ClosedXML, so:
- bool
- double
- string
- DateTime
- TimeSpan

And their nullable variants.

As always, feel free to use it however you desire. But I provide you with no guarantee whatsoever. Enjoy!

## Version history
In version 2 I've added ability to read data dynamically without a need to create a special type. Useful when you need to read/write some not so organized data.

In version 2.2 I've added support for formulas, but as this library is based on ClosedXML it has its limitations, as per [Evaluating Formulas](https://github.com/closedxml/closedxml/wiki/Evaluating-Formulas):
*Not all formulas are included and you'll probably get a nasty error if the formula isn't supported or if there's an error in the formula. Please test your formulas before going to production.*

In version 2.3 I've added support for hyperlinks and improved appending to the existing files.

In version 2.5 I've added Location to ArgumentException Message in ExcelReader. It'll show address of affected cell and the worksheet's name.
