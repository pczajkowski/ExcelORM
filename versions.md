# Version history

| Version | Description |
| ----------- | ----------- |
| 2.0.0 | Added ability to read data dynamically without a need to create a special type. Useful when you need to read/write some not so organized data.|
| 2.2.0 | Added support for formulas, but as this library is based on ClosedXML it has its limitations, as per [Evaluating Formulas](https://github.com/closedxml/closedxml/wiki/Evaluating-Formulas): *Not all formulas are included and you'll probably get a nasty error if the formula isn't supported or if there's an error in the formula. Please test your formulas before going to production.*|
| 2.3.0 | Added support for hyperlinks and improved appending to the existing files.|
| 2.5.0 | Added Location to ArgumentException Message in ExcelReader. It'll show address of affected cell and the worksheet's name.|
| 2.6.0 | Added support for appending starting from given row.|
| 2.7.0 | Added support for reading properties of type Guid and enum.|
| 2.8.0 | Handling more number types. Properly handling appending to and reading from empty file.|
| 2.8.1 | Ability to start writing from given row.|
