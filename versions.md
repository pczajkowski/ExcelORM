# Version history

| Version | Description |
| ----------- | ----------- |
| 2.0.0 | I've added ability to read data dynamically without a need to create a special type. Useful when you need to read/write some not so organized data.|
| 2.2.0 | I've added support for formulas, but as this library is based on ClosedXML it has its limitations, as per [Evaluating Formulas](https://github.com/closedxml/closedxml/wiki/Evaluating-Formulas):
*Not all formulas are included and you'll probably get a nasty error if the formula isn't supported or if there's an error in the formula. Please test your formulas before going to production.*|
| 2.3.0 | I've added support for hyperlinks and improved appending to the existing files.|

| 2.5.0 | I've added Location to ArgumentException Message in ExcelReader. It'll show address of affected cell and the worksheet's name.|
