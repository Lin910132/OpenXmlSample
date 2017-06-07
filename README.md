# OpenXmlSample
Sample code for utilizing OpenXml library to edit/create excel files with(out) template

There are two versions.
One has DataReader, the other doesn't. The dataReader is used to avoid keeping to much data in the memory.

There are two Approaches. One is DOM, the other is SAX.
DOM approach is type safe, you can have a acurate edit on the sheet, row and cell. However, it might cause memory issues if working with too much data.
SAX approach is best for adding tons of data and creating new Excel File. However, it's working directly with XML for the Excel file,
which means it's not type safe, and has strict format rules.

I've implemented both approaches, mostly in SAX, to make manipulating Excel files much easier.
