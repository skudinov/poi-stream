## Real Streaming Version of Apache POI SXSSFCell, SXSSFRow, SXSSFSheet, SXSSFWorkbook

Streaming classes from poi-ooxml v3.9 extended to really stream data as soon as a row has been created.
The usage of temporary data files removed. The Worksheet XML rows are provided as Java 8 Stream<String>.
The XML stream supposed to be mapped out of data Stream with help of StreamSheet.flushRow(): String,
see sample:

```java
        Stream<Integer> stream = Stream.of(1, 2, 3);
        StreamWorkbook wb = new StreamWorkbook();
        Sheet sheet = wb.createSheet(sheetName);
        // create header row
        sheet.createRow(0).createCell(0).setCellValue(prefix + 0 + "0");
        // flush header row even if there might not be data rows
        wb.setStreamSource(st -> Stream.concat(Stream.of(st.flushRowsUnchecked()), stream.map(i -> {
            st.createRow(i).createCell(0).setCellValue(prefix + i + "0");
            return st.flushRowsUnchecked();
        })));
        wb.write(outputStream);
```