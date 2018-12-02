## Real Streaming Version of Apache POI SXSSFCell, SXSSFRow, SXSSFSheet, SXSSFWorkbook

Streaming classes from poi-ooxml v3.9 extended to really stream data as soon as a row has been created.
The usage of temporary data files removed. The Worksheet XML rows are provided as Java 8 Stream<String>.
The XML stream supposed to be mapped out of data Stream with help of StreamSheet.flushRow(): String,
see sample:

```java
        Stream<Integer> stream = Stream.of(0, 1, 2, 3);
        StreamWorkbook wb = new StreamWorkbook();
        wb.createSheet(sheetName);
        wb.setStreamSource(st -> stream.map(i -> {
            try {
                st.createRow(i).createCell(0).setCellValue(prefix + i + "0");
                return st.flushRows();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }));
        wb.write(outputStream);

```