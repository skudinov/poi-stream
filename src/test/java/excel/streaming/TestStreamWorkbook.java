package excel.streaming;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.*;
import java.util.stream.Stream;

import static org.junit.Assert.assertEquals;

public class TestStreamWorkbook {
    @Test
    public void workbookStreamWrite() throws IOException {
        StreamWorkbook streamWb = createWorkbook("S1", "value", Stream.of(1, 2, 3));
        XSSFWorkbook xssfWb = (XSSFWorkbook) writeOutAndReadBack(streamWb);
        XSSFSheet st = xssfWb.getSheet("S1");
        assertEquals(3, st.getLastRowNum());
        for (int i = 0; i < 4; i++) {
            assertEquals("value" + i + "0", st.getRow(i).getCell(0).getStringCellValue());
        }
    }

    private StreamWorkbook createWorkbook(String sheetName, String prefix, Stream<Integer> stream) {
        StreamWorkbook wb = new StreamWorkbook();
        Sheet sheet = wb.createSheet(sheetName);
        sheet.createRow(0).createCell(0).setCellValue(prefix + 0 + "0");
        wb.setStreamSource(st -> Stream.concat(Stream.of(st.flushRowsUnchecked()), stream.map(i -> {
            st.createRow(i).createCell(0).setCellValue(prefix + i + "0");
            return st.flushRowsUnchecked();
        })));
        return wb;
    }

    private Workbook writeOutAndReadBack(StreamWorkbook wb) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream(8192);
        wb.write(baos);
        return new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()));
    }

}
