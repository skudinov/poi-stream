package excel.streaming;

import org.apache.poi.ss.usermodel.Cell;
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
        StreamWorkbook streamWb = createWorkbook("S1", "value", Stream.of(0, 1, 2));
        XSSFWorkbook xssfWb = (XSSFWorkbook) writeOutAndReadBack(streamWb);
        XSSFSheet st = xssfWb.getSheet("S1");
        for (int i = 0; i < 2; i++) {
            assertEquals("value" + i + "0", st.getRow(i).getCell(0).getStringCellValue());
        }
    }

    private StreamWorkbook createWorkbook(String sheetName, String prefix, Stream<Integer> stream) {
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
        return wb;
    }

    private Workbook writeOutAndReadBack(StreamWorkbook wb) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream(8192);
        wb.write(baos);
        return new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()));
    }

}
