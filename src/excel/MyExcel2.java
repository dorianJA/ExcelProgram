package excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class MyExcel2 {
    public static void main(String[] args) throws IOException {
        getFfiles("D:\\Desktop.xlsx","D:\\new.xlsx");
    }

    public static void getFfiles(String readFile,String WriteFile) throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream(WriteFile);
        Workbook wb = new XSSFWorkbook(new FileInputStream(readFile));
        int count = 0;

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();

        for(Row r: wb.getSheetAt(0)) {
            Row row = sheet.createRow(count);
            if(r.getCell(1)!=null)
            row.createCell(0).setCellValue(r.getCell(1).getStringCellValue()); // записываем в ячейку нового файла,
                                                                                                                    // данные первой строки и второго столбца считывающего файла
            count++;
        }

        workbook.write(fileOutputStream);
        fileOutputStream.close();

    }
}
