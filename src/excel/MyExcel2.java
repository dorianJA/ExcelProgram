package excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class MyExcel2 {
    public static void main(String[] args) throws IOException {
        addNewFile("D:\\Desktop.xlsx","D:\\new.xlsx");
    }

    public static void addNewFile(String readFile,String WriteFile) throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream(WriteFile);
        Workbook wb = new XSSFWorkbook(new FileInputStream(readFile));
        int count = 0;
        double numberCount = 0;
        List<String> list = new ArrayList<>();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        Row row = sheet.createRow(count);
        for(Row r: wb.getSheetAt(0)) {
            //Row row = sheet.createRow(count);
            if(r.getCell(1)!=null)
            //row.createCell(0).setCellValue(r.getCell(1).getStringCellValue()); // записываем в ячейку нового файла,
                                                                                                                    // данные первой строки и второго столбца считывающего файла
            list.add(r.getCell(1).getStringCellValue());
            //count++;
        }



        for(int i = 0; i < list.size()-1;i++){
            numberCount++;
            if (list.get(i)!=null && !list.get(i).equals(list.get(i+1))){
                row = sheet.createRow(count);
                count++;
                row.createCell(0).setCellValue(list.get(i));
                row.createCell(1).setCellValue(numberCount);
                numberCount = 0;

            }
        }
        row = sheet.createRow(count);
        row.createCell(0).setCellValue(list.get(list.size()-1));
        row.createCell(1).setCellValue(numberCount+1);




//        for(int i = 0; i < wb.getSheetAt(0).getLastRowNum()-1; i++){
//            if( wb.getSheetAt(0).getRow(i).getCell(1)!=null && !(wb.getSheetAt(0).getRow(i).getCell(1).getStringCellValue().equals(wb.getSheetAt(0).getRow(i+1).getCell(1).getStringCellValue()))){
//                Row row = sheet.createRow(count);
//                count++;
//                row.createCell(0).setCellValue(wb.getSheetAt(0).getRow(i).getCell(1).getStringCellValue());
//            }
//        }
        workbook.write(fileOutputStream);
        fileOutputStream.close();

    }
}
