package excel;

import org.apache.commons.collections4.list.GrowthList;
import org.apache.commons.collections4.list.TreeList;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class MyExcelP {
    public static void main(String[] args) throws IOException {
        getAllCell("D:\\Desktop.xlsx");

    }


    public static void getAllCell(String path) throws IOException {
        ArrayList<String> list = new ArrayList<>();
        Map<String, Integer> map = new TreeMap<>();
        Workbook wb = new XSSFWorkbook(new FileInputStream(path));
        int count = 0;
        int count2 = 0;

        Sheet sheet = wb.getSheetAt(0);

        for (Row row : sheet) {
            if (row.getCell(1) != null) {
                count++;
                list.add(row.getCell(1).getStringCellValue());
            }
        }


        for (int i = 0; i < list.size() - 1; i++) {
            count2++;
            if (list.get(i).equals(list.get(i + 1))) {

            } else if (!list.get(i).equals(list.get(i + 1))) {
                map.put(list.get(i), count2);
                count2 = 0;
            }
        }
        map.put(list.get(list.size() - 1), count2 + 1);

        for (Map.Entry<String, Integer> p: map.entrySet()) {
            System.out.println(p.getKey()+" = "+p.getValue());
        }
//        int adv = count/map.size();
//        System.out.println(adv);


//        System.out.println("=====================");
//        System.out.println(count);
//        System.out.println(count2);
        wb.close();

    }


}
