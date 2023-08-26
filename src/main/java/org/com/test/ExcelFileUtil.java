package org.com.test;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ExcelFileUtil {

    public static void writeData(Contact contact) {
        try
        {
            File file = new File("/Users/senthilkumar/file/" + "Contact_Form_Details.xlsx");
            XSSFWorkbook workbook = null;
            XSSFSheet sheet = null;
            int rownum = 0;
            if(!file.exists()){
                workbook = new XSSFWorkbook();
                //Create a blank sheet
                sheet = workbook.createSheet("Contact Form Data");
                rownum = 0;
                Row header = sheet.createRow(rownum);
                header.createCell(0).setCellValue("Name");
                header.createCell(1).setCellValue("Email");
                header.createCell(2).setCellValue("Message");
            } else {
                FileInputStream inputStream = new FileInputStream(file);
                workbook = (XSSFWorkbook) WorkbookFactory.create(inputStream);
                sheet = workbook.getSheetAt(0);
                rownum = sheet.getLastRowNum();
            }
            rownum = rownum + 1;
            //new row
            Row row = sheet.createRow(rownum);
            int cellnum = 0;
            Cell cell = row.createCell(cellnum++);
            cell.setCellValue(contact.getName());

            cell = row.createCell(cellnum++);
            cell.setCellValue(contact.getEmail());

            cell = row.createCell(cellnum++);
            cell.setCellValue(contact.getMessage());

            FileOutputStream out = new FileOutputStream(file);
            workbook.write(out);
            out.close();
            System.out.println("Contact Form Details written successfully on disk.");
            System.out.println();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}
