package dev.proceane.excel.poi523.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
public class Excel523Service {

    public InputStream getEmojiExcelFile() throws IOException {
        Workbook workbook = getEmojiSXSSFWorkbook();
        return getInputStream(workbook);
    }

    private static Workbook getEmojiSXSSFWorkbook() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("emoji");
        cell = row.createCell(1);
        cell.setCellValue("😊😊😊💙💙💙");

        row = sheet.createRow(1);
        cell = row.createCell(0);
        cell.setCellValue("korean");
        cell = row.createCell(1);
        cell.setCellValue("안녕하세요");

        row = sheet.createRow(2);
        cell = row.createCell(0);
        cell.setCellValue("korean+emoji");
        cell = row.createCell(1);
        cell.setCellValue("안녕하세요💙💙💙");
        return workbook;
    }

    private static ByteArrayInputStream getInputStream(Workbook workbook) throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        workbook.close();
        return new ByteArrayInputStream(outputStream.toByteArray());
    }
}
