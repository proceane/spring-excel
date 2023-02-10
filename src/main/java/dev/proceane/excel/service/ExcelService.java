package dev.proceane.excel.service;

import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.*;
import java.security.GeneralSecurityException;

@Service
public class ExcelService {

    public InputStream getExcelFile() throws IOException {
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        SXSSFSheet sheet = workbook.createSheet();
        SXSSFRow row = sheet.createRow(0);
        SXSSFCell cell = row.createCell(0);
        cell.setCellValue("test");

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        workbook.close();
        return new ByteArrayInputStream(outputStream.toByteArray());
    }

    public InputStream getXssfEncryptExcelFile() throws IOException, GeneralSecurityException {
        XSSFWorkbook workbook = getXSSFWorkbook();

        EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
        Encryptor enc = info.getEncryptor();
        enc.confirmPassword("1234");

        POIFSFileSystem fs = new POIFSFileSystem();
        OutputStream os = enc.getDataStream(fs);
        workbook.write(os);

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        fs.writeFilesystem(outputStream);

        os.close();
        workbook.close();
        fs.close();

        return new ByteArrayInputStream(outputStream.toByteArray());
    }

    private XSSFWorkbook getXSSFWorkbook() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("test");
        return workbook;
    }
}
