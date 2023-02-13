package dev.proceane.excel.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.security.GeneralSecurityException;

import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
public class ExcelService {

    public InputStream getExcelFile() throws IOException {
        SXSSFWorkbook workbook = getSXSSFWorkbook();
        return getInputStream(workbook);
    }

    public static InputStream getXssfEncryptExcelFile() throws IOException, GeneralSecurityException {
        XSSFWorkbook workbook = getXSSFWorkbook();
        return getEncryptInputStream(workbook);
    }

    public static InputStream getSXSSFEncryptExcelFile() throws IOException, GeneralSecurityException {
        SXSSFWorkbook workbook = getSXSSFWorkbook();
        return getEncryptInputStream(workbook);
    }

    private static XSSFWorkbook getXSSFWorkbook() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("test");
        return workbook;
    }

    private static SXSSFWorkbook getSXSSFWorkbook() {
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("test");
        return workbook;
    }

    private static InputStream getInputStream(Workbook workbook) throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        workbook.close();
        return new ByteArrayInputStream(outputStream.toByteArray());
    }


    private static InputStream getEncryptInputStream(Workbook workbook) throws IOException, GeneralSecurityException {
        EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
        Encryptor enc = info.getEncryptor();
        enc.confirmPassword("1234");

        POIFSFileSystem fs = new POIFSFileSystem();
        OutputStream os = enc.getDataStream(fs);
        workbook.write(os);

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        fs.writeFilesystem(outputStream);

        workbook.close();
        os.close();
        fs.close();
        return new ByteArrayInputStream(outputStream.toByteArray());
    }
}
