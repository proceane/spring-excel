package dev.proceane.excel.service;

import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;

import org.springframework.stereotype.Service;

import dev.proceane.excel.poi316.service.Excel316Service;
import dev.proceane.excel.poi523.service.Excel523Service;
import lombok.RequiredArgsConstructor;

@Service
@RequiredArgsConstructor
public class ExcelService {

    private final Excel316Service service316;
    private final Excel523Service service523;

    public InputStream getExcelFile() throws IOException {
        return service316.getExcelFile();
    }

    public InputStream getEmojiExcelFile() throws IOException {
        return service523.getEmojiExcelFile();
    }

    public InputStream getXssfEncryptExcelFile() throws IOException, GeneralSecurityException {
        return Excel316Service.getXssfEncryptExcelFile();
    }

    public InputStream getSXSSFEncryptExcelFile() throws IOException, GeneralSecurityException {
        return Excel316Service.getSXSSFEncryptExcelFile();
    }

}
