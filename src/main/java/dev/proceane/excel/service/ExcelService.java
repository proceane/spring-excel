package dev.proceane.excel.service;

import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;

import org.springframework.stereotype.Service;

import dev.proceane.excel.poi317.service.Excel317Service;
import dev.proceane.excel.poi523.service.Excel523Service;
import lombok.RequiredArgsConstructor;

@Service
@RequiredArgsConstructor
public class ExcelService {

    private final Excel317Service service317;
    private final Excel523Service service523;

    public InputStream getExcelFile() throws IOException {
        return service317.getExcelFile();
    }

    public InputStream getEmojiExcelFile() throws IOException {
        return service523.getEmojiExcelFile();
    }

    public InputStream getXssfEncryptExcelFile() throws IOException, GeneralSecurityException {
        return Excel317Service.getXssfEncryptExcelFile();
    }

    public InputStream getSXSSFEncryptExcelFile() throws IOException, GeneralSecurityException {
        return Excel317Service.getSXSSFEncryptExcelFile();
    }

}
