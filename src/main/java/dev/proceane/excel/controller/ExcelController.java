package dev.proceane.excel.controller;

import java.nio.charset.StandardCharsets;

import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import dev.proceane.excel.service.ExcelService;
import lombok.RequiredArgsConstructor;
import lombok.SneakyThrows;

@RestController
@RequiredArgsConstructor
public class ExcelController {

    private final ExcelService service;

    @SneakyThrows
    @GetMapping("/excel")
    public ResponseEntity<Resource> downloadExcelFile() {
        InputStreamResource resource = new InputStreamResource(service.getExcelFile());
        return ResponseEntity.ok()
                             .header(HttpHeaders.CONTENT_DISPOSITION, ContentDisposition.attachment()
                                     .filename("excel.xlsx", StandardCharsets.UTF_8)
                                     .build().toString())
                             .contentType(MediaType.APPLICATION_OCTET_STREAM)
                             .body(resource);
    }

    @SneakyThrows
    @GetMapping("/excel/emoji")
    public ResponseEntity<Resource> downloadEmojiExcelFile() {
        InputStreamResource resource = new InputStreamResource(service.getEmojiExcelFile());
        return ResponseEntity.ok()
                             .header(HttpHeaders.CONTENT_DISPOSITION, ContentDisposition.attachment()
                                                                                        .filename("excel_emoji.xlsx", StandardCharsets.UTF_8)
                                                                                        .build().toString())
                             .contentType(MediaType.APPLICATION_OCTET_STREAM)
                             .body(resource);
    }

    @SneakyThrows
    @GetMapping("/excel/xssf/encrypt")
    public ResponseEntity<Resource> downloadXssfEncryptExcelFile() {
        InputStreamResource resource = new InputStreamResource(service.getXssfEncryptExcelFile());
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, ContentDisposition.attachment()
                        .filename("excel_encrypt.xlsx", StandardCharsets.UTF_8)
                        .build().toString())
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(resource);
    }

    @SneakyThrows
    @GetMapping("/excel/sxssf/encrypt")
    public ResponseEntity<Resource> downloadSXSSFEncryptExcelFile() {
        InputStreamResource resource = new InputStreamResource(service.getSXSSFEncryptExcelFile());
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, ContentDisposition.attachment()
                        .filename("excel_encrypt.xlsx", StandardCharsets.UTF_8)
                        .build().toString())
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(resource);
    }

}
