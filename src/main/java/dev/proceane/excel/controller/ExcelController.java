package dev.proceane.excel.controller;

import java.io.IOException;
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

@RestController
@RequiredArgsConstructor
public class ExcelController {

    private final ExcelService service;

    @GetMapping("/excel")
    public ResponseEntity<Resource> downloadExcelFile() throws IOException {
        InputStreamResource resource = new InputStreamResource(service.getExcelFile());
        return ResponseEntity.ok()
                             .header(HttpHeaders.CONTENT_DISPOSITION, ContentDisposition.attachment()
                                     .filename("excel.xlsx", StandardCharsets.UTF_8)
                                     .build().toString())
                             .contentType(MediaType.APPLICATION_OCTET_STREAM)
                             .body(resource);
    }

}
