package com.example.jxlstesting3.controller;

import com.example.jxlstesting3.service.EmployeeService;
import lombok.extern.slf4j.Slf4j;
import org.jxls.builder.JxlsOutputFile;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Map;
@RestController
@Slf4j
public class EmployeeController {

    @Value("classpath:templates/employee-result.xlsx")
    private Resource resource;

    private static final String XLSX_MEDIA_TYPE = "application/vnd.openxmlformats";

    private static final String FILENAME_DATE_FORMAT = "yyyyMMdd'T'HHmmss";
    private final EmployeeService employeeService;

    @Autowired
    public EmployeeController(EmployeeService employeeService) {
        this.employeeService = employeeService;
    }

    @GetMapping(value = "/xls/employees", produces = XLSX_MEDIA_TYPE)
    public ResponseEntity<Resource> getAllEmployees() {

        Map<String, Object> data = new HashMap<>();
        data.put("employees", employeeService.getAllEmployees());


        File file = getFileTemp();
        JxlsOutputFile jxlsOutputFile = new JxlsOutputFile(file);

        xlsxBuilder(data,jxlsOutputFile);

        return getResponseEntity("employee-result", file);
    }

    private void xlsxBuilder(Map<String, Object> data, JxlsOutputFile jxlsOutputFile) {
        try {
            JxlsPoiTemplateFillerBuilder.newInstance()
                    .withTemplate(resource.getFile())
                    .build()
                    .fill(data, jxlsOutputFile);
        } catch (IOException e) {
            log.error(e.getMessage());
            throw new RuntimeException(e.getMessage());
        }
    }

    private File getFileTemp()  {
        try {
            File tempFile = File.createTempFile("tempfile", ".xlsx");
            tempFile.deleteOnExit();
            return tempFile;
        } catch (IOException e) {
            log.error(e.getMessage());
            throw new RuntimeException(e.getMessage());
        }
    }

    private ResponseEntity<Resource> getResponseEntity(String nomeFile, File file) {
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\""+buildNomeFile(nomeFile)+"\"")
                .contentLength(file.length())
                .contentType(MediaType.valueOf(XLSX_MEDIA_TYPE))
                .body(new FileSystemResource(file));
    }

    private String buildNomeFile(String nomeFile) {
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern(FILENAME_DATE_FORMAT);
        return String.format("%s_%s.xlsx",nomeFile, LocalDateTime.now().format(dtf));
    }
}
