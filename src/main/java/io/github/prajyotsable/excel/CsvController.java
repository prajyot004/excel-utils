package io.github.prajyotsable.excel;

import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.util.List;

/**
 * REST endpoint that streams a CSV file directly to the browser.
 *
 * GET /api/users/download/csv
 *
 * Uses StreamingResponseBody so rows are piped straight from the generator
 * into the HTTP response socket — zero intermediate byte[] buffer in heap.
 * Safe for any number of rows.
 */
@RestController
@RequestMapping("/api/users")
public class CsvController {

    private final CsvGenerationUtil   csvUtil   = new CsvGenerationUtil();
    private final ExcelGenerationUtil excelUtil = new ExcelGenerationUtil();

    /**
     * Streams the CSV file to the frontend.
     *
     * The browser receives Content-Disposition: attachment, so it triggers
     * a file-save dialog / automatic download automatically.
     *
     * Endpoint : GET http://localhost:8080/api/users/download/csv
     */
    @GetMapping("/download/csv")
    public ResponseEntity<StreamingResponseBody> downloadCsv() {

        // Replace this with your real service / repository call
        List<UserRecord> records = TestDataGenerator.generateUsers(1_000_000);

        // StreamingResponseBody is a functional interface:
        //   void writeTo(OutputStream outputStream) throws IOException
        // Spring Boot runs it on an async thread and flushes chunks to the
        // client as they are written — the web thread is never blocked.
        StreamingResponseBody responseBody = outputStream ->
                csvUtil.streamCsvToResponse(records, UserRecord.class, outputStream);

        return ResponseEntity.ok()
                // tells the browser: save this as a file named users.csv
                .header(HttpHeaders.CONTENT_DISPOSITION,
                        "attachment; filename=\"users.csv\"")
                // no Content-Length — size is unknown until fully written
                .contentType(MediaType.parseMediaType("text/csv; charset=UTF-8"))
                .body(responseBody);
    }

    /**
     * Streams the Excel (.xlsx) file to the frontend.
     *
     * Endpoint : GET http://localhost:8080/api/users/download/xlsx
     */
    @GetMapping("/download/xlsx")
    public ResponseEntity<StreamingResponseBody> downloadExcel() {

        // Replace this with your real service / repository call
        List<UserRecord> records = TestDataGenerator.generateUsers(1_000_000);

        StreamingResponseBody responseBody = outputStream ->
                excelUtil.streamExcelToResponse(records, UserRecord.class, outputStream);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION,
                        "attachment; filename=\"users.xlsx\"")
                .contentType(MediaType.parseMediaType(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(responseBody);
    }
}
