package io.github.prajyotsable.excel;

import org.apache.poi.util.RecordFormatException;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.server.ResponseStatusException;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.IOException;
import java.util.List;
import java.util.Map;

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

        private final CsvGenerationUtil csvUtil = new CsvGenerationUtil();
        private final ExcelGenerationUtil excelUtil = new ExcelGenerationUtil();
        private final ExcelToJsonStreamingUtil excelToJsonUtil = new ExcelToJsonStreamingUtil();

        /**
         * Accepts an uploaded Excel file (.xlsx) and streams JSON response.
         * Designed for very large files (100K+ rows) using SAX/event parsing,
         * so rows are never fully loaded into memory.
         *
         * Endpoint: POST /api/users/upload/xlsx/to-json
         * Form-data: file=<excel-file>
         * Query params:
         * - mode=first|all|name|index (default: first)
         * - sheetName=<sheet-name> (required when mode=name)
         * - sheetIndex=<0-based-index> (required when mode=index)
         */
        @PostMapping(value = "/upload/xlsx/to-json", consumes = MediaType.MULTIPART_FORM_DATA_VALUE, produces = MediaType.APPLICATION_JSON_VALUE)
        public ResponseEntity<StreamingResponseBody> uploadExcelToJson(
                        @RequestParam("file") MultipartFile file,
                        @RequestParam(value = "mode", defaultValue = "first") String mode,
                        @RequestParam(value = "sheetName", required = false) String sheetName,
                        @RequestParam(value = "sheetIndex", required = false) Integer sheetIndex) {

                if (file == null || file.isEmpty()) {
                        throw new ResponseStatusException(HttpStatus.BAD_REQUEST, "Uploaded file is empty");
                }

                String originalFileName = file.getOriginalFilename();
                if (originalFileName == null ||
                                !originalFileName.toLowerCase(java.util.Locale.ROOT).endsWith(".xlsx")) {
                        throw new ResponseStatusException(HttpStatus.BAD_REQUEST, "Only .xlsx files are supported");
                }

                StreamingResponseBody responseBody = outputStream -> {
                        try (java.io.InputStream inputStream = file.getInputStream()) {
                                String normalizedMode = mode == null ? "first"
                                                : mode.trim().toLowerCase(java.util.Locale.ROOT);
                                switch (normalizedMode) {
                                        case "all":
                                                excelToJsonUtil.streamAllSheetsAsJson(inputStream, outputStream);
                                                break;
                                        case "name":
                                                if (sheetName == null || sheetName.trim().isEmpty()) {
                                                        throw new ResponseStatusException(HttpStatus.BAD_REQUEST,
                                                                        "sheetName is required when mode=name");
                                                }
                                                excelToJsonUtil.streamSheetByNameAsJson(inputStream, outputStream,
                                                                sheetName);
                                                break;
                                        case "index":
                                                if (sheetIndex == null) {
                                                        throw new ResponseStatusException(HttpStatus.BAD_REQUEST,
                                                                        "sheetIndex is required when mode=index");
                                                }
                                                if (sheetIndex < 0) {
                                                        throw new ResponseStatusException(HttpStatus.BAD_REQUEST,
                                                                        "sheetIndex must be >= 0");
                                                }
                                                excelToJsonUtil.streamSheetByIndexAsJson(inputStream, outputStream,
                                                                sheetIndex);
                                                break;
                                        case "first":
                                                excelToJsonUtil.streamFirstSheetAsJson(inputStream, outputStream);
                                                break;
                                        default:
                                                throw new ResponseStatusException(HttpStatus.BAD_REQUEST,
                                                                "Unsupported mode. Use first|all|name|index");
                                }
                        } catch (ResponseStatusException e) {
                                throw e;
                        } catch (RecordFormatException e) {
                                throw new IOException(
                                                "Failed to convert uploaded Excel to JSON: file has very large internal records. "
                                                                + "Increase JVM system property excel.poi.byteArrayMaxOverride (bytes), "
                                                                + "for example -Dexcel.poi.byteArrayMaxOverride=1073741824",
                                                e);
                        } catch (Exception e) {
                                throw new IOException("Failed to convert uploaded Excel to JSON", e);
                        }
                };

                return ResponseEntity.ok()
                                .contentType(MediaType.APPLICATION_JSON)
                                .body(responseBody);
        }

        /**
         * Accepts an uploaded Excel file (.xlsx) and streams NDJSON events that the
         * frontend can consume incrementally.
         *
         * Endpoint: POST /api/users/upload/xlsx/to-json/events
         * Form-data: file=<excel-file>
         * Query params:
         * - mode=first|all|name|index (default: first)
         * - sheetName=<sheet-name> (required when mode=name)
         * - sheetIndex=<0-based-index> (required when mode=index)
         */
        @PostMapping(value = "/upload/xlsx/to-json/events", consumes = MediaType.MULTIPART_FORM_DATA_VALUE, produces = "application/x-ndjson")
        public ResponseEntity<StreamingResponseBody> uploadExcelToJsonEvents(
                        @RequestParam("file") MultipartFile file,
                        @RequestParam(value = "mode", defaultValue = "first") String mode,
                        @RequestParam(value = "sheetName", required = false) String sheetName,
                        @RequestParam(value = "sheetIndex", required = false) Integer sheetIndex) {

                validateUploadedExcel(file);

                StreamingResponseBody responseBody = outputStream -> {
                        try (java.io.InputStream inputStream = file.getInputStream()) {
                                String normalizedMode = mode == null ? "first"
                                                : mode.trim().toLowerCase(java.util.Locale.ROOT);
                                switch (normalizedMode) {
                                        case "all":
                                                excelToJsonUtil.streamAllSheetsAsNdjson(inputStream, outputStream);
                                                break;
                                        case "name":
                                                if (sheetName == null || sheetName.trim().isEmpty()) {
                                                        throw new ResponseStatusException(HttpStatus.BAD_REQUEST,
                                                                        "sheetName is required when mode=name");
                                                }
                                                excelToJsonUtil.streamSheetByNameAsNdjson(inputStream, outputStream,
                                                                sheetName);
                                                break;
                                        case "index":
                                                if (sheetIndex == null) {
                                                        throw new ResponseStatusException(HttpStatus.BAD_REQUEST,
                                                                        "sheetIndex is required when mode=index");
                                                }
                                                if (sheetIndex < 0) {
                                                        throw new ResponseStatusException(HttpStatus.BAD_REQUEST,
                                                                        "sheetIndex must be >= 0");
                                                }
                                                excelToJsonUtil.streamSheetByIndexAsNdjson(inputStream, outputStream,
                                                                sheetIndex);
                                                break;
                                        case "first":
                                                excelToJsonUtil.streamFirstSheetAsNdjson(inputStream, outputStream);
                                                break;
                                        default:
                                                throw new ResponseStatusException(HttpStatus.BAD_REQUEST,
                                                                "Unsupported mode. Use first|all|name|index");
                                }
                        } catch (ResponseStatusException e) {
                                throw e;
                        } catch (RecordFormatException e) {
                                throw new IOException(
                                                "Failed to convert uploaded Excel to NDJSON: file has very large internal records. "
                                                                + "Increase JVM system property excel.poi.byteArrayMaxOverride (bytes), "
                                                                + "for example -Dexcel.poi.byteArrayMaxOverride=1073741824",
                                                e);
                        } catch (Exception e) {
                                throw new IOException("Failed to convert uploaded Excel to NDJSON events", e);
                        }
                };

                return ResponseEntity.ok()
                                .contentType(MediaType.parseMediaType("application/x-ndjson"))
                                .body(responseBody);
        }

        private void validateUploadedExcel(MultipartFile file) {
                if (file == null || file.isEmpty()) {
                        throw new ResponseStatusException(HttpStatus.BAD_REQUEST, "Uploaded file is empty");
                }

                String originalFileName = file.getOriginalFilename();
                if (originalFileName == null ||
                                !originalFileName.toLowerCase(java.util.Locale.ROOT).endsWith(".xlsx")) {
                        throw new ResponseStatusException(HttpStatus.BAD_REQUEST, "Only .xlsx files are supported");
                }
        }

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
                // void writeTo(OutputStream outputStream) throws IOException
                // Spring Boot runs it on an async thread and flushes chunks to the
                // client as they are written — the web thread is never blocked.
                StreamingResponseBody responseBody = outputStream -> csvUtil.streamCsvToResponse(records,
                                UserRecord.class, outputStream);

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

                StreamingResponseBody responseBody = outputStream -> excelUtil.streamExcelToResponse(records,
                                UserRecord.class, outputStream);

                return ResponseEntity.ok()
                                .header(HttpHeaders.CONTENT_DISPOSITION,
                                                "attachment; filename=\"users.xlsx\"")
                                .contentType(MediaType.parseMediaType(
                                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                                .body(responseBody);
        }

        // ── Method 2: generateCsvAsBytes ─────────────────────────────────────────
        /**
         * Returns the entire CSV as a byte[] in the response body.
         * Spring sets Content-Length automatically so the browser shows exact %
         * progress.
         * Capped at 100 000 rows — the full byte[] lives in JVM heap.
         *
         * Endpoint: GET /api/users/download/csv/bytes
         */
        @GetMapping("/download/csv/bytes")
        public ResponseEntity<byte[]> downloadCsvAsBytes() throws IOException {
                List<UserRecord> records = TestDataGenerator.generateUsers(1_000_000);
                byte[] bytes = csvUtil.generateCsvAsBytes(records, UserRecord.class);
                return ResponseEntity.ok()
                                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"users.csv\"")
                                .header(HttpHeaders.CONTENT_LENGTH, String.valueOf(bytes.length))
                                .contentType(MediaType.parseMediaType("text/csv; charset=UTF-8"))
                                .body(bytes);
        }

        // ── Method 3: generateCsvAsBase64 ────────────────────────────────────────
        /**
         * Returns a JSON body { fileName, data } where data is Base64-encoded CSV.
         * The browser JS decodes it locally and triggers a Blob download — no extra
         * request.
         * Capped at 100 000 rows — Base64 is ~33% larger than raw bytes.
         *
         * Endpoint: GET /api/users/download/csv/base64
         */
        @GetMapping("/download/csv/base64")
        public ResponseEntity<Map<String, String>> downloadCsvAsBase64() throws IOException {
                List<UserRecord> records = TestDataGenerator.generateUsers(1_000_000);
                String base64 = csvUtil.generateCsvAsBase64(records, UserRecord.class);
                return ResponseEntity.ok(Map.of("fileName", "users.csv", "data", base64));
        }

        // ── Method 1: generateCsvToFile ──────────────────────────────────────────
        /**
         * Saves the CSV directly to the server's local filesystem.
         * Returns JSON with the absolute path where the file was written.
         * Nothing is downloaded to the browser — this is a server-side write.
         *
         * Endpoint: GET /api/users/download/csv/file
         */
        @GetMapping("/download/csv/file")
        public ResponseEntity<Map<String, String>> saveCsvToFile() throws IOException {
                List<UserRecord> records = TestDataGenerator.generateUsers(1_000_000);
                String path = "users_export.csv";
                csvUtil.generateCsvToFile(records, UserRecord.class, path);
                return ResponseEntity.ok(Map.of(
                                "method", "generateCsvToFile",
                                "message", "File saved on server",
                                "path", new java.io.File(path).getAbsolutePath()));
        }

        // ── Custom headers: stream endpoint ──────────────────────────────────────
        /**
         * Demonstrates custom-header CSV generation.
         *
         * <p>
         * Headers passed: {@code id, firstName, email, salary, department, createdDate}
         * <br>
         * UserRecord fields: {@code id, firstName, lastName, email, age, createdDate}
         * <br>
         * → {@code salary} and {@code department} are NOT in the DTO → blank columns.
         * <br>
         * → {@code lastName} and {@code age} are in the DTO but NOT in headers →
         * omitted.
         *
         * Endpoint: GET /api/users/download/csv/custom-headers
         */
        @GetMapping("/download/csv/custom-headers")
        public ResponseEntity<StreamingResponseBody> downloadCsvCustomHeaders() {
                List<UserRecord> records = TestDataGenerator.generateUsers(1_000_000);

                // 6 headers: salary and department have no matching DTO field → blank columns
                List<String> headers = List.of(
                                "firstName", "email", "salary", "department", "createdDate", "id", "lastName");

                StreamingResponseBody body = out -> csvUtil.streamCsvToResponse(records, UserRecord.class, headers,
                                out);

                return ResponseEntity.ok()
                                .header(HttpHeaders.CONTENT_DISPOSITION,
                                                "attachment; filename=\"users_custom.csv\"")
                                .contentType(MediaType.parseMediaType("text/csv; charset=UTF-8"))
                                .body(body);
        }

        // ── Custom headers: generateCsvAsBytes ────────────────────────────────────
        /**
         * Returns custom-header CSV as byte[] — Content-Length set for exact
         * download-progress %.
         * Endpoint: GET /api/users/download/csv/custom-headers/bytes
         */
        @GetMapping("/download/csv/custom-headers/bytes")
        public ResponseEntity<byte[]> downloadCsvCustomHeadersAsBytes() throws IOException {
                List<UserRecord> records = TestDataGenerator.generateUsers(1_000_000);
                List<String> headers = List.of(
                                "firstName", "email", "salary", "department", "createdDate", "id", "lastName");
                byte[] bytes = csvUtil.generateCsvAsBytes(records, UserRecord.class, headers);
                return ResponseEntity.ok()
                                .header(HttpHeaders.CONTENT_DISPOSITION,
                                                "attachment; filename=\"users_custom.csv\"")
                                .header(HttpHeaders.CONTENT_LENGTH, String.valueOf(bytes.length))
                                .contentType(MediaType.parseMediaType("text/csv; charset=UTF-8"))
                                .body(bytes);
        }

        // ── Custom headers: generateCsvAsBase64 ──────────────────────────────────
        /**
         * Returns custom-header CSV as Base64 JSON { fileName, data }.
         * Endpoint: GET /api/users/download/csv/custom-headers/base64
         */
        @GetMapping("/download/csv/custom-headers/base64")
        public ResponseEntity<Map<String, String>> downloadCsvCustomHeadersAsBase64() throws IOException {
                List<UserRecord> records = TestDataGenerator.generateUsers(1_000_000);
                List<String> headers = List.of(
                                "firstName", "email", "salary", "department", "createdDate", "id", "lastName");
                String base64 = csvUtil.generateCsvAsBase64(records, UserRecord.class, headers);
                return ResponseEntity.ok(Map.of("fileName", "users_custom.csv", "data", base64));
        }

        // ── Custom headers: generateCsvToFile ─────────────────────────────────────
        /**
         * Saves custom-header CSV to server disk; returns JSON { method, message, path
         * }.
         * Endpoint: GET /api/users/download/csv/custom-headers/file
         */
        @GetMapping("/download/csv/custom-headers/file")
        public ResponseEntity<Map<String, String>> saveCsvCustomHeadersToFile() throws IOException {
                List<UserRecord> records = TestDataGenerator.generateUsers(1_000_000);
                List<String> headers = List.of(
                                "firstName", "email", "salary", "department", "createdDate", "id", "lastName");
                String path = "users_custom_export.csv";
                csvUtil.generateCsvToFile(records, UserRecord.class, headers, path);
                return ResponseEntity.ok(Map.of(
                                "method", "generateCsvToFile (custom headers)",
                                "message", "File saved on server",
                                "path", new java.io.File(path).getAbsolutePath()));
        }

        // =========================================================================
        // Excel endpoints (mirrors all CSV endpoints)
        // =========================================================================

        // ── xlsx/bytes ─────────────────────────────────────────────────────────
        /** Endpoint: GET /api/users/download/xlsx/bytes */
        @GetMapping("/download/xlsx/bytes")
        public ResponseEntity<byte[]> downloadExcelAsBytes() throws IOException {
                List<UserRecord> records = TestDataGenerator.generateUsers(100_000);
                byte[] bytes = excelUtil.generateExcelAsBytes(records, UserRecord.class);
                return ResponseEntity.ok()
                                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"users.xlsx\"")
                                .header(HttpHeaders.CONTENT_LENGTH, String.valueOf(bytes.length))
                                .contentType(MediaType.parseMediaType(
                                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                                .body(bytes);
        }

        // ── xlsx/base64 ───────────────────────────────────────────────────────
        /** Endpoint: GET /api/users/download/xlsx/base64 */
        @GetMapping("/download/xlsx/base64")
        public ResponseEntity<Map<String, String>> downloadExcelAsBase64() throws IOException {
                List<UserRecord> records = TestDataGenerator.generateUsers(100_000);
                String base64 = excelUtil.generateExcelAsBase64(records, UserRecord.class);
                return ResponseEntity.ok(Map.of("fileName", "users.xlsx", "data", base64));
        }

        // ── xlsx/file ──────────────────────────────────────────────────────────
        /** Endpoint: GET /api/users/download/xlsx/file */
        @GetMapping("/download/xlsx/file")
        public ResponseEntity<Map<String, String>> saveExcelToFile() throws IOException {
                List<UserRecord> records = TestDataGenerator.generateUsers(1_000_0000);
                String path = "users_export.xlsx";
                excelUtil.generateExcel(records, UserRecord.class, path);
                return ResponseEntity.ok(Map.of(
                                "method", "generateExcel",
                                "message", "Excel file saved on server",
                                "path", new java.io.File(path).getAbsolutePath()));
        }

        // ── xlsx/custom-headers ───────────────────────────────────────────────
        /**
         * Demonstrates custom-header Excel generation.
         * salary and department have no matching field in UserRecord → blank columns.
         *
         * Endpoint: GET /api/users/download/xlsx/custom-headers
         */
        @GetMapping("/download/xlsx/custom-headers")
        public ResponseEntity<StreamingResponseBody> downloadExcelCustomHeaders() {
                List<UserRecord> records = TestDataGenerator.generateUsers(10000);
                List<String> headers = List.of(
                                "id", "firstName", "email", "salary", "department", "createdDate");
                StreamingResponseBody body = out -> excelUtil.streamExcelToResponse(records, UserRecord.class, headers,
                                out);
                return ResponseEntity.ok()
                                .header(HttpHeaders.CONTENT_DISPOSITION,
                                                "attachment; filename=\"users_custom.xlsx\"")
                                .contentType(MediaType.parseMediaType(
                                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                                .body(body);
        }

        // ── xlsx/custom-headers/bytes ─────────────────────────────────────────────
        /**
         * Returns custom-header Excel as byte[] — Content-Length set.
         * Endpoint: GET /api/users/download/xlsx/custom-headers/bytes
         */
        @GetMapping("/download/xlsx/custom-headers/bytes")
        public ResponseEntity<byte[]> downloadExcelCustomHeadersAsBytes() throws IOException {
                List<UserRecord> records = TestDataGenerator.generateUsers(10000);
                List<String> headers = List.of(
                                "id", "firstName", "email", "salary", "department", "createdDate");
                byte[] bytes = excelUtil.generateExcelAsBytes(records, UserRecord.class, headers);
                return ResponseEntity.ok()
                                .header(HttpHeaders.CONTENT_DISPOSITION,
                                                "attachment; filename=\"users_custom.xlsx\"")
                                .header(HttpHeaders.CONTENT_LENGTH, String.valueOf(bytes.length))
                                .contentType(MediaType.parseMediaType(
                                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                                .body(bytes);
        }

        // ── xlsx/custom-headers/base64 ────────────────────────────────────────────
        /**
         * Returns custom-header Excel as Base64 JSON { fileName, data }.
         * Endpoint: GET /api/users/download/xlsx/custom-headers/base64
         */
        @GetMapping("/download/xlsx/custom-headers/base64")
        public ResponseEntity<Map<String, String>> downloadExcelCustomHeadersAsBase64() throws IOException {
                List<UserRecord> records = TestDataGenerator.generateUsers(10000);
                List<String> headers = List.of(
                                "id", "firstName", "email", "salary", "department", "createdDate");
                String base64 = excelUtil.generateExcelAsBase64(records, UserRecord.class, headers);
                return ResponseEntity.ok(Map.of("fileName", "users_custom.xlsx", "data", base64));
        }

        // ── xlsx/custom-headers/file ──────────────────────────────────────────────
        /**
         * Saves custom-header Excel to server disk; returns JSON { method, message,
         * path }.
         * Endpoint: GET /api/users/download/xlsx/custom-headers/file
         */
        @GetMapping("/download/xlsx/custom-headers/file")
        public ResponseEntity<Map<String, String>> saveExcelCustomHeadersToFile() throws IOException {
                List<UserRecord> records = TestDataGenerator.generateUsers(10000);
                List<String> headers = List.of(
                                "id", "firstName", "email", "salary", "department", "createdDate");
                String path = "users_custom_export.xlsx";
                excelUtil.generateExcel(records, UserRecord.class, headers, path);
                return ResponseEntity.ok(Map.of(
                                "method", "generateExcel (custom headers)",
                                "message", "Excel file saved on server",
                                "path", new java.io.File(path).getAbsolutePath()));
        }
}
