package io.github.prajyotsable.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.BufferedOutputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Base64;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelGenerationUtil {

    /**
     * Rows held in heap before flushing to the on-disk temp file.
     * 500 → ~2 000 flushes for 1 M rows (vs ~1 000 flushes at window=1000,
     * and ~200 flushes at window=5 000).
     */
    private static final int ROW_ACCESS_WINDOW  = 500;

    /** 64 KB write buffer — batches final file writes, cuts OS syscall count. */
    private static final int OUTPUT_BUFFER_SIZE = 64 * 1024;

    /**
     * Pre-computed per-column type category.
     * Avoids re-running a 4-branch {@code instanceof} chain on every single cell
     * (= 6 M checks for 1 M rows × 6 fields).
     */
    private enum FieldType { NUMBER, BOOLEAN, TEMPORAL, OTHER }

    // =========================================================================
    //  1. FILE  –  write directly to disk
    // =========================================================================

    /** Generates an .xlsx file and saves it to disk. */
    public <T> void generateExcel(List<T> records, Class<T> type, String fileName)
            throws IOException {
        validate(records);
        try (BufferedOutputStream bos = new BufferedOutputStream(
                new FileOutputStream(fileName), OUTPUT_BUFFER_SIZE)) {
            buildAndWrite(records, type, bos);
        }
        System.out.println("Excel generated successfully: " + fileName);
    }

    // =========================================================================
    //  2. STREAM  –  pipe directly into any OutputStream (HTTP response)
    // =========================================================================

    /** Streams the .xlsx directly into the caller’s {@link OutputStream} (e.g. HTTP response). */
    public <T> void streamExcelToResponse(List<T> records, Class<T> type,
                                          OutputStream responseStream)
            throws IOException {
        validate(records);
        buildAndWrite(records, type, responseStream);
    }

    // =========================================================================
    //  3. BYTE ARRAY  –  for small / medium REST responses
    //
    //  The entire workbook lives in heap. Keep to ≤ 100 K rows / ~50 MB.
    //  Benefit: Spring can set Content-Length, so the browser shows exact %.
    // =========================================================================

    /** Returns the .xlsx as a {@code byte[]} (Content-Length friendly). */
    public <T> byte[] generateExcelAsBytes(List<T> records, Class<T> type)
            throws IOException {
        validate(records);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        buildAndWrite(records, type, baos);
        return baos.toByteArray();
    }

    // =========================================================================
    //  4. BASE64  –  for JSON responses / data-URIs / e-mail attachments
    // =========================================================================

    /**
     * Returns the .xlsx as a Base64-encoded {@link String}.
     * Use in a JSON response: {@code { "fileName": "users.xlsx", "data": "<base64>" }}
     */
    public <T> String generateExcelAsBase64(List<T> records, Class<T> type)
            throws IOException {
        return Base64.getEncoder().encodeToString(generateExcelAsBytes(records, type));
    }

    // =========================================================================
    //  5. CUSTOM HEADERS  –  overloads for all 4 output types
    //
    //  Pass the ordered column names you want in the sheet.  Each name is
    //  matched case-sensitively against DTO field names:
    //    • match found  → value from the DTO field
    //    • no match     → blank cell  (column still appears with the given header)
    //  DTO fields that are NOT in the headers list are silently ignored.
    // =========================================================================

    /** Saves .xlsx to disk using caller-supplied ordered column headers. */
    public <T> void generateExcel(List<T> records, Class<T> type,
                                  List<String> headers, String fileName)
            throws IOException {
        validate(records);
        validateHeaders(headers);
        try (BufferedOutputStream bos = new BufferedOutputStream(
                new FileOutputStream(fileName), OUTPUT_BUFFER_SIZE)) {
            buildAndWriteWithHeaders(records, type, headers, bos);
        }
        System.out.println("Excel (custom headers) generated: " + fileName);
    }

    /** Returns the .xlsx as {@code byte[]} using caller-supplied column headers. */
    public <T> byte[] generateExcelAsBytes(List<T> records, Class<T> type,
                                           List<String> headers)
            throws IOException {
        validate(records);
        validateHeaders(headers);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        buildAndWriteWithHeaders(records, type, headers, baos);
        return baos.toByteArray();
    }

    /** Returns the .xlsx as a Base64 string using caller-supplied column headers. */
    public <T> String generateExcelAsBase64(List<T> records, Class<T> type,
                                            List<String> headers)
            throws IOException {
        return Base64.getEncoder().encodeToString(
                generateExcelAsBytes(records, type, headers));
    }

    /** Streams the .xlsx with caller-supplied column headers into any {@link OutputStream}. */
    public <T> void streamExcelToResponse(List<T> records, Class<T> type,
                                          List<String> headers,
                                          OutputStream responseStream)
            throws IOException {
        validate(records);
        validateHeaders(headers);
        buildAndWriteWithHeaders(records, type, headers, responseStream);
    }

    // =========================================================================
    //  Private core (auto headers — field names used as column names)
    // =========================================================================

    /**
     * Builds the SXSSF workbook and writes it to {@code outputStream}.
     * Shared by all auto-header public methods.
     */
    private <T> void buildAndWrite(List<T> records, Class<T> type, OutputStream outputStream)
            throws IOException {

        Field[]        fields  = type.getDeclaredFields();
        MethodHandle[] getters = buildGetters(fields);  // built once, reused every row
        FieldType[]    ftypes  = detectTypes(fields);   // built once, reused every cell

        try (SXSSFWorkbook workbook = new SXSSFWorkbook(ROW_ACCESS_WINDOW)) {

            // Compression OFF: gzip-ing each flush adds ~10-20 s for 1 M rows.
            workbook.setCompressTempFiles(false);

            Sheet sheet = workbook.createSheet(type.getSimpleName());

            // ── Header row ────────────────────────────────────────────────────
            CellStyle headerStyle = buildHeaderStyle(workbook);
            Row headerRow = sheet.createRow(0);
            for (int col = 0; col < fields.length; col++) {
                Cell cell = headerRow.createCell(col);
                cell.setCellValue(fields[col].getName());
                cell.setCellStyle(headerStyle);
            }

            // ── Data rows ─────────────────────────────────────────────────────
            // Plain indexed for-loop: no Stream wrapper, no AtomicInteger CAS,
            // no lambda allocation — measurably faster over 1 M iterations.
            int totalRows = records.size();
            for (int i = 0; i < totalRows; i++) {
                T   record = records.get(i);
                Row row    = sheet.createRow(i + 1);
                for (int col = 0; col < fields.length; col++) {
                    Cell cell = row.createCell(col);
                    try {
                        setCellValue(cell, getters[col].invoke(record), ftypes[col]);
                    } catch (Throwable e) {
                        cell.setCellValue("N/A");
                    }
                }
            }

            // ── Write to stream ───────────────────────────────────────────────
            workbook.write(outputStream);
            workbook.dispose(); // delete POI temp files from disk
        }
    }

    // =========================================================================
    //  Private core (custom headers)
    // =========================================================================

    /**
     * Builds the SXSSF workbook using caller-supplied column headers and writes
     * it to {@code outputStream}.
     *
     * <p>Algorithm:
     * <ol>
     *   <li>Build {@code fieldName → MethodHandle} and {@code fieldName → FieldType}
     *       maps from every declared field of {@code type}.</li>
     *   <li>Walk {@code headers} in order; for each name look up its handle and type.
     *       Store {@code null} handle when no matching field exists.</li>
     *   <li>Write the header row using the supplied names verbatim (bold style).</li>
     *   <li>For every data row: invoke the handle if non-null, otherwise write an
     *       empty string (blank cell).</li>
     * </ol>
     */
    private <T> void buildAndWriteWithHeaders(List<T> records, Class<T> type,
                                              List<String> headers,
                                              OutputStream outputStream)
            throws IOException {

        // 1. Build per-field maps from all declared fields
        Field[] allFields = type.getDeclaredFields();
        Map<String, MethodHandle> handleMap = new HashMap<>(allFields.length * 2);
        Map<String, FieldType>    typeMap   = new HashMap<>(allFields.length * 2);
        MethodHandles.Lookup lookup = MethodHandles.lookup();
        for (Field f : allFields) {
            f.setAccessible(true);
            try {
                handleMap.put(f.getName(), lookup.unreflectGetter(f));
                typeMap.put(f.getName(), detectType(f));
            } catch (IllegalAccessException e) {
                throw new RuntimeException("Cannot create getter for: " + f.getName(), e);
            }
        }

        // 2. Resolve handle + type in header order (null handle = no matching DTO field)
        int cols = headers.size();
        MethodHandle[] getters = new MethodHandle[cols];
        FieldType[]    ftypes  = new FieldType[cols];
        for (int i = 0; i < cols; i++) {
            getters[i] = handleMap.get(headers.get(i));
            ftypes[i]  = typeMap.getOrDefault(headers.get(i), FieldType.OTHER);
        }

        try (SXSSFWorkbook workbook = new SXSSFWorkbook(ROW_ACCESS_WINDOW)) {
            workbook.setCompressTempFiles(false);
            Sheet sheet = workbook.createSheet(type.getSimpleName());

            // 3. Header row — use passed names verbatim
            CellStyle headerStyle = buildHeaderStyle(workbook);
            Row headerRow = sheet.createRow(0);
            for (int col = 0; col < cols; col++) {
                Cell cell = headerRow.createCell(col);
                cell.setCellValue(headers.get(col));
                cell.setCellStyle(headerStyle);
            }

            // 4. Data rows
            int totalRows = records.size();
            for (int i = 0; i < totalRows; i++) {
                T   record = records.get(i);
                Row row    = sheet.createRow(i + 1);
                for (int col = 0; col < cols; col++) {
                    Cell cell = row.createCell(col);
                    if (getters[col] == null) {
                        cell.setCellValue(""); // no matching DTO field → blank
                        continue;
                    }
                    try {
                        setCellValue(cell, getters[col].invoke(record), ftypes[col]);
                    } catch (Throwable e) {
                        cell.setCellValue("N/A");
                    }
                }
            }

            workbook.write(outputStream);
            workbook.dispose();
        }
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    /**
     * Builds a {@link MethodHandle} getter for each declared field.
     * <p>
     * MethodHandles are ~3-5× faster than reflective {@code Field.get()} per
     * invocation once the JIT inlines them — critical when called 6 M times for
     * 1 M rows × 6 fields.
     */
    private static MethodHandle[] buildGetters(Field[] fields) {
        MethodHandles.Lookup lookup = MethodHandles.lookup();
        MethodHandle[] handles = new MethodHandle[fields.length];
        for (int i = 0; i < fields.length; i++) {
            fields[i].setAccessible(true);
            try {
                handles[i] = lookup.unreflectGetter(fields[i]);
            } catch (IllegalAccessException e) {
                throw new RuntimeException("Cannot create getter for: " + fields[i].getName(), e);
            }
        }
        return handles;
    }

    /**
     * Classifies each field's declared type once at setup time so that
     * {@link #setCellValue} never needs to run an {@code instanceof} chain.
     */
    private static FieldType[] detectTypes(Field[] fields) {
        FieldType[] types = new FieldType[fields.length];
        for (int i = 0; i < fields.length; i++) {
            Class<?> c = fields[i].getType();
            if (Number.class.isAssignableFrom(c)
                    || c == int.class || c == long.class || c == double.class
                    || c == float.class || c == short.class || c == byte.class) {
                types[i] = FieldType.NUMBER;
            } else if (c == boolean.class || c == Boolean.class) {
                types[i] = FieldType.BOOLEAN;
            } else if (c == LocalDate.class || c == LocalDateTime.class) {
                types[i] = FieldType.TEMPORAL;
            } else {
                types[i] = FieldType.OTHER;
            }
        }
        return types;
    }

    /** Writes one cell value using the pre-computed column type — zero instanceof checks. */
    private static void setCellValue(Cell cell, Object value, FieldType type) {
        if (value == null) { cell.setCellValue(""); return; }
        switch (type) {
            case NUMBER  -> cell.setCellValue(((Number) value).doubleValue());
            case BOOLEAN -> cell.setCellValue((Boolean) value);
            default      -> cell.setCellValue(value.toString()); // TEMPORAL + OTHER
        }
    }

    /** Creates a bold {@link CellStyle} for the header row. */
    private static CellStyle buildHeaderStyle(SXSSFWorkbook workbook) {
        Font font = workbook.createFont();
        font.setBold(true);
        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        return style;
    }

    /** Classifies a single field’s declared type — used by custom-header core. */
    private static FieldType detectType(Field f) {
        Class<?> c = f.getType();
        if (Number.class.isAssignableFrom(c)
                || c == int.class || c == long.class || c == double.class
                || c == float.class || c == short.class || c == byte.class)
            return FieldType.NUMBER;
        if (c == boolean.class || c == Boolean.class)
            return FieldType.BOOLEAN;
        if (c == LocalDate.class || c == LocalDateTime.class)
            return FieldType.TEMPORAL;
        return FieldType.OTHER;
    }

    private static void validate(List<?> records) {
        if (records == null || records.isEmpty())
            throw new IllegalArgumentException("Records list must not be null or empty");
    }

    private static void validateHeaders(List<String> headers) {
        if (headers == null || headers.isEmpty())
            throw new IllegalArgumentException("Headers list must not be null or empty");
    }
}
