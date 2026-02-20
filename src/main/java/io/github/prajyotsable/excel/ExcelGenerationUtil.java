package io.github.prajyotsable.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;

public class ExcelGenerationUtil {

    /**
     * Rows held in heap before flushing to the on-disk temp file.
     * 5 000 → only ~200 flushes for 1 M rows (vs 1 000 flushes at window=1000).
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

    /**
     * Generates an .xlsx file from a list of objects in a memory-efficient way.
     * <p>
     * Uses {@link SXSSFWorkbook} (POI Streaming API): only {@code ROW_ACCESS_WINDOW}
     * rows are kept in the JVM heap at any time; the rest are written to a temp
     * file on disk.
     * <p>
     * Column headers are derived automatically from the declared field names of
     * {@code type}.  Supported cell types: {@code Number}, {@code Boolean},
     * {@code LocalDate}, {@code LocalDateTime}, and everything else via
     * {@code toString()}.
     *
     * @param <T>      type of the record objects
     * @param records  non-null, non-empty list of records to export
     * @param type     {@code Class<T>} used to read field metadata
     * @param fileName output file path, e.g. {@code "output/users.xlsx"}
     * @throws IOException              if the file cannot be written
     * @throws IllegalArgumentException if {@code records} is null or empty
     */
    /**
     * Generates an .xlsx file written to disk.
     */
    public <T> void generateExcel(List<T> records, Class<T> type, String fileName)
            throws IOException {
        if (records == null || records.isEmpty()) {
            throw new IllegalArgumentException("Records list must not be null or empty");
        }
        try (BufferedOutputStream bos = new BufferedOutputStream(
                new FileOutputStream(fileName), OUTPUT_BUFFER_SIZE)) {
            buildAndWrite(records, type, bos);
        }
        System.out.println("Excel generated successfully: " + fileName);
    }

    /**
     * Streams the generated .xlsx directly into any {@link OutputStream}.
     * <p>
     * Designed for HTTP responses: Spring's {@code StreamingResponseBody} passes
     * the raw response socket stream here, so rows flow straight to the client —
     * zero intermediate byte[] buffer in the JVM heap.
     *
     * @param records        list of records to export
     * @param type           {@code Class<T>}
     * @param responseStream target stream (e.g. from {@code HttpServletResponse})
     */
    public <T> void streamExcelToResponse(List<T> records, Class<T> type, OutputStream responseStream)
            throws IOException {
        if (records == null || records.isEmpty()) {
            throw new IllegalArgumentException("Records list must not be null or empty");
        }
        buildAndWrite(records, type, responseStream);
    }

    // ── Private core ──────────────────────────────────────────────────────────

    /**
     * Builds the SXSSF workbook and writes it to {@code outputStream}.
     * Shared by {@link #generateExcel} (file) and
     * {@link #streamExcelToResponse} (HTTP response).
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
}
