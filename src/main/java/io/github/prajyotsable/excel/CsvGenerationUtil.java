package io.github.prajyotsable.excel;

import java.io.BufferedWriter;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.reflect.Field;
import java.nio.charset.StandardCharsets;
import java.util.Base64;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Memory-efficient CSV writer for any list of POJOs.
 *
 * <p>Three public methods, all delegating to the same private
 * {@link #writeContent} core so CSV logic lives in exactly one place:
 *
 * <ol>
 *   <li>{@link #generateCsvToFile}   – write directly to a file on disk.</li>
 *   <li>{@link #generateCsvAsBytes}  – return CSV as {@code byte[]}.</li>
 *   <li>{@link #streamCsvToResponse} – pipe directly into an HTTP response
 *       {@link OutputStream} (recommended for large / frontend downloads).</li>
 * </ol>
 *
 * <p>RFC 4180 escaping: fields containing {@code ,}, {@code "}, {@code \r},
 * or {@code \n} are double-quoted; any {@code "} inside is doubled.
 */
public class CsvGenerationUtil {

    /**
     * 256 KB buffer — batches writes, minimises OS syscall count.
     */
    private static final int BUFFER_SIZE = 256 * 1024;

    // =========================================================================
    //  1. FILE  –  write directly to disk
    // =========================================================================

    /**
     * Generates a CSV file and saves it to disk.
     * Best for: batch jobs, scheduled exports, CLI tools.
     */
    public <T> void generateCsvToFile(List<T> records, Class<T> type, String fileName)
            throws IOException {
        validate(records);
        try (FileOutputStream fos = new FileOutputStream(fileName)) {
            writeContent(records, type, fos);
        }
        System.out.println("CSV saved to file: " + fileName);
    }

    // =========================================================================
    //  2. BYTE ARRAY  –  for small / medium REST responses
    // =========================================================================

    /**
     * Generates the CSV entirely in memory and returns it as a {@code byte[]}.
     *
     * <p>Use in a Spring Boot controller like this:
     * <pre>{@code
     * byte[] csv = csvUtil.generateCsvAsBytes(records, UserRecord.class);
     * return ResponseEntity.ok()
     *     .header("Content-Disposition", "attachment; filename=\"users.csv\"")
     *     .contentType(MediaType.parseMediaType("text/csv"))
     *     .body(csv);
     * }</pre>
     *
     * <p><b>Pros:</b> simple; {@code Content-Length} header can be set so the
     * browser shows exact download progress %.
     * <br><b>Cons:</b> entire file lives in heap – risk of OOM for very large
     * exports. Keep to &lt; 500 K rows / ~50 MB.
     */
    public <T> byte[] generateCsvAsBytes(List<T> records, Class<T> type)
            throws IOException {
        validate(records);
        // pre-size at ~64 bytes per row to avoid costly internal reallocations
        ByteArrayOutputStream baos = new ByteArrayOutputStream(records.size() * 64);
        writeContent(records, type, baos);
        return baos.toByteArray();
    }

    // =========================================================================
    //  3. BASE64  –  for embedding in JSON / API responses
    // =========================================================================

    /**
     * Generates the CSV as a Base64-encoded {@link String}.
     *
     * <p>Useful when the file must travel inside a JSON body
     * (e.g. a REST API that returns {@code { "file": "<base64>" }}),
     * in e-mail attachments, or in data-URI links.
     *
     * <p>Use in a Spring Boot controller like this:
     * <pre>{@code
     * String b64 = csvUtil.generateCsvAsBase64(records, UserRecord.class);
     * return Map.of("fileName", "users.csv", "data", b64);
     * }</pre>
     *
     * On the frontend you can decode and download it without a server round-trip:
     * <pre>{@code
     * const a = document.createElement('a');
     * a.href     = 'data:text/csv;base64,' + response.data;
     * a.download = response.fileName;
     * a.click();
     * }</pre>
     *
     * <p><b>Pros:</b> embeds cleanly in JSON; works across any transport.
     * <br><b>Cons:</b> ~33 % size overhead vs raw bytes; entire file in heap
     * (same constraint as {@link #generateCsvAsBytes}).
     */
    public <T> String generateCsvAsBase64(List<T> records, Class<T> type)
            throws IOException {
        byte[] bytes = generateCsvAsBytes(records, type);
        return Base64.getEncoder().encodeToString(bytes);
    }

    // =========================================================================
    //  4. HTTP STREAM  –  recommended for large / frontend downloads
    // =========================================================================

    /**
     * Streams CSV rows directly into the HTTP response {@link OutputStream}.
     *
     * <p>Only {@value #BUFFER_SIZE} bytes live in heap at any time regardless
     * of data size – rows flow straight from generation to the network socket.
     *
     * <p>Use in a Spring Boot controller like this:
     * <pre>{@code
     * @GetMapping("/download")
     * public ResponseEntity<StreamingResponseBody> download() {
     *     List<UserRecord> records = userService.findAll();
     *     StreamingResponseBody body = out ->
     *         csvUtil.streamCsvToResponse(records, UserRecord.class, out);
     *     return ResponseEntity.ok()
     *         .header("Content-Disposition", "attachment; filename=\"users.csv\"")
     *         .contentType(MediaType.parseMediaType("text/csv; charset=UTF-8"))
     *         .body(body);
     * }
     * }</pre>
     *
     * <p><b>Pros:</b> zero extra heap; download starts in the browser
     * immediately; safe for any data size.
     * <br><b>Cons:</b> {@code Content-Length} cannot be set upfront (browser
     * shows bytes downloaded, not %).
     *
     * <p><b>NOTE:</b> the caller owns the stream lifecycle –
     * do <em>not</em> close {@code responseStream} inside this method.
     */
    public <T> void streamCsvToResponse(List<T> records, Class<T> type,
                                        OutputStream responseStream)
            throws IOException {
        validate(records);
        writeContent(records, type, responseStream);
    }

    // =========================================================================
    //  Private core – all three public methods delegate here
    // =========================================================================

    /**
     * Writes CSV header + all data rows into {@code outputStream}.
     * Does NOT close the stream – the caller owns the lifecycle.
     */
    private <T> void writeContent(List<T> records, Class<T> type,
                                  OutputStream outputStream) throws IOException {

        Field[] fields = type.getDeclaredFields();
        MethodHandle[] getters = buildGetters(fields); // built once, reused every row

        // NOT try-with-resources: must not close the caller's stream
        BufferedWriter writer = new BufferedWriter(
                new OutputStreamWriter(outputStream, StandardCharsets.UTF_8), BUFFER_SIZE);

        // Header
        writer.write(buildHeaderLine(fields));
        writer.newLine();

        // Data rows – reuse one StringBuilder, zero per-row heap allocation
        StringBuilder sb = new StringBuilder(fields.length * 16);
        int total = records.size();

        for (int i = 0; i < total; i++) {
            T record = records.get(i);
            sb.setLength(0); // reset without reallocating

            for (int col = 0; col < fields.length; col++) {
                if (col > 0) sb.append(',');
                try {
                    Object val = getters[col].invoke(record);
                    appendEscaped(sb, val == null ? "" : val.toString());
                } catch (Throwable e) {
                    sb.append("N/A");
                }
            }

            writer.write(sb.toString());
            writer.newLine();
        }

        writer.flush(); // push all buffered bytes to the underlying stream
    }

    // =========================================================================
    //  5. CUSTOM HEADERS  –  overloads for all 4 output types
    //
    //  Pass the column names you want in the CSV.  Each name is matched
    //  case-sensitively against the DTO field names:
    //    • match found  → value from the DTO field
    //    • no match     → blank cell (column still appears in the file)
    //  DTO fields that are NOT in the headers list are silently ignored.
    // =========================================================================

    /**
     * Writes a CSV file to disk using a caller-supplied ordered list of column
     * headers.  Headers that do not correspond to any field of {@code type}
     * produce a blank column.
     *
     * @param headers ordered column names, e.g. {@code List.of("id","email","salary")}
     */
    public <T> void generateCsvToFile(List<T> records, Class<T> type,
                                      List<String> headers, String fileName)
            throws IOException {
        validate(records);
        validateHeaders(headers);
        try (FileOutputStream fos = new FileOutputStream(fileName)) {
            writeContentWithHeaders(records, type, headers, fos);
        }
        System.out.println("CSV (custom headers) saved to file: " + fileName);
    }

    /** Returns the CSV as {@code byte[]} using caller-supplied column headers. */
    public <T> byte[] generateCsvAsBytes(List<T> records, Class<T> type,
                                         List<String> headers)
            throws IOException {
        validate(records);
        validateHeaders(headers);
        ByteArrayOutputStream baos = new ByteArrayOutputStream(records.size() * 64);
        writeContentWithHeaders(records, type, headers, baos);
        return baos.toByteArray();
    }

    /** Returns the CSV as a Base64 string using caller-supplied column headers. */
    public <T> String generateCsvAsBase64(List<T> records, Class<T> type,
                                          List<String> headers)
            throws IOException {
        return Base64.getEncoder().encodeToString(
                generateCsvAsBytes(records, type, headers));
    }

    /**
     * Streams the CSV with caller-supplied column headers into any
     * {@link OutputStream} (e.g. an HTTP response socket).
     */
    public <T> void streamCsvToResponse(List<T> records, Class<T> type,
                                        List<String> headers,
                                        OutputStream responseStream)
            throws IOException {
        validate(records);
        validateHeaders(headers);
        writeContentWithHeaders(records, type, headers, responseStream);
    }

    // =========================================================================
    //  Private core (custom headers)
    // =========================================================================

    /**
     * Writes CSV content using a caller-supplied ordered list of column names.
     *
     * <p>Algorithm:
     * <ol>
     *   <li>Build a {@code fieldName → MethodHandle} map from every declared
     *       field of {@code type}.</li>
     *   <li>Walk {@code headers} in order; for each name look up its handle.
     *       Store {@code null} when no matching field exists.</li>
     *   <li>Write the header row using the supplied names verbatim.</li>
     *   <li>For every data row: invoke the handle if non-null, otherwise
     *       write an empty string (blank cell).</li>
     * </ol>
     */
    private <T> void writeContentWithHeaders(List<T> records, Class<T> type,
                                             List<String> headers,
                                             OutputStream outputStream) throws IOException {

        // ─ 1. Build fieldName → MethodHandle map ────────────────────────────────
        Field[] allFields = type.getDeclaredFields();
        Map<String, MethodHandle> handleMap = new HashMap<>(allFields.length * 2);
        MethodHandles.Lookup lookup = MethodHandles.lookup();
        for (Field f : allFields) {
            f.setAccessible(true);
            try {
                handleMap.put(f.getName(), lookup.unreflectGetter(f));
            } catch (IllegalAccessException e) {
                throw new RuntimeException("Cannot create getter for: " + f.getName(), e);
            }
        }

        // ─ 2. Resolve handles in header order (null = no matching DTO field) ───
        int cols = headers.size();
        MethodHandle[] getters = new MethodHandle[cols];
        for (int i = 0; i < cols; i++) {
            getters[i] = handleMap.get(headers.get(i)); // null if not found in DTO
        }

        // NOT try-with-resources: must not close the caller's stream
        BufferedWriter writer = new BufferedWriter(
                new OutputStreamWriter(outputStream, StandardCharsets.UTF_8), BUFFER_SIZE);

        // ─ 3. Header row ─────────────────────────────────────────────────
        StringBuilder sb = new StringBuilder(cols * 16);
        for (int i = 0; i < cols; i++) {
            if (i > 0) sb.append(',');
            appendEscaped(sb, headers.get(i));
        }
        writer.write(sb.toString());
        writer.newLine();

        // ─ 4. Data rows ─────────────────────────────────────────────────
        int total = records.size();
        for (int i = 0; i < total; i++) {
            T record = records.get(i);
            sb.setLength(0);
            for (int col = 0; col < cols; col++) {
                if (col > 0) sb.append(',');
                if (getters[col] == null) {
                    // header has no matching DTO field → blank cell
                    continue;
                }
                try {
                    Object val = getters[col].invoke(record);
                    appendEscaped(sb, val == null ? "" : val.toString());
                } catch (Throwable e) {
                    sb.append("N/A");
                }
            }
            writer.write(sb.toString());
            writer.newLine();
        }

        writer.flush();
    }

    // =========================================================================
    //  RFC 4180 helpers
    // =========================================================================

    private static String buildHeaderLine(Field[] fields) {
        StringBuilder sb = new StringBuilder(fields.length * 12);
        for (int i = 0; i < fields.length; i++) {
            if (i > 0) sb.append(',');
            appendEscaped(sb, fields[i].getName());
        }
        return sb.toString();
    }

    /**
     * Wraps {@code value} in double-quotes if it contains , " \r \n; doubles any " inside.
     */
    private static void appendEscaped(StringBuilder sb, String value) {
        boolean needsQuoting = value.indexOf(',') >= 0
                || value.indexOf('"') >= 0
                || value.indexOf('\n') >= 0
                || value.indexOf('\r') >= 0;
        if (!needsQuoting) {
            sb.append(value);
            return;
        }
        sb.append('"');
        for (int i = 0; i < value.length(); i++) {
            char c = value.charAt(i);
            if (c == '"') sb.append('"'); // double the quote per RFC 4180
            sb.append(c);
        }
        sb.append('"');
    }

    // =========================================================================
    //  MethodHandle builder
    // =========================================================================

    /**
     * Builds a {@link MethodHandle} getter for each field.
     * ~3-5x faster than {@code Field.get()} after JIT inlining –
     * critical for millions of invocations on large exports.
     */
    private static MethodHandle[] buildGetters(Field[] fields) {
        MethodHandles.Lookup lookup = MethodHandles.lookup();
        MethodHandle[] handles = new MethodHandle[fields.length];
        for (int i = 0; i < fields.length; i++) {
            fields[i].setAccessible(true);
            try {
                handles[i] = lookup.unreflectGetter(fields[i]);
            } catch (IllegalAccessException e) {
                throw new RuntimeException(
                        "Cannot create getter for: " + fields[i].getName(), e);
            }
        }
        return handles;
    }

    // =========================================================================
    //  Guard
    // =========================================================================

    private static void validate(List<?> records) {
        if (records == null || records.isEmpty())
            throw new IllegalArgumentException("Records list must not be null or empty");
    }

    private static void validateHeaders(List<String> headers) {
        if (headers == null || headers.isEmpty())
            throw new IllegalArgumentException("Headers list must not be null or empty");
    }
}
