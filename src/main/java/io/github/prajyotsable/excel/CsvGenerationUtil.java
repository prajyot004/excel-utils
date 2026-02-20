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
import java.util.List;

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

    /** 256 KB buffer — batches writes, minimises OS syscall count. */
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
    //  3. HTTP STREAM  –  recommended for large / frontend downloads
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

        Field[]        fields  = type.getDeclaredFields();
        MethodHandle[] getters = buildGetters(fields); // built once, reused every row

        // NOT try-with-resources: must not close the caller's stream
        BufferedWriter writer = new BufferedWriter(
                new OutputStreamWriter(outputStream, StandardCharsets.UTF_8), BUFFER_SIZE);

        // Header
        writer.write(buildHeaderLine(fields));
        writer.newLine();

        // Data rows – reuse one StringBuilder, zero per-row heap allocation
        StringBuilder sb    = new StringBuilder(fields.length * 16);
        int           total = records.size();

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

    /** Wraps {@code value} in double-quotes if it contains , " \r \n; doubles any " inside. */
    private static void appendEscaped(StringBuilder sb, String value) {
        boolean needsQuoting = value.indexOf(',')  >= 0
                            || value.indexOf('"')  >= 0
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
        MethodHandles.Lookup lookup  = MethodHandles.lookup();
        MethodHandle[]       handles = new MethodHandle[fields.length];
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
}
