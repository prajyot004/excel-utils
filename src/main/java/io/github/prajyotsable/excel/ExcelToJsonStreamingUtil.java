package io.github.prajyotsable.excel;

import com.fasterxml.jackson.core.JsonFactory;
import com.fasterxml.jackson.core.JsonGenerator;
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.IOUtils;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;
import java.util.function.BiPredicate;

public class ExcelToJsonStreamingUtil {

    private static final String POI_OVERRIDE_PROPERTY = "excel.poi.byteArrayMaxOverride";
    private static final String JSON_FLUSH_EVERY_ROWS_PROPERTY = "excel.json.flushEveryRows";
    private static final int DEFAULT_POI_BYTE_ARRAY_MAX = 512 * 1024 * 1024;
    private static final int DEFAULT_JSON_FLUSH_EVERY_ROWS = 2000;
    private static volatile boolean poiByteArrayOverrideConfigured = false;

    private final JsonFactory jsonFactory = new JsonFactory();

    public void streamFirstSheetAsJson(InputStream excelInputStream, OutputStream outputStream) throws Exception {
        validateStreams(excelInputStream, outputStream);
        configurePoiByteArrayOverride();

        try (BufferedInputStream bis = new BufferedInputStream(excelInputStream);
                OPCPackage opcPackage = OPCPackage.open(bis);
                JsonGenerator generator = jsonFactory.createGenerator(outputStream)) {

            XSSFReader reader = new XSSFReader(opcPackage);
            StylesTable stylesTable = reader.getStylesTable();
            SharedStrings sharedStrings = reader.getSharedStringsTable();

            XSSFReader.SheetIterator sheetIterator = (XSSFReader.SheetIterator) reader.getSheetsData();
            if (!sheetIterator.hasNext()) {
                generator.writeStartObject();
                generator.writeStringField("sheetName", "");
                generator.writeArrayFieldStart("rows");
                generator.writeEndArray();
                generator.writeNumberField("recordCount", 0);
                generator.writeEndObject();
                generator.flush();
                return;
            }

            try (InputStream sheetStream = sheetIterator.next()) {
                String sheetName;
                try {
                    sheetName = sheetIterator.getSheetName();
                } catch (RuntimeException ex) {
                    sheetName = "";
                }

                ExcelSheetToJsonHandler handler = new ExcelSheetToJsonHandler(
                    generator,
                    sheetName,
                    0,
                    resolveJsonFlushEveryRows());
                XMLReader parser = SAXHelper.newXMLReader();
                DataFormatter formatter = new DataFormatter(Locale.ROOT);
                XSSFSheetXMLHandler xmlHandler = new XSSFSheetXMLHandler(
                        stylesTable,
                        null,
                        sharedStrings,
                        handler,
                        formatter,
                        false);
                parser.setContentHandler(xmlHandler);
                parser.parse(new InputSource(sheetStream));
                handler.finish();
                generator.flush();
            }
        }
    }

    public void streamAllSheetsAsJson(InputStream excelInputStream, OutputStream outputStream) throws Exception {
        streamSelectedSheetsAsJson(excelInputStream, outputStream, (index, name) -> true, false);
    }

    public void streamSheetByIndexAsJson(InputStream excelInputStream, OutputStream outputStream, int targetSheetIndex)
            throws Exception {
        if (targetSheetIndex < 0) {
            throw new IllegalArgumentException("sheetIndex must be >= 0");
        }
        streamSelectedSheetsAsJson(excelInputStream, outputStream,
                (index, name) -> index == targetSheetIndex,
                true);
    }

    public void streamSheetByNameAsJson(InputStream excelInputStream, OutputStream outputStream, String targetSheetName)
            throws Exception {
        String expectedName = targetSheetName == null ? "" : targetSheetName.trim();
        if (expectedName.isEmpty()) {
            throw new IllegalArgumentException("sheetName cannot be blank");
        }
        streamSelectedSheetsAsJson(excelInputStream, outputStream,
                (index, name) -> expectedName.equalsIgnoreCase(name),
                true);
    }

    private void streamSelectedSheetsAsJson(
            InputStream excelInputStream,
            OutputStream outputStream,
            BiPredicate<Integer, String> selector,
            boolean stopAfterFirstMatch) throws Exception {

        validateStreams(excelInputStream, outputStream);
        configurePoiByteArrayOverride();

        try (BufferedInputStream bis = new BufferedInputStream(excelInputStream);
                OPCPackage opcPackage = OPCPackage.open(bis);
                JsonGenerator generator = jsonFactory.createGenerator(outputStream)) {

            XSSFReader reader = new XSSFReader(opcPackage);
            StylesTable stylesTable = reader.getStylesTable();
            SharedStrings sharedStrings = reader.getSharedStringsTable();

            XSSFReader.SheetIterator sheetIterator = (XSSFReader.SheetIterator) reader.getSheetsData();

            generator.writeStartObject();
            generator.writeArrayFieldStart("sheets");

            int sheetIndex = 0;
            int matchedSheetCount = 0;
            long totalRecordCount = 0;
            int flushEveryRows = resolveJsonFlushEveryRows();
            
            while (sheetIterator.hasNext()) {
                try (InputStream sheetStream = sheetIterator.next()) {
                    String sheetName = resolveSheetName(sheetIterator);
                    boolean shouldProcess = selector.test(sheetIndex, sheetName);

                    if (shouldProcess) {
                        long sheetRecords = parseOneSheetToJson(
                                sheetStream,
                                sheetName,
                                sheetIndex,
                                stylesTable,
                                sharedStrings,
                                generator,
                                flushEveryRows);

                        matchedSheetCount++;
                        totalRecordCount += sheetRecords;

                        if (stopAfterFirstMatch) {
                            break;
                        }
                    }
                }
                sheetIndex++;
            }

            generator.writeEndArray();
            generator.writeNumberField("sheetCount", matchedSheetCount);
            generator.writeNumberField("totalRecordCount", totalRecordCount);
            generator.writeEndObject();
            generator.flush();
        }
    }

    private static void configurePoiByteArrayOverride() {
        if (poiByteArrayOverrideConfigured) {
            return;
        }
        synchronized (ExcelToJsonStreamingUtil.class) {
            if (poiByteArrayOverrideConfigured) {
                return;
            }

            int overrideValue = DEFAULT_POI_BYTE_ARRAY_MAX;
            String configured = System.getProperty(POI_OVERRIDE_PROPERTY);
            if (configured != null && !configured.trim().isEmpty()) {
                try {
                    int parsed = Integer.parseInt(configured.trim());
                    if (parsed > 0) {
                        overrideValue = parsed;
                    }
                } catch (NumberFormatException ignored) {
                    // Keep default when provided value is invalid.
                }
            }

            IOUtils.setByteArrayMaxOverride(overrideValue);
            poiByteArrayOverrideConfigured = true;
        }
    }

    private long parseOneSheetToJson(
            InputStream sheetStream,
            String sheetName,
            int sheetIndex,
            StylesTable stylesTable,
            SharedStrings sharedStrings,
            JsonGenerator generator,
            int flushEveryRows) throws Exception {

        ExcelSheetToJsonHandler handler = new ExcelSheetToJsonHandler(
                generator,
                sheetName,
                sheetIndex,
                flushEveryRows);
        XMLReader parser = SAXHelper.newXMLReader();
        DataFormatter formatter = new DataFormatter(Locale.ROOT);
        XSSFSheetXMLHandler xmlHandler = new XSSFSheetXMLHandler(
                stylesTable,
                null,
                sharedStrings,
                handler,
                formatter,
                false);
        parser.setContentHandler(xmlHandler);
        parser.parse(new InputSource(sheetStream));
        return handler.finish();
    }

    private int resolveJsonFlushEveryRows() {
        String configured = System.getProperty(JSON_FLUSH_EVERY_ROWS_PROPERTY);
        if (configured == null || configured.trim().isEmpty()) {
            return DEFAULT_JSON_FLUSH_EVERY_ROWS;
        }
        try {
            int parsed = Integer.parseInt(configured.trim());
            return Math.max(parsed, 0);
        } catch (NumberFormatException ignored) {
            return DEFAULT_JSON_FLUSH_EVERY_ROWS;
        }
    }

    private String resolveSheetName(XSSFReader.SheetIterator sheetIterator) {
        try {
            String sheetName = sheetIterator.getSheetName();
            return sheetName == null ? "" : sheetName;
        } catch (RuntimeException ex) {
            return "";
        }
    }

    private void validateStreams(InputStream excelInputStream, OutputStream outputStream) {
        if (excelInputStream == null) {
            throw new IllegalArgumentException("Excel input stream cannot be null");
        }
        if (outputStream == null) {
            throw new IllegalArgumentException("Output stream cannot be null");
        }
    }

    private static final class ExcelSheetToJsonHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

        private final JsonGenerator generator;
        private final List<String> headers = new ArrayList<>();
        private final List<String> currentRowValues = new ArrayList<>();
        private final Set<String> usedHeaderNames = new HashSet<>();
        private final int flushEveryRows;
        private int currentColumnIndex = -1;
        private long recordCount = 0;

        ExcelSheetToJsonHandler(JsonGenerator generator, String sheetName, int sheetIndex, int flushEveryRows)
                throws IOException {
            this.generator = generator;
            this.flushEveryRows = flushEveryRows;
            generator.writeStartObject();
            generator.writeNumberField("sheetIndex", sheetIndex);
            generator.writeStringField("sheetName", sheetName == null ? "" : sheetName);
            generator.writeArrayFieldStart("rows");
        }

        @Override
        public void startRow(int rowNum) {
            currentRowValues.clear();
            currentColumnIndex = -1;
        }

        @Override
        public void endRow(int rowNum) {
            try {
                if (rowNum == 0) {
                    buildHeaders(currentRowValues);
                    return;
                }
                if (headers.isEmpty()) {
                    buildDefaultHeaders(currentRowValues.size());
                }

                generator.writeStartObject();
                for (int i = 0; i < headers.size(); i++) {
                    String header = headers.get(i);
                    String value = i < currentRowValues.size() ? currentRowValues.get(i) : "";
                    generator.writeStringField(header, value);
                }
                generator.writeEndObject();
                recordCount++;
                if (flushEveryRows > 0 && recordCount % flushEveryRows == 0) {
                    generator.flush();
                }
            } catch (IOException e) {
                throw new RuntimeException("Failed to stream JSON row", e);
            }
        }

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            int cellColumnIndex = resolveColumnIndex(cellReference, currentColumnIndex + 1);
            while (currentRowValues.size() < cellColumnIndex) {
                currentRowValues.add("");
            }
            currentRowValues.add(Objects.toString(formattedValue, ""));
            currentColumnIndex = cellColumnIndex;
        }

        @Override
        public void headerFooter(String text, boolean isHeader, String tagName) {
        }

        long finish() throws IOException {
            generator.writeEndArray();
            generator.writeNumberField("recordCount", recordCount);
            generator.writeEndObject();
            return recordCount;
        }

        private void buildHeaders(List<String> firstRowValues) {
            headers.clear();
            usedHeaderNames.clear();

            for (int i = 0; i < firstRowValues.size(); i++) {
                String raw = Objects.toString(firstRowValues.get(i), "").trim();
                String candidate = raw.isEmpty() ? "column_" + (i + 1) : raw;
                candidate = deduplicate(candidate);
                headers.add(candidate);
            }
        }

        private void buildDefaultHeaders(int size) {
            headers.clear();
            usedHeaderNames.clear();
            for (int i = 0; i < size; i++) {
                headers.add(deduplicate("column_" + (i + 1)));
            }
        }

        private String deduplicate(String candidate) {
            String normalized = candidate;
            int suffix = 1;
            while (usedHeaderNames.contains(normalized)) {
                normalized = candidate + "_" + suffix;
                suffix++;
            }
            usedHeaderNames.add(normalized);
            return normalized;
        }

        private int resolveColumnIndex(String cellReference, int fallback) {
            if (cellReference == null || cellReference.isEmpty()) {
                return fallback;
            }
            int i = 0;
            int index = 0;
            while (i < cellReference.length() && Character.isLetter(cellReference.charAt(i))) {
                char c = Character.toUpperCase(cellReference.charAt(i));
                index = index * 26 + (c - 'A' + 1);
                i++;
            }
            return Math.max(0, index - 1);
        }
    }
}
