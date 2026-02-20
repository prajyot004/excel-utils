//package io.github.prajyotsable;
//
//import io.github.prajyotsable.excel.CsvGenerationUtil;
//import io.github.prajyotsable.excel.TestDataGenerator;
//import io.github.prajyotsable.excel.UserRecord;
//
//import java.io.IOException;
//import java.util.Date;
//import java.util.List;
//
//public class App {
//
//    public static void main(String[] args) throws IOException {
//
//        List<UserRecord> users = TestDataGenerator.generateUsers(1_000_000);
//        System.out.println("Size of data: " + users.size());
//
//        CsvGenerationUtil util = new CsvGenerationUtil();
//
//        // 1. Write to disk
//        benchmark("File write", () ->
//                util.generateCsvToFile(users, UserRecord.class, "users.csv"));
//
//        // 2. Return as byte[] (simulates small/medium REST response)
//        benchmark("Byte array", () -> {
//            byte[] csv = util.generateCsvAsBytes(users, UserRecord.class);
//            System.out.println("  byte[] size: " + csv.length + " bytes");
//        });
//
//        // 3. Stream to OutputStream (simulates large REST/frontend response)
//        benchmark("Stream to response", () ->
//                util.generateCsvToFile(users, UserRecord.class, "users_streamed.csv"));
//    }
//
//    private static void benchmark(String label, ThrowingRunnable task) throws IOException {
//        System.out.println("[" + label + "] started at: " + new Date());
//        long start = System.nanoTime();
//        task.run();
//        System.out.println("[" + label + "] done in: " + (System.nanoTime() - start) / 1_000_000 + " ms\n");
//    }
//
//    @FunctionalInterface
//    interface ThrowingRunnable { void run() throws IOException; }
//}
