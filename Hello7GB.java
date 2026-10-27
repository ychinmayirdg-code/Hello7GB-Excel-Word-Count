
import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Hello50GB {

    static final String PATH = "D:\\Hello50GBFiles";

    // Target: approximately 10 GiB total
    static final long TARGET_SIZE = 7L * 1024 * 1024 * 1024;

    public static void main(String[] args) throws Exception {

        File folder = new File(PATH);

        if (!folder.exists() && !folder.mkdirs()) {
            System.out.println("Could not create output folder.");
            return;
        }

        // Count existing Excel files and find the next file number
        long totalSize = 0;
        int fileNumber = 1;

        File[] existing = folder.listFiles();

        if (existing != null) {
            for (File f : existing) {

                if (f.isFile() && f.getName().endsWith(".xlsx")) {
                    totalSize += f.length();

                    String name = f.getName();

                    if (name.startsWith("hello_")
                            && name.endsWith(".xlsx")) {
                        try {
                            int number = Integer.parseInt(
                                name.substring(6, name.length() - 5)
                            );

                            fileNumber = Math.max(
                                fileNumber, number + 1
                            );

                        } catch (NumberFormatException ignored) {
                            // Ignore filenames that do not match the pattern
                        }
                    }
                }
            }
        }

        System.out.printf(
            "Existing Excel files size: %.2f GiB%n",
            totalSize / (1024.0 * 1024 * 1024)
        );

        if (totalSize >= TARGET_SIZE) {
            System.out.println("Target size already reached!");
            return;
        }

        // Generate additional Excel files until the target is reached
        while (totalSize < TARGET_SIZE) {

            // Stop if less than 5 GiB of free space remains
            if (folder.getFreeSpace() < 5L * 1024 * 1024 * 1024) {
                System.out.println(
                    "Stopped: less than 5 GiB disk space remains."
                );
                break;
            }

            File file = new File(
                folder,
                String.format("hello_%06d.xlsx", fileNumber)
            );

            try (Workbook workbook = new XSSFWorkbook()) {

                Sheet sheet = workbook.createSheet("Data");

                String filler =
                    "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789".repeat(400);

                for (int i = 0; i < 500; i++) {

                    Row row = sheet.createRow(i);

                    row.createCell(0).setCellValue("Hello");

                    row.createCell(1).setCellValue(
                        "Hello data row " + i + " " + filler
                    );
                }

                try (FileOutputStream out =
                         new FileOutputStream(file)) {
                    workbook.write(out);
                }
            }

            totalSize += file.length();

            System.out.printf(
                "File %d | Total size: %.2f GiB (%.2f%%)%n",
                fileNumber,
                totalSize / (1024.0 * 1024 * 1024),
                Math.min(100.0, totalSize * 100.0 / TARGET_SIZE)
            );

            fileNumber++;
        }

        System.out.printf(
            "Generation finished. Total Excel size: %.2f GiB%n",
            totalSize / (1024.0 * 1024 * 1024)
        );

        System.out.println("Excel files saved in: " + PATH);
    }
}
