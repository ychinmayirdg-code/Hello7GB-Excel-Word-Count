
import java.io.File;
import java.util.concurrent.*;
import java.util.concurrent.atomic.AtomicLong;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.openxml4j.util.ZipSecureFile;
public class ReadExcelMultithread {
    static final String PATH = "D:\\Hello50GBFiles";
    static final int THREAD_COUNT = 10;
    static final AtomicLong filesProcessed = new AtomicLong(0);
    static final AtomicLong totalRows = new AtomicLong(0);
    static final AtomicLong helloCount = new AtomicLong(0);
    static final AtomicLong failedFiles = new AtomicLong(0);
    static final AtomicLong[] threadTotals =
            new AtomicLong[THREAD_COUNT];
    static final AtomicLong[] threadFiles =
            new AtomicLong[THREAD_COUNT];
    public static void main(String[] args) throws InterruptedException {
        long startTime = System.nanoTime();
        for (int i = 0; i < THREAD_COUNT; i++) {
            threadTotals[i] = new AtomicLong(0);
            threadFiles[i] = new AtomicLong(0);
        }
        // For trusted, self-generated Excel files only.
        ZipSecureFile.setMinInflateRatio(0.001);
        File folder = new File(PATH);
        if (!folder.isDirectory()) {
            System.out.println("Folder not found: " + PATH);
            return;
        }
        File[] files = folder.listFiles(
                (dir, name) ->
                        name.toLowerCase().endsWith(".xlsx")
        );
        if (files == null) {
            System.out.println("Could not read the folder.");
            return;
        }
        System.out.println("Total Excel files found: " + files.length);
        System.out.println("Number of threads: " + THREAD_COUNT);
        System.out.println("Starting multithreaded reading...\n");
        ThreadFactory threadFactory = new ThreadFactory() {
            private int threadNumber = 1;
            @Override
            public synchronized Thread newThread(Runnable task) {
                return new Thread(task, "Thread " + threadNumber++);
            }
        };
        ThreadPoolExecutor executor = new ThreadPoolExecutor(
                THREAD_COUNT,
                THREAD_COUNT,
                0L,
                TimeUnit.MILLISECONDS,
                new ArrayBlockingQueue<>(THREAD_COUNT * 2),
                threadFactory
        );
        try {
            for (File file : files) {
                while (true) {
                    try {
                        executor.execute(() -> processFile(file));
                        break;
                    } catch (RejectedExecutionException e) {
                        if (executor.isShutdown()) {
                            throw e;
                        }
                        // Wait briefly when the queue is full.
                        Thread.sleep(1);
                    }
                }
            }
        } finally {
            executor.shutdown();
        }
        // Wait for every file-reading task to finish.
        executor.awaitTermination(
                Long.MAX_VALUE, TimeUnit.NANOSECONDS
        );
        // Print final totals only once, after completion.
        System.out.println("\n========== FINAL TOTALS ==========");
        for (int i = 0; i < THREAD_COUNT; i++) {
            System.out.printf(
                    "Thread %d total = %d Hello cells (%d files)%n",
                    i + 1,
                    threadTotals[i].get(),
                    threadFiles[i].get()
            );
        }
        double seconds =
                (System.nanoTime() - startTime) / 1_000_000_000.0;
        System.out.println("\n========== FINAL RESULTS ==========");
        System.out.println("Total Excel files: " + files.length);
        System.out.println("Files processed: " + filesProcessed.get());
        System.out.println("Total rows read: " + totalRows.get());
        System.out.println("Grand total of Hello = " + helloCount.get());
        System.out.println("Failed files: " + failedFiles.get());
        System.out.printf("Execution time: %.2f seconds%n", seconds);
        System.out.println("Multithreaded reading completed.");
    }
    static void processFile(File file) {
        String threadName = Thread.currentThread().getName();
        int threadIndex = Integer.parseInt(
                threadName.substring("Thread ".length())
        ) - 1;
        long rowsInFile = 0;
        long helloInFile = 0;
        try (Workbook workbook = WorkbookFactory.create(file)) {
            for (Sheet sheet : workbook) {
                for (Row row : sheet) {
                    rowsInFile++;
                    // Count "Hello" in column A only.
                    Cell cell = row.getCell(
                            0,
                            Row.MissingCellPolicy.RETURN_BLANK_AS_NULL
                    );
                    if (cell != null
                            && cell.getCellType() == CellType.STRING
                            && "Hello".equals(cell.getStringCellValue())) {
                        helloInFile++;
                    }
                }
            }
            totalRows.addAndGet(rowsInFile);
            helloCount.addAndGet(helloInFile);
            threadTotals[threadIndex].addAndGet(helloInFile);
            threadFiles[threadIndex].incrementAndGet();
            // Print every successful file result; no 100-line limit.
            System.out.printf(
                    "%s: %s -> Count = %d%n",
                    threadName,
                    file.getAbsolutePath(),
                    helloInFile
            );
        } catch (Exception e) {
            failedFiles.incrementAndGet();
            System.out.printf(
                    "%s: %s -> FAILED (%s)%n",
                    threadName,
                    file.getAbsolutePath(),
                    e.getClass().getSimpleName()
            );
        } finally {
            filesProcessed.incrementAndGet();
        }
    }
}