package com.tp.asset_ap.util;

import com.tp.asset_ap.exception.InternalServerErrorException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.UUID;
import java.util.function.Consumer;

public final class FileSystemUtils {

    private static final Logger logger = LoggerFactory.getLogger(FileSystemUtils.class);

    private FileSystemUtils() {
    }

    public static boolean isDir(Path path) {
        ObjectUtils.checkNull(path);
        return path.toFile().isDirectory();
    }

    public static boolean isFile(Path path) {
        ObjectUtils.checkNull(path);
        return path.toFile().isFile();
    }

    public static boolean existsDir(Path path) {
        ObjectUtils.checkNull(path);
        return FileSystemUtils.isDir(path) && path.toFile().exists();
    }

    public static void createDirsIfNotExist(Path path) {
        if (!Files.exists(path)) {
            try {
                Files.createDirectories(path);
            } catch (IOException e) {
                throw new InternalServerErrorException(e.getMessage(), e);
            }
        }
    }

    public static void iterFile(Path path, Consumer<? super File> consumer) {
        if (!path.toFile().isDirectory()) {
            consumer.accept(path.toFile());
            return;
        }

        File[] files = path.toFile().listFiles();

        if (files == null) {
            return;
        }

        for (File file : files) {
            if (file.isDirectory()) {
                iterFile(file.toPath(), consumer);
            } else {
                consumer.accept(file);
            }
        }
    }

    public static void deleteDirContent(File dirPath) {
        if (dirPath == null || !dirPath.exists() || !dirPath.isDirectory()) {
            return;
        }

        File[] files = dirPath.listFiles();
        if (files != null) {
            for (File file : files) {
                if (file.isDirectory()) {
                    deleteDirContent(file);
                }
                try {
                    Files.delete(file.toPath());
                } catch (IOException e) {
                    logger.error(e.getMessage(), e);
                }
            }
        }
    }

    public static void deleteDir(File file) {
        deleteDir(file, true);
    }

    public static void safeDeleteDir(File file) {
        deleteDir(file, false);
    }

    public static void safeDeleteFile(Path path) {
        try {
            Files.delete(path);
        } catch (IOException ioException) {
            logger.error(ioException.getMessage(), ioException);
        }
    }

    public static void safeDeleteFile(File file) {
        try {
            Files.delete(file.toPath());
        } catch (IOException ioException) {
            logger.error(ioException.getMessage(), ioException);
        }
    }

    private static void deleteDir(File file, boolean throwOrNot) {
        if (file == null || !file.exists()) {
            return;
        }

        if (file.isFile()) {
            try {
                Files.delete(file.toPath());
            } catch (IOException e) {
                if (throwOrNot) {
                    throw new InternalServerErrorException(e);
                } else {
                    logger.error(e.getMessage(), e);
                }
            }
        } else {
            File[] listFiles = file.listFiles();
            if (listFiles != null) {
                for (File file2 : listFiles) {
                    safeDeleteDir(file2);
                }
            }
            try {
                Files.delete(file.toPath());
            } catch (IOException e) {
                logger.error(e.getMessage(), e);
            }
        }
    }

    public static void moveFile(File file, String to) throws IOException {
        File dir = new File(to);
        // 如果資料夾不存在 則建立新資料夾
        if (!dir.exists()) {
            Files.createDirectories(dir.toPath());
        }

        Path toPath = PathUtils.getPathAfterCheckFileNameExtension(dir.getAbsolutePath(), file.getName());
        Files.move(file.toPath(), toPath, StandardCopyOption.REPLACE_EXISTING);
    }

    public static long getFreeSpaceInBytes(Path path) {
        return path.toFile().getFreeSpace();
    }

    public static long getTotalSpaceInBytes(Path path) {
        return path.toFile().getTotalSpace();
    }

    public static void hasEnoughSpaceToSave(Path path, long fileSizeInBytes) {
        if ((FileSystemUtils.getFreeSpaceInBytes(path) - fileSizeInBytes <= 0)) {
            throw new InternalServerErrorException("指定路徑儲存空間不足");
        }
    }

    public static Path createTempDirUnderThe(String path) {
        if (TPStringUtils.isNullOrEmpty(path)) {
            throw new InternalServerErrorException("建立暫存資料夾失敗");
        }
        Path tmpPath = PathUtils.getPathAfterCheckFileNameExtension(path,
            UUID.randomUUID().toString()
        );
        if (!new File(tmpPath.toString()).mkdirs()) {
            throw new InternalServerErrorException("建立暫存資料夾失敗");
        }

        return tmpPath;
    }
}

