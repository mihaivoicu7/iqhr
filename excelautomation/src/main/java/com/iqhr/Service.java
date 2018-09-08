package com.iqhr;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.Reader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class Service {

    public static final String XLSX = ".xlsx";

    public static final String XLS = "xls";

    public static final String DATE_FORMAT = "dd/MM/yyyy";

    public static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern(DATE_FORMAT);

    public static final Integer MAXIM_CHARS = 255 * 256;

    public static final String CNP = "cnp";

    public static void processPath(Path pathToProcess, String frontNameParam, String endNameParam, List<String> sortColumns) throws IOException {
        Comparator<CSVRecord> comparator = buildComparator(sortColumns);
        if (Files.isDirectory(pathToProcess)) {
            List<Path> filesToProcess = Files.list(pathToProcess).collect(Collectors.toList());
            filesToProcess.stream().filter(path -> FilenameUtils.isExtension(path.getFileName().toString(), XLS))
                    .filter(path -> !path.getFileName().startsWith(frontNameParam))
                    .filter(path -> !FilenameUtils.removeExtension(path.getFileName().toString()).endsWith(endNameParam))
                    .forEach(path -> processXlsSingleFile(path,
                            comparator,
                            frontNameParam, endNameParam));
            filesToProcess.stream().filter(path -> !FilenameUtils.isExtension(path.getFileName().toString(), XLS) && !FilenameUtils.isExtension(path.getFileName().toString(),
                    XLSX))
                    .filter(path -> !path.getFileName().startsWith(frontNameParam))
                    .filter(path -> !FilenameUtils.removeExtension(path.getFileName().toString()).endsWith(endNameParam))
                    .parallel().forEach(path -> processDiffXlsFile(path,
                    frontNameParam, endNameParam));
        } else {
            processXlsSingleFile(pathToProcess, comparator, frontNameParam, endNameParam);
        }
    }

    private static Comparator<CSVRecord> buildComparator(List<String> sortColumns) {
        Comparator<CSVRecord> comparator = new Comparator<CSVRecord>() {
            @Override
            public int compare(CSVRecord o1, CSVRecord o2) {
                for (String column : sortColumns) {
                    if (o1.isMapped(column) && o2.isMapped(column)) {
                        int compResult = o1.get(column).compareToIgnoreCase(o2.get(column));
                        if (compResult != 0) {
                            return compResult;
                        }
                    }
                }
                return 0;
            }
        };
        return comparator;
    }

    private static void processDiffXlsFile(Path pathToFile, String frontNameParam, String endNameParam) {
        try {
            String newFileName = frontNameParam + FilenameUtils.removeExtension(pathToFile.getFileName().toString()) + "_" + endNameParam + "." + FilenameUtils.getExtension
                    (pathToFile.toString());
            Files.deleteIfExists(pathToFile.getParent().resolve(newFileName));
            Files.copy(pathToFile, pathToFile.getParent().resolve(newFileName));
            System.out.println("Processing finished for: " + pathToFile.toString() + " Created file: " + newFileName);
        } catch (Exception e) {
            System.out.println("Can't process file: " + pathToFile.toString());
        }
    }

    private static void processXlsSingleFile(Path pathToFile, Comparator<CSVRecord> comparator, String frontNameParam, String endNameParam) {
        try (Workbook workbook = new XSSFWorkbook();) {
            System.out.println("Started processing for: " + pathToFile.toString());
            Reader reader = Files.newBufferedReader(pathToFile);
            CSVParser parse = CSVFormat.newFormat('\t')
                    .withFirstRecordAsHeader()
                    .withIgnoreHeaderCase()
                    .withTrim()
                    .parse(reader);
            List<CSVRecord> records = parse.getRecords();
            records.sort(comparator);
            String sheetName = frontNameParam + FilenameUtils.removeExtension(pathToFile.getFileName().toString()) + "_" + endNameParam;
            String newFileName = sheetName + XLSX;
            Path newFilePath = pathToFile.getParent().resolve(newFileName);
            CreationHelper createHelper = workbook.getCreationHelper();
            CellStyle dateStyle = workbook.createCellStyle();
            short dateFormat = createHelper.createDataFormat().getFormat(DATE_FORMAT);
            dateStyle.setDataFormat(dateFormat);
            // Create a Sheet
            Sheet sheet = workbook.createSheet(sheetName);
            List<String> headers = new ArrayList<>();
            Map<String, Integer> headersMap = parse.getHeaderMap();
            headers.addAll(parse.getHeaderMap().keySet());
            headers.sort(new Comparator<String>() {
                @Override
                public int compare(String o1, String o2) {
                    return headersMap.get(o1).compareTo(headersMap.get(o2));
                }
            });
            Map<Integer, Integer> sizeMap = new HashMap<>();
            IntStream.range(0, headers.size()).forEach(i -> sizeMap.put(i, headers.get(i).length()));
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers.get(i));
            }

            int rowNum = 1;
            Map<String, ColumnType> columnTypeMap = new HashMap<>();
            for (CSVRecord record : records) {
                Row row = sheet.createRow(rowNum);
                for (int i = 0; i < headers.size(); i++) {

                    String initialCellValue = record.get(headers.get(i));
                    if (sizeMap.get(i) < initialCellValue.length()) {
                        sizeMap.replace(i, initialCellValue.length());
                    }
                    createCell(row, initialCellValue, i, headers.get(i), dateStyle, columnTypeMap);
                }
                rowNum = rowNum + 1;
            }
            sizeMap.entrySet().forEach(entry -> sheet.setColumnWidth(entry.getKey(), (entry.getValue() + 2) * 256 > MAXIM_CHARS ? MAXIM_CHARS : (entry.getValue() + 2) * 256));
            sheet.createFreezePane(4, 1);
            ((XSSFWorkbook) workbook).lockStructure();
            // Write the output to a file
            Files.deleteIfExists(newFilePath);
            try (FileOutputStream fileOut = new FileOutputStream(newFilePath.toFile());) {
                workbook.write(fileOut);
            }
            System.out.println("Processing finished for: " + pathToFile.toString() + " Created file: " + newFilePath.getFileName().toString());
        } catch (Exception e) {
            System.out.println("Could not process: " + pathToFile);
            System.out.println("Error message: " + e);
            e.printStackTrace();
        }
    }

    private static void createCell(Row row, String cellValue, int potision, String colName, CellStyle dateStyle, Map<String, ColumnType> columnTypeMap) {
        SimpleDateFormat sdf = new SimpleDateFormat(DATE_FORMAT);
        Cell recordCell = row.createCell(potision);
        if (colName.equalsIgnoreCase(CNP)) {
            CellType cellType = CellType.STRING;
            recordCell.setCellValue(cellValue);
            recordCell.setCellType(cellType);
        } else {
            ColumnType columnType = columnTypeMap.get(colName);
            if(columnType == null) {
                columnType = new ColumnType();
                columnType.setColumnName(colName);
                columnTypeMap.put(colName, columnType);
            }
            if (columnTypeMap.get(colName).getTrials() < 10) {
                if (NumberUtils.isParsable(cellValue)) {
                    columnType.setCellType(CellTypeCustom.NUMERIC);
                    columnType.setTrials(columnType.getTrials()+1);
                    CellType cellType = CellType.NUMERIC;
                    recordCell.setCellValue(Double.valueOf(cellValue));
                    recordCell.setCellType(cellType);
                } else {
                    try {

                        Date date = sdf.parse(cellValue);
                        recordCell.setCellStyle(dateStyle);
                        recordCell.setCellValue(date);
                        columnType.setCellType(CellTypeCustom.DATE);
                        columnType.setTrials(columnType.getTrials() + 1);
                    } catch (ParseException dateFormatException) {
                        if(StringUtils.isNotEmpty(cellValue)) {
                            columnType.setCellType(CellTypeCustom.STRING);
                            columnType.setTrials(columnType.getTrials() + 1);
                        }
                        CellType cellType = CellType.STRING;
                        recordCell.setCellValue(cellValue);
                        recordCell.setCellType(cellType);
                    }
                }
            } else {
                recordCell.setCellStyle(columnType.getCellStyle());
                switch (columnType.getCellType()) {
                    case NUMERIC:
                        if(StringUtils.isNotEmpty(cellValue)) {
                            recordCell.setCellValue(Double.valueOf(cellValue));
                            recordCell.setCellType(CellType.NUMERIC);
                        }
                        break;
                    case DATE:
                        try {
                            recordCell.setCellValue(sdf.parse(cellValue));
                            recordCell.setCellStyle(dateStyle);
                        } catch (ParseException e) {
                            recordCell.setCellValue(cellValue);
                        }
                        break;
                    case STRING:
                        recordCell.setCellValue(cellValue);
                        recordCell.setCellType(CellType.STRING);
                        break;
                }
            }
        }
    }

}
