package com.yelisoft;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

public class SheetsProcessor {
    private static final Logger log = LoggerFactory.getLogger(SheetsProcessor.class);
    private static final Set<String> absents = new HashSet<>();

    public static Set<String> getAbsents() {
        return absents;
    }

    public static void process(Sheet inSheet, Sheet outSheet, int outServiceNameColumn, String outFileName) {
        log.info("Started for {} -> {} ----------------------------------------------",
                inSheet.getSheetName(), outFileName + " " + outSheet.getSheetName());

        Config config = Config.getInstance();
        int inSheetStartRow = getStartRow(inSheet);
        String reportMonth = config.getMonth().toLowerCase();
        int dataColumn = getDataColumn(inSheet, inSheetStartRow, reportMonth);
        int vydachaColumn = dataColumn + 1;
        if (dataColumn >= 100) {
            log.error("Не найдены колонки данных для месяца {}. Выход из программы", reportMonth);
        }

        Map<String, Double[]> outNumbers = new HashMap<>();
        Map<String, Double> consultsAccum = new HashMap<>();
        int consultColumn = vydachaColumn + 1;
        int typeOfServiceColumn = 4;
        int inServiceNameColumn = 3;

        for (int rowNumber = inSheetStartRow; rowNumber < 2000; rowNumber++) {  //2000 - защита от зацикливания
            Row row = inSheet.getRow(rowNumber);
            if (null == row)
                break;
            Cell cell = row.getCell(inServiceNameColumn);
            if (cell == null)
                continue;

            String inServiceName = cell.getStringCellValue();
            if ("".equals(inServiceName) || null == inServiceName)
                continue;

            String outService = config.getOutForInService(inServiceName);
            if (null == outService)
                outService = inServiceName;

            Double[] sums = new Double[2];
            sums[0] = row.getCell(dataColumn).getNumericCellValue();
            sums[1] = row.getCell(vydachaColumn).getNumericCellValue();
            Double[] outSums = outNumbers.get(outService);
            if (null == outSums)
                outNumbers.put(outService, sums);
            else {
                // Суммируем Обращения(outSums[0]) и Выдачи(outSums[1])
                outSums[0] += sums[0];
                outSums[1] += sums[1];
            }

            // Суммируем консультации по типу услуг (федеральные, региональные и прочие)
            try {
                String key = row.getCell(typeOfServiceColumn).getStringCellValue().substring(0, 1);
                Double acc = consultsAccum.getOrDefault(key, 0.0);
                acc += row.getCell(consultColumn).getNumericCellValue();
                consultsAccum.put(key, acc);
            } catch (Exception e) {
                e.printStackTrace();
                log.info("Ячейка: {}:{}", row.getRowNum(), typeOfServiceColumn);
                System.exit(1);
            }
        }

        int outStringNumber = outServiceNameColumn - 1;
        int dataRowNumber;
        for (dataRowNumber = 7; true; dataRowNumber++) {
            Cell stringNumberCell = outSheet.getRow(dataRowNumber).getCell(outStringNumber);

            if (null != stringNumberCell
                    && CellType.STRING.equals(stringNumberCell.getCellType())
                    && stringNumberCell.getStringCellValue().startsWith("Общее")) {
                break;
            }

            String outService = outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn).getStringCellValue();
            Cell outCell;
            int totalNumberOfOrdersColumn = 10;
            outCell = outSheet.getRow(dataRowNumber).getCell(totalNumberOfOrdersColumn);
            Double[] outSums = outNumbers.get(outService);
            if (null == outSums)
                continue;
            outCell.setBlank();
            outCell.setCellValue(outSums[0]);
            int totalResultsIssuedToApplicants = 17;
            outCell = outSheet.getRow(dataRowNumber).getCell(totalResultsIssuedToApplicants);
            if (config.hasOutputForService(outService)) {
                outCell.setBlank();
                outCell.setCellValue(outSums[1]);
            }

            if ("".equals(outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn).getStringCellValue())) break;
        }

        // Прописываем количество общее консультаций
        int outColumnOfTotalCell = outServiceNameColumn + 14;
        Cell destinationCell = outSheet.getRow(dataRowNumber).getCell(outColumnOfTotalCell);
        destinationCell.setCellValue(consultsAccum.getOrDefault(getKeyForFilename(outFileName), 0.0));

        log.info("==========Finished for {} -> {}==========", inSheet.getSheetName(), outSheet.getSheetName());
    }

    // Searching start row
    private static int getStartRow(Sheet inSheet) {
        int inSheetStartRow = 6;
        for (; ; inSheetStartRow++) {
            if (inSheet.getRow(inSheetStartRow) == null)
                break;
            if (inSheet.getRow(inSheetStartRow).getCell(0) == null)
                continue;
            if (CellType.NUMERIC.equals(inSheet.getRow(inSheetStartRow).getCell(0).getCellType())) {
                if (inSheet.getRow(inSheetStartRow).getCell(0).getNumericCellValue() == 1)
                    break;
            }
        }
        return inSheetStartRow;
    }

    // Searching dataColumn
    private static int getDataColumn(Sheet inSheet, int inSheetStartRow, String reportMonth) {
        int dataColumn = 8;
        for (; dataColumn < 100; dataColumn++) {
            Cell cell = inSheet.getRow(inSheetStartRow - 2)
                    .getCell(dataColumn);
            if (null == cell) continue;
            if (reportMonth.equalsIgnoreCase(cell.getStringCellValue())) {
                break;
            }
        }
        return dataColumn;
    }

    private static String getKeyForFilename(String fileName) {
        String fileFirstLetter = fileName.substring(0, 1);
        switch (fileFirstLetter) {
            case "f":
                return "Ф";
            case "r":
                return "Р";
            case "o":
                return "И";
        }
        return "";
    }

}
