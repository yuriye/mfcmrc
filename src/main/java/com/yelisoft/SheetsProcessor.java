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

    public static void process(Sheet inSheet,
                               Sheet outSheet,
                               int outServiceNameColumn,
                               String outFileName) {

        log.info("Started for {} -> {} ----------------------------------------------",
                inSheet.getSheetName(), outFileName + " " + outSheet.getSheetName());

        Config config = Config.getInstance();

        int inServiceNameColumn = 3;
        int inSheetStartRow;
        int dataColumn;
        int vydachaColumn;
        int outColumnOfTotalCell = outServiceNameColumn + 14;
        String reportMonth = config.getMonth().toLowerCase();

        int totalNumberOfOrdersColumn = 10;
        int totalResultsIssuedToApplicants = 17;

        Map<String, Double> consultsAccum = new HashMap<>();

        //Ищем стартовую строку
        for (inSheetStartRow = 6; true; inSheetStartRow++) {
            if (inSheet.getRow(inSheetStartRow) == null) break;
            if (inSheet.getRow(inSheetStartRow).getCell(0) == null) continue;
            if (CellType.NUMERIC.equals(inSheet.getRow(inSheetStartRow).getCell(0).getCellType())) {
                if (inSheet.getRow(inSheetStartRow).getCell(0).getNumericCellValue() == 1)
                    break;
            }

        }

        //Найти колонку данных (выдача, приём, консультации)
        String capMonth = reportMonth.substring(0, 1).toUpperCase() + reportMonth.substring(1).toLowerCase();
        for (dataColumn = 8; dataColumn < 100; dataColumn++) {
            Cell cell = inSheet.getRow(inSheetStartRow - 2)
                    .getCell(dataColumn);
            if (null == cell) continue;
            if (capMonth.equalsIgnoreCase(cell.getStringCellValue())) {

                break;
            }
        }

        vydachaColumn = dataColumn + 1;
        if (dataColumn >= 100) {
            log.error("Не найдены колонки данных для месяца {}. Выход из программы", capMonth);
        }

        Map<String, Double[]> outNumbers = new HashMap<>();

        int consultColumn = vydachaColumn + 1;
        int typeOfServiceColumn = 4;

        for (int i = inSheetStartRow; i < 2000; i++) {  //2000 - защита от зацикливания
            Row row = inSheet.getRow(i);

            if (null == row)
                break;
            Cell cell = row.getCell(inServiceNameColumn);

            if (cell == null)
                continue;

            String inServiceName = cell.getStringCellValue();
            if ("".equals(inServiceName) || null == inServiceName) continue;

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
                outSums[0] += sums[0];
                outSums[1] += sums[1];
            }

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
            outCell = outSheet.getRow(dataRowNumber).getCell(totalNumberOfOrdersColumn);
            Double[] outSums = outNumbers.get(outService);
            if (null == outSums)
                continue;
            outCell.setBlank();
            outCell.setCellValue(outSums[0]);

            outCell = outSheet.getRow(dataRowNumber).getCell(totalResultsIssuedToApplicants);
            if (config.hasOutputForService(outService)) {
                outCell.setBlank();
                outCell.setCellValue(outSums[1]);
            }

            if ("".equals(outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn).getStringCellValue())) break;
        }

        int inRow;
        for (inRow = 30; inRow < 2000; inRow++) {

            Cell tmpCell = null;
            try {
                tmpCell = inSheet.getRow(inRow).getCell(inServiceNameColumn + 1);
                if (tmpCell.getCellType() != CellType.STRING) {
                    continue;
                }
            } catch (Exception e) {
                e.printStackTrace();
                log.error(e.getMessage());
                log.error("inRow = {}; inColumn = {}", inRow, inServiceNameColumn + 1);
                log.error(inSheet.getRow(inRow - 1).getCell(inServiceNameColumn + 1).getStringCellValue());
                System.exit(1);
            }

            String rawVal;
            rawVal = tmpCell.getStringCellValue();
            if (rawVal.toLowerCase().startsWith("итого:"))
                break;
        }

        Cell destinationCell = outSheet.getRow(dataRowNumber).getCell(outColumnOfTotalCell);

        String key = "";
        String fileFirstLetter = outFileName.substring(0, 1);
        switch (fileFirstLetter) {
            case "f":
                key = "Ф";
                break;
            case "r":
                key = "Р";
                break;
            case "o":
                key = "И";
                break;
        }

        destinationCell.setCellValue(consultsAccum.getOrDefault(key, 0.0));

        log.info("==========Finished for {} -> {}==========", inSheet.getSheetName(), outSheet.getSheetName());
    }

}
