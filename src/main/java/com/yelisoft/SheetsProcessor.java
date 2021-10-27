package com.yelisoft;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
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

    public static void process(XSSFSheet inSheet,
                               HSSFSheet outSheet,
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

        //Ищем стартовую строку
        for (inSheetStartRow = 6; true; inSheetStartRow++) {
            if (inSheet.getRow(inSheetStartRow) == null) break;
            if (inSheet.getRow(inSheetStartRow).getCell(0) == null) continue;
            if (CellType.NUMERIC.equals(inSheet.getRow(inSheetStartRow).getCell(0).getCellTypeEnum())) {
                if (inSheet.getRow(inSheetStartRow).getCell(0).getNumericCellValue() == 1)
                    break;
            }

        }

        //Найти колонку данных (выдача, приём, консультации)
        String capMonth = reportMonth.substring(0, 1).toUpperCase() + reportMonth.substring(1).toLowerCase();
        for (dataColumn = 8; dataColumn < 100; dataColumn++) {
            XSSFCell cell = inSheet.getRow(inSheetStartRow - 2)
                    .getCell(dataColumn);
            if (null == cell) continue;
            if (capMonth.equalsIgnoreCase(cell.getStringCellValue())) {

                break;
            }
        }
        vydachaColumn = dataColumn + 1;
        if (dataColumn >= 100) {
            log.info("Не найдены колонки данных для месяца {}. Выход из программы", capMonth);
        }

        Map<String, Double[]> outNumbers= new HashMap<>();

        for (int i = inSheetStartRow; i < 2000; i++) {  //2000 - защита от зацикливания
            XSSFRow row = inSheet.getRow(i);
            if (null == row) break;
            XSSFCell cell = row.getCell(inServiceNameColumn);
            if (cell == null) continue;
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
        }

        int outStringNumber = outServiceNameColumn - 1;
        int dataRowNumber;
        for (dataRowNumber = 7; true; dataRowNumber++) {
            HSSFCell stringNumberCell = outSheet.getRow(dataRowNumber).getCell(outStringNumber);
            if (null != stringNumberCell
                    && CellType.STRING.equals(stringNumberCell.getCellTypeEnum())
                    && stringNumberCell.getStringCellValue().startsWith("Общее")) {
                break;
            }

            String outService = outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn).getStringCellValue();
            HSSFCell outCell;
            outCell = outSheet.getRow(dataRowNumber).getCell(totalNumberOfOrdersColumn);
            Double[] outSums = outNumbers.get(outService);
            if (null == outSums) continue;
            outCell.setCellFormula(null);
            outCell.setCellValue(outSums[0]);

            outCell = outSheet.getRow(dataRowNumber).getCell(totalResultsIssuedToApplicants);
            if (config.hasOutputForService(outService)) {
                outCell.setCellFormula(null);
                outCell.setCellValue(outSums[1]);
            }

            if ("".equals(outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn).getStringCellValue())) break;
        }

        int inRow;
        for (inRow = 30; inRow < 2000; inRow++) {
            XSSFCell tmpCell = inSheet.getRow(inRow).getCell(inServiceNameColumn + 1);
            if (tmpCell.getStringCellValue().toLowerCase().startsWith("итого:"))
                break;
        }

        XSSFCell sourceCell = inSheet.getRow(inRow).getCell(dataColumn + 2);
        HSSFCell destinationCell = outSheet.getRow(dataRowNumber).getCell(outColumnOfTotalCell);
        destinationCell.setCellValue(sourceCell.getNumericCellValue());

        log.info("==========Finished for {} -> {}==========", inSheet.getSheetName(), outSheet.getSheetName());
    }

}
