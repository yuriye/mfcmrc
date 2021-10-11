package com.yelisoft;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

public class SheetsProcessor {

    private static final Logger log = LoggerFactory.getLogger(SheetsProcessor.class);
    private static Set<String> absents = new HashSet<>();

    public static Set<String> getAbsents() {
        return absents;
    }

    public static void process(XSSFSheet inSheet,
                               HSSFSheet outSheet,
                               int outServiceNameColumn,
                               String auth, String outFileName) throws IOException {

        log.info("Started for {} -> {} ----------------------------------------------",
                inSheet.getSheetName(), outFileName + " " + outSheet.getSheetName());

        Config config = Config.getInstance();
        Map<String, Integer> inServiceRow = new HashMap<>();

        int inServiceNameColumn = 3;
        int inSheetStartRow = 7;
        int dataColumn;
        int vydachaColumn = 0;
        int outColumnOfTotalCell = outServiceNameColumn + 14;
        String reportMonth = config.getMonth().toLowerCase();

        int totalNumberOfOrdersColumn = 10;
        int totalResultsIssuedToApplicants = 17;

        int numberOfOrdersFormedForRosreestr = 6;
        int numberOfClosedOrdersForRosreestr = 11;

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

        int sumRosreestrUchTotal = 0;
        int sumRosreestrUchExterrTotal = 0;
        int sumRosreestrUchClosed = 0;
        int sumRosreestrUchExterrClosed = 0;
        int sumRosreestrSvedTotal = 0;
        int sumRosreestrSvedClosed = 0;

        for (int i = inSheetStartRow; i < 2000; i++) {  //2000 - защита от зацикливания
            XSSFRow row = inSheet.getRow(i);
            if (null == row) break;
            XSSFCell cell = row.getCell(inServiceNameColumn);
            if (cell == null) continue;
            String inServiceName = cell.getStringCellValue();
            if ("".equals(inServiceName) || null == inServiceName) continue;
            inServiceRow.put(inServiceName, i);

            if ("(КАМЧАТСКИЙ КРАЙ) Государственная услуга по государственному кадастровому учету недвижимого имущества и (или) государственной регистрации прав на недвижимое имущество и сделок с ним"
                    .equals(inServiceName) ||
                    "(ЭКСТЕР) Государственная услуга по государственному кадастровому учету недвижимого имущества и (или) государственной регистрации прав на недвижимое имущество и сделок с ним"
                            .equals(inServiceName)) {

                sumRosreestrUchTotal += row.getCell(dataColumn).getNumericCellValue();
                sumRosreestrUchClosed += row.getCell(vydachaColumn).getNumericCellValue();
                if (inServiceName.startsWith("(Э")) {
                    sumRosreestrUchExterrTotal += row.getCell(dataColumn).getNumericCellValue();
                    ;
                    sumRosreestrUchExterrClosed += row.getCell(vydachaColumn).getNumericCellValue();
                }

            } else if ("Государственная услуга по предоставлению сведений, cодержащихся в Едином государственном реестре недвижимости (ЕГРН)"
                    .equals(inServiceName)) {

                sumRosreestrSvedTotal += row.getCell(dataColumn).getNumericCellValue();
                sumRosreestrSvedClosed += row.getCell(vydachaColumn).getNumericCellValue();
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

            if ("Государственный кадастровый учет недвижимого имущества и (или) государственная регистрация прав на недвижимое имущество"
                    .equals(outService)) {

                outSheet.getRow(dataRowNumber).getCell(numberOfOrdersFormedForRosreestr).setCellValue(sumRosreestrUchTotal);
                outSheet.getRow(dataRowNumber).getCell(numberOfOrdersFormedForRosreestr + 1).setCellValue(sumRosreestrUchExterrTotal);
                outSheet.getRow(dataRowNumber).getCell(numberOfClosedOrdersForRosreestr).setCellValue(sumRosreestrUchClosed);
                outSheet.getRow(dataRowNumber).getCell(numberOfClosedOrdersForRosreestr + 1).setCellValue(sumRosreestrUchExterrClosed);
                continue;
            }

            if ("Предоставление сведений, содержащихся в Едином государственном реестре недвижимости"
                    .equals(outService)) {
                outSheet.getRow(dataRowNumber).getCell(numberOfOrdersFormedForRosreestr).setCellValue(sumRosreestrSvedTotal);
                outSheet.getRow(dataRowNumber).getCell(numberOfClosedOrdersForRosreestr).setCellValue(sumRosreestrSvedClosed);
                continue;
            }

            String inService = config.getInForOutService(outService);
            if (null == inService) {
                absents.add(outService);
                continue;
            } else {
                HSSFCell outCell = outSheet.getRow(dataRowNumber).getCell(totalNumberOfOrdersColumn);
                outCell.setCellFormula(null);
                outCell.setCellValue(0);
                outCell = outSheet.getRow(dataRowNumber).getCell(totalResultsIssuedToApplicants);
                outCell.setCellFormula(null);
                outCell.setCellValue(0);
            }

            if (null == inServiceRow.get(inService))
                continue;
            XSSFRow inRow = inSheet.getRow(inServiceRow.get(inService));
            HSSFCell outCell;
            XSSFCell inCell;
            outCell = outSheet.getRow(dataRowNumber).getCell(totalNumberOfOrdersColumn);
            outCell.setCellFormula(null);

            inCell = inRow.getCell(dataColumn);
            double cellValue = 0;
            if (null != inCell) cellValue = inCell.getNumericCellValue();

            outCell.setCellValue(cellValue);
            addToCellDoubleValue(outCell, cellValue);

            if (config.hasOutputForService(outService)) {
                outCell = outSheet.getRow(dataRowNumber).getCell(totalResultsIssuedToApplicants);
                outCell.setCellFormula(null);
                inCell = inRow.getCell(vydachaColumn);
                cellValue = 0;
                if (null != inCell) cellValue = inCell.getNumericCellValue();
                addToCellDoubleValue(outCell, cellValue);
            } else {
                outSheet.getRow(dataRowNumber).getCell(totalResultsIssuedToApplicants).setCellValue("нет выдачи через МФЦ");
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

    static void addToCellDoubleValue(HSSFCell cell, double number) {
        double value = number;
        if (CellType.STRING == cell.getCellTypeEnum()) {
            cell.setCellFormula(null);
        } else {
            value += cell.getNumericCellValue();
        }
        cell.setCellValue(value);
    }


}
