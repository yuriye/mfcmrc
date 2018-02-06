package com.yelisoft;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class SheetsProcessor {

    private static final Logger log = LoggerFactory.getLogger(SheetsProcessor.class);

    public static Map<String, Boolean> hasOutputMap = new HashMap<>();

    public static void process(XSSFSheet inSheet,
                               HSSFSheet outSheet,
                               int outServiceNameColumn,
                               String auth) throws IOException {


        Config config = Config.getInstance();
        Map<String, Integer> inServiceRow = new HashMap<>();

        int inServiceNameColumn = 1;
        int inSheetStartRow = 10;
        int dataColumn;
        int vydachaColumn = 19;
        int outRowOfTotalCell = 0;
        int outColumnOfTotalCell = outServiceNameColumn + 5;
        int consultRowNumber = 0;

        int totalNumberOfOrdersColumn = 9;
        int totalResultsIssuedToApplicants = 15;

        int numberOfOrdersFormedForRosreestr = 6;
        int numberOfClosedOrdersForRosreestr = 10;


        //ВЫДАЧА ЗА ...
        String vydZa = "ВЫДАЧА ЗА " + config.getMonth().toUpperCase();
        for (dataColumn = 7; dataColumn < 100; dataColumn++) {
            if (vydZa.equals(inSheet.getRow(inSheetStartRow - 2)
                    .getCell(dataColumn)
                    .getStringCellValue()
                    .toUpperCase())) {
                vydachaColumn = dataColumn;
//                System.out.println("vydachaColumn=" + vydachaColumn);
                log.info("vydachaColumn=" + vydachaColumn);
                break;
            }

        }

        for (dataColumn = 7; dataColumn < vydachaColumn; dataColumn++) {
            if (config.getMonth().toUpperCase().equals(
                    inSheet.getRow(inSheetStartRow - 2).getCell(dataColumn).getStringCellValue().toUpperCase()))
                break;
        }

        int sumOf3cell = 0;
        int sumVydachOf3cell = 0;
        for (int i = inSheetStartRow; true; i++) {
            XSSFRow row = inSheet.getRow(i);
            if (null == row) break;
            try {

                if ( null != row.getCell(0) && CellType.STRING.equals(row.getCell(0).getCellTypeEnum())) {
                    String cell1Value = row.getCell(0).getStringCellValue();
                    if ("Прием запросов на регистрацию на портале Gosuslugi.ru".equals(cell1Value)
                            || "Прием запросов на подтверждение регистрации на портале Gosuslugi.ru".equals(cell1Value)
                            || "Восстановление регистрации на портале Gosuslugi.ru".equals(cell1Value)) {
                        sumOf3cell += null == row.getCell(dataColumn)? 0: row.getCell(dataColumn).getNumericCellValue();
                        sumVydachOf3cell += null == row.getCell(vydachaColumn)? 0: row.getCell(vydachaColumn).getNumericCellValue();
                    } else if ("Консультации".equals(cell1Value))
                        consultRowNumber = row.getRowNum();
                }

            } catch (Exception e) {
                log.error("inSheet scaning row==" + row, e);
                e.printStackTrace();
            }

            XSSFCell cell = row.getCell(inServiceNameColumn);
            if (cell == null) continue;
            String inServiceName = cell.getStringCellValue();
            if ("".equals(inServiceName) || null == inServiceName) continue;
            inServiceRow.put(inServiceName, i);
        }

//        System.out.println(inSheet.getSheetName() + "->" + outSheet.getSheetName());
        log.info(inSheet.getSheetName() + "->" + outSheet.getSheetName());
        for (int dataRowNumber = 7; true; dataRowNumber++) {

            if (CellType.STRING.equals(outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn - 1).getCellTypeEnum())) {
                if (outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn - 1).getStringCellValue().startsWith("Общее")) {
                    outRowOfTotalCell = dataRowNumber;
                    break;
                }
            }

            try {
                HSSFRow row = outSheet.getRow(dataRowNumber);
                if (null == row) break;
                HSSFCell cell = row.getCell(outServiceNameColumn - 1);
                if (cell == null) continue;
//                int rowNumber = Integer.valueOf(cell.getStringCellValue());
            } catch (NumberFormatException mfe) {
                continue;
            }
            String outService = outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn).getStringCellValue();
            if ("Регистрация, подтверждение личности, восстановление доступа граждан в Единой системе идентификации и аутентификации (ЕСИА)"
                    .equals(outService)) {
                outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn + 2).setCellValue(sumOf3cell);
//                outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn + 3).setCellValue(sumVydachOf3cell);
//                System.out.println("sumOf3cell = " + sumOf3cell);
//                System.out.println("sumVydachOf3cell = " + sumVydachOf3cell);
                log.info("inSheet.getSheetName() == " + inSheet.getSheetName() + ", outSheet.getSheetName() == " + outSheet.getSheetName()
                        + ": sumOf3cell == " + sumOf3cell + "\t" + "sumVydachOf3cell == " + sumVydachOf3cell);
                continue;
            }

            int rosreestrRowNumber1 = consultRowNumber + 5;
            int rosreestrRowNumber2 = consultRowNumber + 6;
            int rosreestrColumn1 = 0;
            log.info(inSheet.getSheetName() + " rosreestrRowNumber1 == " + rosreestrRowNumber1);
            for(int i = 0; i < 100; i++) {
                XSSFCell xcell = inSheet.getRow(rosreestrRowNumber1).getCell(i);
                if (null == xcell) {
                    log.info("i == " + i + " null == xcell");
                    continue;
                }
                if (CellType.STRING.equals(xcell.getCellTypeEnum()) &&
                        xcell.getStringCellValue().startsWith("Гос")) {
                    rosreestrColumn1 = i;
                    break;
                }
            }
            log.info("rosreestrColumn1 == " + rosreestrColumn1);
            if(rosreestrColumn1 == 0) {
                log.error("Не нашли ячейку для росреестра");
            }

            if ("Государственный кадастровый учет и (или) государственная регистрация прав на недвижимое имущество"
                    .equals(outService)) {
                copyXToHCell(inSheet.getRow(rosreestrRowNumber1).getCell(rosreestrColumn1),
                        outSheet.getRow(dataRowNumber).getCell(numberOfOrdersFormedForRosreestr));
                copyXToHCell(inSheet.getRow(rosreestrRowNumber1).getCell(rosreestrColumn1 + 1),
                        outSheet.getRow(dataRowNumber).getCell(numberOfClosedOrdersForRosreestr));
            }
            if ("Предоставление сведений, содержащихся в Едином государственном реестре недвижимости"
                    .equals(outService)) {
                copyXToHCell(inSheet.getRow(rosreestrRowNumber2).getCell(rosreestrColumn1),
                        outSheet.getRow(dataRowNumber).getCell(numberOfOrdersFormedForRosreestr));
                copyXToHCell(inSheet.getRow(rosreestrRowNumber2).getCell(rosreestrColumn1 + 1),
                        outSheet.getRow(dataRowNumber).getCell(numberOfClosedOrdersForRosreestr));
            }

            String inService;

            if ("Установление ежемесячной денежной выплаты отдельным категориям граждан в Российской Федерации"
                    .equals(outService)) {
                int sumOfCell = 0;
                int sumVydachOfCells = 0;

                inService = "Прием заявлений о предоставлении набора социальных услуг, об отказе от получения набора социальных услуг или о возобновлении предоставления набора социальных услуг (Установление ЕДВ)";
                XSSFRow inRow = inSheet.getRow(inServiceRow.get(inService));
                XSSFCell inDataCell = inRow.getCell(dataColumn);
                if (null != inServiceRow.get(inService))
                    if (null != inDataCell)
                        if (CellType.NUMERIC.equals(inDataCell.getCellTypeEnum())) {
                            sumOfCell += inDataCell.getNumericCellValue();
                            sumVydachOfCells += inRow.getCell(vydachaColumn).getNumericCellValue();
                        }
                inService = "Доставка ежемесячной денежной выплаты (Установление ЕДВ)";
                inRow = inSheet.getRow(inServiceRow.get(inService));
                inDataCell = inRow.getCell(dataColumn);
                if (null != inServiceRow.get(inService))
                    if (null != inDataCell)
                        if (CellType.NUMERIC.equals(inDataCell.getCellTypeEnum())) {
                            sumOfCell += inDataCell.getNumericCellValue();
                            sumVydachOfCells += inRow.getCell(vydachaColumn).getNumericCellValue();
                        }

                outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn + 2).setCellValue(sumOfCell);
//                outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn + 3).setCellValue(sumVydachOfCells);
                continue;
            }

            inService = config.getInForOutService(outService);
            if (null == inService) continue;

//            System.out.println("inService:" + inService );
//            System.out.println("row=" + inServiceRow.get(inService));
            if (null == inServiceRow.get(inService)) continue;
            XSSFRow inRow = inSheet.getRow(inServiceRow.get(inService));
            XSSFCell inDataCell = inRow.getCell(dataColumn);
            if (null == inDataCell) continue;

            copyXToHCell(inDataCell, outSheet.getRow(dataRowNumber).getCell( totalNumberOfOrdersColumn));

            if (config.hasOutputForService(outService)) {
                copyXToHCell(inRow.getCell(vydachaColumn),
                        outSheet.getRow(dataRowNumber).getCell(totalResultsIssuedToApplicants));
                hasOutputMap.put(outService, true);
            }
            else {
                hasOutputMap.put(outService, false);
            }
            if ("".equals(outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn).getStringCellValue())) break;
        }

        int rowNum;
        for (rowNum = inSheet.getLastRowNum(); rowNum >= 0; rowNum--) {
//            System.out.println("rowNum=" + rowNum);
            if (null == inSheet.getRow(rowNum)) continue;
            if (null == inSheet.getRow(rowNum).getCell(0)) continue;
            if (inSheet.getRow(rowNum).getCell(0).getStringCellValue().toUpperCase().startsWith("КОНС")) break;
            if (inSheet.getRow(rowNum).getCell(0).getStringCellValue().length() > 60) break;
        }
        if ("fed".equals(auth)) {
            copyXToHCell(inSheet.getRow(rowNum).getCell(16),
                    outSheet.getRow(outRowOfTotalCell).getCell(outColumnOfTotalCell));
        } else if ("reg".equals(auth)) {
            copyXToHCell(inSheet.getRow(rowNum).getCell(17),
                    outSheet.getRow(outRowOfTotalCell).getCell(outColumnOfTotalCell));
        } else if ("oth".equals(auth)) {
            copyXToHCell(inSheet.getRow(rowNum).getCell(19),
                    outSheet.getRow(outRowOfTotalCell).getCell(outColumnOfTotalCell));
        }

    }

    private static void copyXToHCell(XSSFCell sourceCell, HSSFCell destinationCell) {
        if (CellType.BLANK.equals(sourceCell.getCellTypeEnum())) {
        } else if (CellType.NUMERIC.equals(sourceCell.getCellTypeEnum())) {
            destinationCell.setCellValue(sourceCell.getNumericCellValue());
        } else if (CellType.STRING.equals(sourceCell.getCellTypeEnum())) {
            destinationCell.setCellValue(sourceCell.getStringCellValue());
        } else if (CellType.BOOLEAN.equals(sourceCell.getCellTypeEnum())) {
            destinationCell.setCellValue(sourceCell.getBooleanCellValue());
        }
    }

}
