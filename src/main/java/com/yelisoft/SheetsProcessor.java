package com.yelisoft;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
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

        log.info("Started for {} -> {} ----------------------------------------------",
                inSheet.getSheetName(), outSheet.getSheetName());

        Config config = Config.getInstance();
        Map<String, Integer> inServiceRow = new HashMap<>();

        int inServiceNameColumn = 1;
        int inSheetStartRow = 10;
        int dataColumn;
        int vydachaColumn = 19;
        int outRowOfTotalCell = 0;
        int outColumnOfTotalCell = outServiceNameColumn + 12;
        int consultRowNumber = 0;

        int totalNumberOfOrdersColumn = 9;
        int totalResultsIssuedToApplicants = 15;

        int numberOfOrdersFormedForRosreestr = 6;
        int numberOfClosedOrdersForRosreestr = 10;

        //Ищем стартовую строу
        for (inSheetStartRow = 6; true; inSheetStartRow++) {
            if (inSheet.getRow(inSheetStartRow) == null) break;
            if (inSheet.getRow(inSheetStartRow).getCell(0) == null) continue;
            if(CellType.NUMERIC.equals(inSheet.getRow(inSheetStartRow).getCell(0).getCellTypeEnum())) {
                if (inSheet.getRow(inSheetStartRow).getCell(0).getNumericCellValue() == 1)
                    break;
            }
        }

        //ВЫДАЧА ЗА ...
        String vydZa = "ВЫДАЧА ЗА " + config.getMonth().toUpperCase();
        for (dataColumn = 7; dataColumn < 100; dataColumn++) {
            XSSFCell cell = inSheet.getRow(inSheetStartRow - 2)
                    .getCell(dataColumn);
            if (null == cell) continue;
            if (vydZa.equals(cell.getStringCellValue().toUpperCase())) {
                vydachaColumn = dataColumn;
                break;
            }
        }
        //Поиск номера колонки за отчетный месяц
        for (dataColumn = 7; dataColumn < vydachaColumn; dataColumn++) {
            if (config.getMonth().toUpperCase().equals(
                    inSheet.getRow(inSheetStartRow - 2).getCell(dataColumn).getStringCellValue().toUpperCase()))
                break;
        }
        log.info("dataColumn == " + dataColumn + ", vydachaColumn == " + vydachaColumn);

        int sumOf3cell = 0;
        int sumVydachOf3cell = 0;
        for (int i = inSheetStartRow; true; i++) {
            XSSFRow row = inSheet.getRow(i);
            if (null == row) break;
            try {
                if (null != row.getCell(0) && CellType.STRING.equals(row.getCell(0).getCellTypeEnum())) {
                    String cell1Value = row.getCell(0).getStringCellValue();
                    cell1Value = cell1Value == null ? "" : cell1Value;
                    if ("Прием запросов на регистрацию на портале Gosuslugi.ru".equals(cell1Value)
                            || "Прием запросов на подтверждение регистрации на портале Gosuslugi.ru".equals(cell1Value)
                            || "Восстановление регистрации на портале Gosuslugi.ru".equals(cell1Value)) {
                        sumOf3cell += null == row.getCell(dataColumn)? 0: row.getCell(dataColumn).getNumericCellValue();
                        sumVydachOf3cell += null == row.getCell(vydachaColumn)? 0: row.getCell(vydachaColumn).getNumericCellValue();
                    }
//                    else if ("КОНСУЛЬТАЦИИ".equals(cell1Value.toUpperCase()))
//                        consultRowNumber = row.getRowNum();
                }
            } catch (Exception e) {
                log.error("inSheet scaning row== {} {}",row, e);
                e.printStackTrace();
            }

            // consultRowNumber
            for (consultRowNumber = inSheet.getLastRowNum(); ; consultRowNumber--) {
                if (null == inSheet.getRow(consultRowNumber)) continue;
                XSSFCell cell = inSheet.getRow(consultRowNumber).getCell(0);
                if (null == cell) continue;
                if (!CellType.STRING.equals(cell.getCellTypeEnum())) continue;
                if (cell.getStringCellValue().toUpperCase().startsWith("КОНСУЛЬТ")
                        && cell.getStringCellValue().length() < 13)
                    break;
            }



            XSSFCell cell = row.getCell(inServiceNameColumn);
            if (cell == null) continue;
            String inServiceName = cell.getStringCellValue();
            if ("".equals(inServiceName) || null == inServiceName) continue;
            inServiceRow.put(inServiceName, i);
        }
        log.info("sumOf3cell == {}  sumVydachOf3cell == {}  consultRowNumber == {}", sumOf3cell, sumVydachOf3cell, consultRowNumber);

        //Вычисление ячеек "В ПК ПВД"
        int rosreestrColumn1 = 0;
        int rosreestrRowNumber1;
        rr:
        for(rosreestrRowNumber1 = consultRowNumber + 1; rosreestrRowNumber1 < 1000; rosreestrRowNumber1++) {
            for (int column = 0; column < vydachaColumn; column++) {
                XSSFCell cell = null;
                try {
                    cell = inSheet.getRow(rosreestrRowNumber1).getCell(column);
                } catch (Exception e) {}
                if(cell == null) continue;
                if (CellType.STRING.equals(cell.getCellTypeEnum()))
                    if(cell.getStringCellValue().startsWith("Гос")) {
                        rosreestrColumn1 = column;
                        break rr;
                    }
            }
        }
        int rosreestrRowNumber2 = rosreestrRowNumber1 + 1;
        log.info("{}: rosreestrColumn1 == {}  rosreestrRowNumber1 == {}  rosreestrRowNumber2 == {}", inSheet.getSheetName(), rosreestrColumn1, rosreestrRowNumber1, rosreestrRowNumber2);

        for(int i = 0; i < 100; i++) {
            XSSFCell xcell = inSheet.getRow(rosreestrRowNumber1).getCell(i);
            if (null == xcell) {
                continue;
            }
            if (CellType.STRING.equals(xcell.getCellTypeEnum()) &&
                    xcell.getStringCellValue().startsWith("Гос")) {
                rosreestrColumn1 = i;
                break;
            }
        }

        if(rosreestrColumn1 == 0) {
            log.error("Не нашли ячейку для росреестра");
        }

        rosreestrColumn1++;


        for (int dataRowNumber = 7; true; dataRowNumber++) {
            if(null != outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn - 1)
                    && CellType.STRING.equals(outSheet.getRow(dataRowNumber)
                            .getCell(outServiceNameColumn - 1).getCellTypeEnum())) {
                if (outSheet.getRow(dataRowNumber)
                        .getCell(outServiceNameColumn - 1)
                        .getStringCellValue().startsWith("Общее")) {
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
            log.info("Для outService dataRowNumber == {} outServiceNameColumn == {}", dataRowNumber, outServiceNameColumn);
            log.info(outService);
            if ("Регистрация, подтверждение личности, восстановление доступа граждан в Единой системе идентификации и аутентификации (ЕСИА)"
                    .equals(outService)) {
                outSheet.getRow(dataRowNumber).getCell(totalNumberOfOrdersColumn).setCellFormula(null);
                outSheet.getRow(dataRowNumber).getCell(totalNumberOfOrdersColumn).setCellValue(sumOf3cell);
                continue;
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

                outSheet.getRow(dataRowNumber).getCell(totalNumberOfOrdersColumn).setCellFormula(null);
                outSheet.getRow(dataRowNumber).getCell(totalNumberOfOrdersColumn).setCellValue(sumOfCell);
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

        int rowNum = consultRowNumber;
//        for (rowNum = inSheet.getLastRowNum(); rowNum >= 0; rowNum--) {
//            if (null == inSheet.getRow(rowNum)) continue;
//            if (null == inSheet.getRow(rowNum).getCell(0)) continue;
//            if (inSheet.getRow(rowNum).getCell(0).getStringCellValue().toUpperCase().startsWith("КОНС")) break;
//            if (inSheet.getRow(rowNum).getCell(0).getStringCellValue().length() > 60) break;
//        }

        int inTotlFedColumn;
        for (inTotlFedColumn = 3; inTotlFedColumn < vydachaColumn; inTotlFedColumn++) {
            XSSFCell cell = inSheet.getRow(consultRowNumber - 1).getCell(inTotlFedColumn);
            if ( null == cell) continue;
            if (!CellType.STRING.equals(cell.getCellTypeEnum())) continue;

            if ( inSheet.getRow(consultRowNumber - 1).getCell(inTotlFedColumn).getStringCellValue().toUpperCase().startsWith("ФЕДЕР"))
                break;
        }

        XSSFCell sourceCell = null;
        HSSFCell destinationCell = null;

        if ("fed".equals(auth)) {
            sourceCell = inSheet.getRow(consultRowNumber).getCell(inTotlFedColumn);
            destinationCell = outSheet.getRow(outRowOfTotalCell).getCell(outColumnOfTotalCell);
        } else if ("reg".equals(auth)) {
            sourceCell = inSheet.getRow(consultRowNumber).getCell(inTotlFedColumn + 1);
            destinationCell = outSheet.getRow(outRowOfTotalCell).getCell(outColumnOfTotalCell);
        } else if ("oth".equals(auth)) {
            sourceCell = inSheet.getRow(consultRowNumber).getCell(inTotlFedColumn + 3);
            destinationCell = outSheet.getRow(outRowOfTotalCell).getCell(outColumnOfTotalCell);
        }
        log.info("consultRowNumber == {}    inTotlFedColumn == {} ",consultRowNumber ,inTotlFedColumn);
        log.info("Totals: {}, {}, {}.", auth,sourceCell, destinationCell);

        copyXToHCell(sourceCell, destinationCell);

        log.info("Finished for {} -> {}==============================================", inSheet.getSheetName(), outSheet.getSheetName());

    }

    private static void copyXToHCell(XSSFCell sourceCell, HSSFCell destinationCell) {
        if (sourceCell == null) {
            log.info("sourceCell == null на входе copyXToHCell");
            return;
        }
        if(destinationCell == null) {
            log.info("destinationCell == null на входе copyXToHCell");
            return;
        }
        destinationCell.setCellFormula(null);
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
