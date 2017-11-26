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

public class SheetsProcessor {

    public static void process(XSSFSheet inSheet,
                               HSSFSheet outSheet,
                               int outServiceNameColumn,
                               String auth) throws IOException {

        Config config = Config.getInstance();
        Map<String, Integer> inServiceRow = new HashMap<>();
        int inServiceNameColumn = 1;
        int inSheetStartRow = 10;
        int dataColumn;
        int totalColumn = 19;
        int outRowOfTotalCell = 0;
        int outColumnOfTotalCell = outServiceNameColumn + 5;

        for (dataColumn = 7; dataColumn < totalColumn; dataColumn++) {
            if (config.getMonth().toUpperCase().equals(
                    inSheet.getRow(inSheetStartRow - 2).getCell(dataColumn).getStringCellValue().toUpperCase()))
                break;
        }

        for(int i = inSheetStartRow; true; i++) {
            XSSFRow row  = inSheet.getRow(i);
            if (null == row) break;
            XSSFCell cell = row.getCell(inServiceNameColumn);
            if (cell == null) continue;
            String inServiceName = cell.getStringCellValue();
            if("".equals(inServiceName) || null == inServiceName) continue;
            inServiceRow.put(inServiceName, i);
        }

        System.out.println(inSheet.getSheetName() + "->" + outSheet.getSheetName());
        for (int dataRowNumber = 7; true ; dataRowNumber++) {
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
                int rowNumber = Integer.valueOf(cell.getStringCellValue());
                System.out.println(rowNumber + ":" );


            }
            catch (NumberFormatException mfe) {
                continue;
            }
            String outService = outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn).getStringCellValue();
            String inService = config.getInForOutService(outService);

            if(null == inService) continue;

            System.out.println("inService:" + inService );
            System.out.println("row=" + inServiceRow.get(inService));
            if (null == inServiceRow.get(inService)) continue;
            XSSFRow inRow = inSheet.getRow(inServiceRow.get(inService));
            XSSFCell inDataCell = inRow.getCell(dataColumn);
            if (null == inDataCell) continue;

            copyXToHCell(inDataCell, outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn + 2));
            copyXToHCell(inRow.getCell(totalColumn), outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn + 3));

            if("".equals(outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn).getStringCellValue())) break;
        }

        int rowNum;
        for(rowNum = inSheet.getLastRowNum(); rowNum >= 0; rowNum--) {
            System.out.println("rowNum=" + rowNum);
            if (null == inSheet.getRow(rowNum)) continue;
            if (null == inSheet.getRow(rowNum).getCell(0)) continue;
            if (inSheet.getRow(rowNum).getCell(0).getStringCellValue().startsWith("Консул")) break;
        }
        if ("fed".equals(auth)) {
            copyXToHCell(inSheet.getRow(rowNum).getCell(16),
                    outSheet.getRow(outRowOfTotalCell).getCell(outColumnOfTotalCell));
        }
        else if ("reg".equals(auth)) {
            copyXToHCell(inSheet.getRow(rowNum).getCell(17),
                    outSheet.getRow(outRowOfTotalCell).getCell(outColumnOfTotalCell));
        }
    }

    private static void copyXToHCell(XSSFCell sourceCell, HSSFCell destinationCell) {
        if (CellType.BLANK.equals(sourceCell.getCellTypeEnum())) {}
        else if (CellType.NUMERIC.equals(sourceCell.getCellTypeEnum())) {
            destinationCell.setCellValue(sourceCell.getNumericCellValue());
        } else if (CellType.STRING.equals(sourceCell.getCellTypeEnum())) {
            destinationCell.setCellValue(sourceCell.getStringCellValue());
        } else if (CellType.BOOLEAN.equals(sourceCell.getCellTypeEnum())) {
            destinationCell.setCellValue(sourceCell.getBooleanCellValue());
        }
    }

}
