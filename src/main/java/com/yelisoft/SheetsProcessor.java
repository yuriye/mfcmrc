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


    public static void process(XSSFSheet inSheet, HSSFSheet outSheet, int outServiceNameColumn, String auth) throws IOException {
        Config config = Config.getInstance();
        Map<String, Integer> inServiceRow = new HashMap<>();
        int inServiceNameColumn = 1;
        int inSheetStartRow = 10;
        int dataColumn;
        int totalColumn = 18;
        int outRowOfTotalCell = 0;
        int outColumnOfTotalCell = outServiceNameColumn + 5;

        for (dataColumn = 7; dataColumn < 18; dataColumn++) {
            if (config.getMonth().toUpperCase().equals(
                    inSheet.getRow(inSheetStartRow - 2).getCell(dataColumn).getStringCellValue().toUpperCase()))
                break;
        }
//        System.out.println(dataColumn);

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
            String monthData;
            String rowTotalData;
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

            Object inData;
            if(CellType.NUMERIC.equals(inDataCell.getCellTypeEnum())) {
                inData = inDataCell.getNumericCellValue();
                outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn + 2).setCellValue((double)inData);

            }
            else {
                inData = inDataCell.getStringCellValue();
                outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn + 2).setCellValue((String)inData);

            }

            XSSFCell inTotalCell = inRow.getCell(totalColumn);
            String totalData = inTotalCell.getStringCellValue();
            outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn + 3).setCellValue(totalData);

            if("".equals(outSheet.getRow(dataRowNumber).getCell(outServiceNameColumn).getStringCellValue())) break;
        }

        int rowNum;
        for(rowNum = inSheet.getLastRowNum(); rowNum >= 0; rowNum--) {
            System.out.println("rowNum=" + rowNum);
            if (null == inSheet.getRow(rowNum)) continue;
            if (null == inSheet.getRow(rowNum).getCell(0)) continue;
            if (inSheet.getRow(rowNum).getCell(0).getStringCellValue().startsWith("Консул")) break;
        }
        Object consTotal;
        if ("fed".equals(auth)) {
            if (CellType.NUMERIC.equals(inSheet.getRow(rowNum).getCell(16).getCellTypeEnum())) {
                consTotal = inSheet.getRow(rowNum).getCell(16).getNumericCellValue();
                outSheet.getRow(outRowOfTotalCell).getCell(outColumnOfTotalCell).setCellValue((double)consTotal);
            }
            else {
                consTotal = inSheet.getRow(rowNum).getCell(16).getStringCellValue();
                outSheet.getRow(outRowOfTotalCell).getCell(outColumnOfTotalCell).setCellValue((String) consTotal);
            }
        }
        else if ("reg".equals(auth)) {
            if (CellType.NUMERIC.equals(inSheet.getRow(rowNum).getCell(17).getCellTypeEnum())) {
                consTotal = inSheet.getRow(rowNum).getCell(17).getNumericCellValue();
                outSheet.getRow(outRowOfTotalCell).getCell(outColumnOfTotalCell).setCellValue((double)consTotal);
            }
            else {
                consTotal = inSheet.getRow(rowNum).getCell(17).getStringCellValue();
                outSheet.getRow(outRowOfTotalCell).getCell(outColumnOfTotalCell).setCellValue((String) consTotal);
            }
        }
    }
}
