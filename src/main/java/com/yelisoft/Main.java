package com.yelisoft;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.filefilter.WildcardFileFilter;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Collection;

public class Main {

    public static void main(String[] args) throws IOException {
        Config config = Config.getInstance();

        XSSFWorkbook pagesComplianceBook = new XSSFWorkbook(new FileInputStream(config.getFullPagesComplianceFileName()));
        XSSFSheet pagesComplianceSheet = pagesComplianceBook.getSheetAt(0);
        for (int i = 0; true; i++) {
            XSSFRow row = pagesComplianceSheet.getRow(i);
            if (null == row) break;
            String val0 = row.getCell(0).getStringCellValue();
            if ("".equals(val0)) break;
            config.setComplianceSheetName(val0, row.getCell(1).getStringCellValue());
        }

        XSSFWorkbook servicesComplianceBook = new XSSFWorkbook(new FileInputStream(config.getFullServicesComplianceFileName()));
        XSSFSheet servicesComplianceSheet = servicesComplianceBook.getSheetAt(0);
        int count = 0;
        for (int i = 2; true; i++) {
            XSSFRow row = servicesComplianceSheet.getRow(i);
            if (null == row) break;
            if (null == row.getCell(0)) continue;
            String outService = row.getCell(0).getStringCellValue();
            if ("".equals(outService)) continue;
            String inService = row.getCell(1).getStringCellValue();
            if ("".equals(inService)) continue;
            config.setInForOutService(outService, inService);
            String tmp = "";
            try {
                tmp = row.getCell(2).getStringCellValue();
            }
            catch (Exception e) {}
            tmp = tmp == null? "": tmp;
            config.setOutputForService(outService, "".equals(tmp)? false: true);

            count = i - 2;
        }

        System.out.println(config.getOutputFolderName());
        FileUtils.deleteDirectory(new File(config.getOutputFolderName()));
        FileUtils.copyDirectory(new File(config.getInputFolderName() + "/" + config.getTemplatesFolderName())
                ,new File(config.getOutputFolderName())
                ,new WildcardFileFilter("feder*.xls*"));
        FileUtils.copyDirectory(new File(config.getInputFolderName() + "/" + config.getTemplatesFolderName())
                ,new File(config.getOutputFolderName())
                ,new WildcardFileFilter("region*.xls*"));
        FileUtils.copyDirectory(new File(config.getInputFolderName() + "/" + config.getTemplatesFolderName())
                ,new File(config.getOutputFolderName())
                ,new WildcardFileFilter("otherServ*.xls*"));
        String[] exts = {"xls", "xlsx" };

        FileInputStream fileInputStream = new FileInputStream(config.getInputFolderName() + "/" + config.getInputFileName());
        XSSFWorkbook inBook = new XSSFWorkbook(fileInputStream);
        System.out.println("Input file: " + config.getInputFolderName() + "/" + config.getInputFileName());

        Collection<File> files = FileUtils.listFiles(new File(config.getOutputFolderName()), exts, false);
        for (File bookFile: files) {
            System.out.println("Output file name: " + bookFile.getName());
            int cellOffset = 4;
            if (bookFile.getName().startsWith("other")) cellOffset = 2;
            HSSFWorkbook outBook = new HSSFWorkbook(new FileInputStream(bookFile));
            int numberOfSheets = outBook.getNumberOfSheets();
            for (int i = 0; i < numberOfSheets; i++) {
                HSSFSheet outSheet = outBook.getSheetAt(i);
                String inSheetName = config.getComplianceSheetName(outSheet.getSheetName());
                if (null == inSheetName) continue;
                XSSFSheet inSheet = inBook.getSheet(inSheetName);
                if (null == inSheet) continue;
                String auth = "";
                if(bookFile.getName().startsWith("fed")) auth = "fed";
                else if(bookFile.getName().startsWith("reg")) auth = "reg";
                SheetsProcessor.process(inSheet, outSheet, cellOffset, auth);
            }
            FileOutputStream outStream = new FileOutputStream(bookFile);
            outBook.write(outStream);
            outStream.close();
            outBook.close();
        }
        inBook.close();

    }



}
