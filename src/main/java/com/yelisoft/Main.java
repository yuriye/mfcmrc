package com.yelisoft;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.filefilter.WildcardFileFilter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.Collection;
import java.util.HashSet;
import java.util.Set;

public class Main {

    private static final Logger log = LoggerFactory.getLogger(Main.class);
    private static final Set<String> absents = new HashSet<>();

    public static void main(String[] args) throws IOException {

        Config config = Config.getInstance();
        String configFileName = "C:/mfcmrc/config.cfg";
        if (args.length > 0) {
            configFileName = args[0];
        }
        config.initFromFile(configFileName);

        XSSFWorkbook pagesComplianceBook = new XSSFWorkbook(new FileInputStream(config.getFullPagesComplianceFileName()));
        XSSFSheet pagesComplianceSheet = pagesComplianceBook.getSheetAt(0);
        for (int i = 0; true; i++) {
            XSSFRow row = pagesComplianceSheet.getRow(i);
            if (null == row) break;
            String val0;
            if (row.getCell(0).getCellTypeEnum() == CellType.NUMERIC) {
                val0 = row.getCell(0).getRawValue();
            } else {
                val0 = row.getCell(0).getStringCellValue();
            }
            if ("".equals(val0)) break;
            config.setComplianceSheetName(val0, row.getCell(1).getStringCellValue());
        }

        XSSFWorkbook servicesComplianceBook = new XSSFWorkbook(new FileInputStream(config.getFullServicesComplianceFileName()));
        XSSFSheet servicesComplianceSheet = servicesComplianceBook.getSheetAt(0);
        for (int i = 2; true; i++) {
            XSSFRow row = servicesComplianceSheet.getRow(i);
            if (null == row) break;
            if (null == row.getCell(0)) continue;
            String outService = row.getCell(0).getStringCellValue();
            if ("".equals(outService)) continue;
            String inService = row.getCell(1).getStringCellValue();
            if ("".equals(inService)) continue;
            config.setOutForInService(inService, outService);

            String tmp;
            try {
                tmp = row.getCell(2).getStringCellValue();
            } catch (Exception e) {
                tmp = "";
            }

            tmp = tmp.trim();
            config.setOutputForService(outService, tmp.length() <= 15);
        }

        System.out.println(config.getOutputFolderName());
        FileUtils.deleteDirectory(new File(config.getOutputFolderName()));
        FileUtils.copyDirectory(new File(config.getInputFolderName() + "/" + config.getTemplatesFolderName())
                , new File(config.getOutputFolderName())
                , new WildcardFileFilter("feder*.xls*"));
        FileUtils.copyDirectory(new File(config.getInputFolderName() + "/" + config.getTemplatesFolderName())
                , new File(config.getOutputFolderName())
                , new WildcardFileFilter("region*.xls*"));
        FileUtils.copyDirectory(new File(config.getInputFolderName() + "/" + config.getTemplatesFolderName())
                , new File(config.getOutputFolderName())
                , new WildcardFileFilter("otherServ*.xls*"));
        String[] exts = {"xls", "xlsx"};

        Workbook inBook = null;
        File inFile = new File(config.getInputFolderName(), config.getInputFileName());
        try {
            inBook = WorkbookFactory.create(inFile, "", true);
        } catch (Exception e) {
            e.printStackTrace();
            log.error(e.getMessage());
            System.exit(1);
        }

        log.info("Input file: " + config.getInputFolderName() + "/" + config.getInputFileName());
        Collection<File> files = FileUtils.listFiles(new File(config.getOutputFolderName()), exts, false);
        for (File bookFile : files) {
            log.info("Output file name: " + bookFile.getCanonicalFile());
            int cellOffset = 4;
            HSSFWorkbook outBook = new HSSFWorkbook(new FileInputStream(bookFile));
            log.info("outBook.isWriteProtected() = {}", outBook.isWriteProtected());
            int numberOfSheets = outBook.getNumberOfSheets();
            for (int i = 0; i < numberOfSheets; i++) {
                HSSFSheet outSheet = outBook.getSheetAt(i);
                String inSheetName = config.getComplianceSheetName(outSheet.getSheetName());
                if (null == inSheetName) continue;
                Sheet inSheet = inBook.getSheet(inSheetName);
                if (null == inSheet) continue;
                SheetsProcessor.process(inSheet, outSheet, cellOffset, bookFile.getName());
                absents.addAll(SheetsProcessor.getAbsents());
            }
            FileOutputStream outStream = new FileOutputStream(bookFile);
            outBook.write(outStream);
            outStream.close();
            outBook.close();
        }

        inBook.close();

        FileWriter writer = new FileWriter(config.getOutputFolderName() + "/отсутствующие услуги.csv");
        writer.write("Наименование\n");
        absents.forEach(s -> writeAbsent(writer, s));
        writer.flush();
        writer.close();
        log.info("Main finished");
    }

    public static void writeAbsent(FileWriter writer, String s) {
        try {
            writer.write(s + "\n");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
