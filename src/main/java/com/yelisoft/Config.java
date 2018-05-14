package com.yelisoft;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Created by Yuriy Yeliseyev on 21.11.2017.
 */
public class Config {
    private static final Logger log = LoggerFactory.getLogger(Config.class);
    private static Config instance;

    private Map<String, String> sheetsMap = new HashMap<>();
    private String inputFolderName = "C:\\mfcmrc";
    private String templatesFolderName = "templates";
    private String outputFolderName = "out";
    private String pagesComplianceFileName = "pagescompliance.xlsx";
//    private String inputFileName = "АКТУАЛЬНЫЙ МФЦ  2017 ЗАКРЫТЫЙ Октябрь.xlsx";
//    private String month = "октябрь";
    private String inputFileName = "+++";
    private String month = "";

    private String servicesComplianceFileName = "соответствия.xlsx";
    private Map<String, String> servicesOutToInMap = new HashMap<>();
    private Map<String, Boolean> hasOutputOfDocs = new HashMap<>();

    public Map<String, Boolean> getHasOutputOfDocs() {
        return hasOutputOfDocs;
    }

    public void setHasOutputOfDocs(Map<String, Boolean> hasOutputOfDocs) {
        this.hasOutputOfDocs = hasOutputOfDocs;
    }

    public boolean hasOutputForService(String service) {
        Boolean has = hasOutputOfDocs.get(service);
        if (null != has && has) return true;
        return false;
    }

    public void setOutputForService(String service, Boolean has) {
        if("".equals(service)) return;
        if(null == service || null == has) return;
        hasOutputOfDocs.put(service, has);
    }


    private Config() {}

    public static Config getInstance() {
        if(instance == null) instance = new Config();
        return instance;
    }

    public void initFromFile(String fileName) throws IOException {
//        BufferedReader br = new BufferedReader(new FileReader(fileName));
        BufferedReader br = Files.newBufferedReader(Paths.get(fileName), Charset.forName("UTF-8"));
        String line;
        line = br.readLine();
        while (true) {
            line = br.readLine();
//            System.out.println(line);
            log.debug(line);
            if (null == line) break;
            if ("".equals(line)) continue;
            if (line.startsWith("#")) continue;
            if (line.startsWith("//")) continue;
            String[] array = line.split("=");
            if(array.length < 2) {
//                System.out.println("Неправильная строка каонфигурации:" + line);
                log.debug("Неправильная строка каонфигурации:" + line);
                continue;
            }
            array[0] = array[0].trim();
            array[1] = array[1].trim();
            switch (array[0]) {
                case "month":
                    month = array[1];
                    break;
                case "baseDirectory":
                    inputFolderName = array[1];
                    break;
                case "templatesDirectory":
                    templatesFolderName = array[1];
                    break;
                case "outputDirectory":
                    outputFolderName = array[1];
                case "pagesComplianceFileName":
                    pagesComplianceFileName = array[1];
                    break;
                case "inputFileName":
                    inputFileName = array[1];
                    break;
                case "servicesComplianceFileName":
                    break;
            }
        }
        br.close();
    }

    public String getTemplatesFolderName() {
        return templatesFolderName;
    }

    public void setTemplatesFolderName(String templatesFolderName) {
        this.templatesFolderName = templatesFolderName;
    }

    public String getFullPagesComplianceFileName() {
        return inputFolderName + "/" + pagesComplianceFileName;
    }

    public String getFullInputFileName() {
        return inputFolderName + "/" + inputFileName;
    }

    public String getInputFolderName() {
        return inputFolderName;
    }

    public String getComplianceSheetName(String name) {
        return sheetsMap.get(name);
    }

    public void setComplianceSheetName(String in, String out) {
        sheetsMap.put(in, out);
    }

    public static void setInstance(Config instance) {
        Config.instance = instance;
    }

    public void setInputFolderName(String inputFolderName) {
        this.inputFolderName = inputFolderName;
    }

    public String getOutputFolderName() {
        return inputFolderName + "\\" + outputFolderName;
    }

    public void setOutputFolderName(String outputFolderName) {
        this.outputFolderName = outputFolderName;
    }

    public String getPagesComplianceFileName() {
        return pagesComplianceFileName;
    }

    public void setPagesComplianceFileName(String pagesComplianceFileName) {
        this.pagesComplianceFileName = pagesComplianceFileName;
    }

    public String getInputFileName() {
        return inputFileName;
    }

    public void setInputFileName(String inputFileName) {
        this.inputFileName = inputFileName;
    }

    public String getMonth() {
        return month;
    }

    public void setMonth(String month) {
        this.month = month;
    }

    public Map<String, String> getSheetsMap() {
        return sheetsMap;
    }

    public void setSheetsMap(Map<String, String> sheetsMap) {
        this.sheetsMap = sheetsMap;
    }

    public String getServicesComplianceFileName() {
        return servicesComplianceFileName;
    }

    public void setServicesComplianceFileName(String servicesComplianceFileName) {
        this.servicesComplianceFileName = servicesComplianceFileName;
    }

    public Map<String, String> getServicesOutToInMap() {
        return servicesOutToInMap;
    }

    public void setServicesOutToInMap(Map<String, String> servicesOutToInMap) {
        this.servicesOutToInMap = servicesOutToInMap;
    }

    public String getInForOutService(String outService) {
        return servicesOutToInMap.get(outService);
    }

    public void setInForOutService(String outService, String inService) {
        servicesOutToInMap.put(outService, inService);
    }

    public String getFullServicesComplianceFileName() {
        return inputFolderName + "\\" + servicesComplianceFileName;
    }

}
