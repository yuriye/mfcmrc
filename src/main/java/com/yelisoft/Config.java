package com.yelisoft;

import java.util.HashMap;
import java.util.Map;

/**
 * Created by Yuriy Yeliseyev on 21.11.2017.
 */
public class Config {
    private static Config instance;

    private Map<String, String> sheetsMap = new HashMap<>();
    private String inputFolderName = "C:/mfcmrc";
    private String outputFolderName = "out";
    private String pagesComplianceFileName = "pagescompliance.xlsx";
    private String inputFileName = "АКТУАЛЬНЫЙ МФЦ  2017 ЗАКРЫТЫЙ Октябрь.xlsx";
    private String month = "октябрь";
    private String servicesComplianceFileName = "соответствия.xlsx";
    private Map<String, String> servicesOutToInMap = new HashMap<>();

    private Config() {}

    public static Config getInstance() {
        if(instance == null) instance = new Config();
        return instance;
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
        return inputFolderName + "/" + outputFolderName;
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
        return inputFolderName + "/" + servicesComplianceFileName;
    }

}
