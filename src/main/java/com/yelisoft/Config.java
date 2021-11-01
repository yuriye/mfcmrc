package com.yelisoft;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedReader;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by Yuriy Yeliseyev on 21.11.2017.
 */
public class Config {
    private static final Logger log = LoggerFactory.getLogger(Config.class);
    private static Config instance;

    private final Map<String, String> sheetsMap = new HashMap<>();
    private String inputFolderName = "C:\\mfcmrc";
    private String templatesFolderName = "templates";
    private String outputFolderName = "out";
    private String pagesComplianceFileName = "pagescompliance.xlsx";
    private String inputFileName = "+++";
    private String month = "";

    private String servicesComplianceFileName = "соответствия.xlsx";
    private final Map<String, String> servicesInToOutMap = new HashMap<>();
    private final Map<String, Boolean> hasOutputOfDocs = new HashMap<>();

    public boolean hasOutputForService(String service) {
        Boolean has = hasOutputOfDocs.get(service);
        return null != has && has;
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
        BufferedReader br = Files.newBufferedReader(Paths.get(fileName), StandardCharsets.UTF_8);
        br.readLine();
        String line;
        while (true) {
            line = br.readLine();
            log.debug(line);
            if (null == line) break;
            if ("".equals(line)) continue;
            if (line.startsWith("#")) continue;
            if (line.startsWith("//")) continue;
            String[] array = line.split("=");
            if(array.length < 2) {
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
                    servicesComplianceFileName = array[1];
                    break;
            }
        }
        br.close();
    }

    public String getTemplatesFolderName() {
        return templatesFolderName;
    }

    public String getFullPagesComplianceFileName() {
        return inputFolderName + "/" + pagesComplianceFileName;
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

    public String getOutputFolderName() {
        return inputFolderName + "\\" + outputFolderName;
    }

    public String getInputFileName() {
        return inputFileName;
    }

    public String getMonth() {
        return month;
    }

    public void setOutForInService(String inService, String outService) {
        servicesInToOutMap.put(inService, outService);
    }

    public String getOutForInService(String inService) {
        return servicesInToOutMap.get(inService);
    }

    public String getFullServicesComplianceFileName() {
        return inputFolderName + "\\" + servicesComplianceFileName;
    }

}
