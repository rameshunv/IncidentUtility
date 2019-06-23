package com.mphasis.incident.incidentUtility;

import org.apache.poi.ss.format.CellTextFormatter;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

public class IncidentUtility {

    public static final String MASTER_SERVER_LIST_XLSX_FILE_PATH = "../MasterData/MasterServerList-1.xlsx";
    public static final String TICKET_OWNER_XLSX_FILE_PATH = "../MasterData/TicketOwner_AppLead_master.xlsx";
    public static final String SOURCE_DIR = "../SourceData/";
    public static final String DESTINATION_DIR = "../DestinationData/";

   /* public static final String MASTER_SERVER_LIST_XLSX_FILE_PATH = "./MasterData/MasterServerList-1_new.xlsx";
    public static final String TICKET_OWNER_XLSX_FILE_PATH = "./MasterData/TicketOwner_AppLead_master.xlsx";
    public static final String SOURCE_DIR = "./SourceData/";
    public static final String DESTINATION_DIR = "./DestinationData/";*/

    static String[] appendixKeywords = {"Fulfillment","Automic_Scheduler","CID","CRDB","Delivery-Service","DocStore","EAS","eCommerce","Focal","Location Service",
            "Location-Service","MobileCentral-App","NAVSSC","NextGen","OnDemand","OrderTracking","OTP","PARC","Payment","Payment Server",
            "PPA","PrintService","PSG","PSSWLS","Sales-Navigator","UploadServer","Utility Service","VertexO","webMethods","WOSV","Xanedu"};


    static Map<String, String>  masterServerDataMap;
    static Map<String, String> hostMap;

    static Map<String, String> ticketOwnerDataMap;

    public static void main(String a[]) throws Exception{

        masterServerDataMap = MasterServerListExcelReader.loadMasterServerDataMap(MASTER_SERVER_LIST_XLSX_FILE_PATH);
       // System.out.println("masterServerDataMap >>>>>>>>>>"+masterServerDataMap.size());

        hostMap = prepareHostMap(masterServerDataMap.keySet());

        ticketOwnerDataMap = TicketOwnerExcelReader.loadTicketOwnerDataMap(TICKET_OWNER_XLSX_FILE_PATH);
      // System.out.println("ticketOwnerDataMap >>>>>>>>>>"+ticketOwnerDataMap.size());

     //  System.out.println("ticketOwnerDataMap >>>>>>>>>>"+ticketOwnerDataMap.get("Center-Repository"));


       List<String> fileNamesList  = UnZipFile.unzipFiles(SOURCE_DIR,DESTINATION_DIR);

       // System.out.println("fileNamesList >>>>>>>>>>"+fileNamesList.toString());

        for(String fileName : fileNamesList){
            if(isNotEmpty(fileName))
            processData(DESTINATION_DIR+fileName);

        }

     //   processData(DESTINATION_DIR+"InProgress_Incidents_11_06_2019.xlsx");

    }


    public  static void processData(String filePath) throws  Exception{

        File file = new File(filePath);
        Workbook workbook = WorkbookFactory.create(file);
        Sheet sheet = workbook.getSheetAt(0);

        Row headerRow = sheet.getRow(0);
        int lastCellNum = headerRow.getLastCellNum();

        Cell ticketOwnerCellHeader = headerRow.createCell(lastCellNum );
        ticketOwnerCellHeader.setCellType(CellType.STRING);
        ticketOwnerCellHeader.setCellValue("Ticket Owner & App Leader");
        ticketOwnerCellHeader.setCellStyle(headerRow.getRowStyle());


        Cell vzCorrelationIDCellHeader = headerRow.createCell(lastCellNum +1);
        vzCorrelationIDCellHeader.setCellType(CellType.STRING);
        vzCorrelationIDCellHeader.setCellValue("VZ Correlation ID");
        vzCorrelationIDCellHeader.setCellStyle(headerRow.getRowStyle());

        Cell dxcCorrelationIdCellHeader = headerRow.createCell(lastCellNum +2);
        dxcCorrelationIdCellHeader.setCellType(CellType.STRING);
        dxcCorrelationIdCellHeader.setCellValue("DXC Correction ID");
        dxcCorrelationIdCellHeader.setCellStyle(headerRow.getRowStyle());


        sheet.forEach(row -> {
            if(row.getRowNum() != 0) {


                String configurationItem = row.getCell(11).getRichStringCellValue().getString();

                if (isNotEmpty(configurationItem)) {
                  //  if (!staticApplicationName.contains(configurationItem)) {
                        configurationItem = getApplicationNameByMasterKey(configurationItem);

                   // }
                } else {
                    configurationItem = row.getCell(139).getRichStringCellValue().getString();
                    if (isNotEmpty(configurationItem)) {

                       String key = getApplicationNameAFromKeywords(configurationItem);

                       if(isNotEmpty(key)){
                           configurationItem = key;

                       }else{
                           configurationItem = getApplicationNameByHostMapKey(configurationItem);
                       }

                    }

                }
                if (isNotEmpty(configurationItem)) {
                    row.getCell(11).setCellType(CellType.STRING);
                    row.getCell(11).setCellValue(configurationItem+"");

                    System.out.println("Last sellllll "+row.getLastCellNum());
                    Cell cell =  row.createCell(lastCellNum);
                    cell.setCellType(CellType.STRING);
                    cell.setCellValue(ticketOwnerDataMap.get(configurationItem));

                    System.out.println(configurationItem + "\t" + ticketOwnerDataMap.get(configurationItem));

                }

               System.out.println(" Processed Records >>>>>>>>>>>>>>>>>>>>>>>>>>>>>"+ row.getRowNum() + " Out of " + sheet.getLastRowNum());
            }
        });


        String newFilePath =  filePath.replace("InProgress","Clean");

        workbook.write(new FileOutputStream(new File(newFilePath)));
        workbook.close();

        System.out.println("Final Processed File Location >>>>>>>>>>>>>>>>>>>>>>>>>>>>>"+newFilePath);
        file.delete();
    }


    // sample :- acsfxf01.auth.rf.fedex.com ? key = acsfxf01 , value = acsfxf01.auth.rf.fedex.com
    public static Map<String,String> prepareHostMap(Set<String> keySet){

        Map<String,String> map = new HashMap<>();
        for (String part : keySet){
             if(part != null){

                 map.put(part.substring(0,part.indexOf(".")),part);
             }

             }
      return  map;

    }

    public static String getApplicationNameAFromKeywords(String configurationItem){

        for(String key : appendixKeywords){
            if(key.toUpperCase().contains(configurationItem.toUpperCase())){
                return key;
            }
        }

        return null;
    }



    public static String getApplicationNameByHostMapKey(String configurationItem){

        for(String key : hostMap.keySet()){
            if(key != null && configurationItem.toUpperCase().contains(key.toUpperCase())){

               return  masterServerDataMap.get(hostMap.get(key));

            }
        }

        return null;
    }


    //Getting AppicationName from MasterServerList
    public static String getApplicationNameByMasterKey(String configurationItem){

       for(String key : masterServerDataMap.keySet()){

           if(key != null && key.toUpperCase().contains(configurationItem.toUpperCase())){
               String appName = masterServerDataMap.get(key);
             //  return  masterServerDataMap.get(key);
               return appName;
           }
       }

        return configurationItem;
    }

    public static boolean isNotEmpty(String data){
        if(data == null) return false;

        if("".equals(data.trim())) return false;
     //   if("null".equals(data.trim())) return false;
     //   if("NULL".equals(data.trim())) return false;

        return true;
    }
}
