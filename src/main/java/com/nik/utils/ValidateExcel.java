package com.nik.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.NotOLE2FileException;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ValidateExcel {

    static boolean validateExcel(String filePath){

        try(FileInputStream fileInputStream = new FileInputStream(filePath)){
            System.out.println("check point1");

            if (filePath.endsWith(".xls") ) {
                System.out.println("file extension- .xls ");
            } else if(filePath.endsWith(".xlsx")) {
                System.out.println("file extension- .xlsx");
            }
            if(fileInputStream.available()>0){
                System.out.println("check point2");
            }

            new XSSFWorkbook(fileInputStream);
            return true;
        }catch (OfficeXmlFileException exception){
            System.out.println("Invalid file format (possibly an XLSX file): "+exception.getMessage());
        }
        catch (NotOLE2FileException e1){
            System.out.println("Invalid file format : "+e1.getMessage());
        }
        catch (IOException e) {
            System.out.println("File not found Exception");
            throw new RuntimeException(e);
        }


        return false;
    }

    public static void main(String[] args) {

        String filePath = "src/main/resources/iShares-Core-US-Aggregate-Bond-ETF_fund1.xlsx";
        System.out.println("Validation excel file: "+validateExcel(filePath));


    }
}
