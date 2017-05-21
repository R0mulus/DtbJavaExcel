/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelport;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;

enum MyEnum{
    T,F
}

/**
 *
 * @author Joseph
 */
public class ConnectionProvider {
    private static final String USERNAME = "root";
    private static final String PASSWORD = "";
    private static final String URL = "jdbc:mysql://localhost/";
    private static final String DBNAME = "world_x";
    private static final String DRIVER = "com.mysql.jdbc.Driver";
    private HSSFWorkbook workbook = new HSSFWorkbook();
    
    private Connection getConnection() {
        Connection conn = null;
        try {
            conn = DriverManager.getConnection(URL + DBNAME, USERNAME, PASSWORD);
        } catch (SQLException ex) {
            System.out.println("Error: " + ex.toString());
        }
        return conn;
    }
    
    public void fillExcelSheetWithCities(){
        try {
            Class.forName(DRIVER);
            String query = "Select * from city";
            Connection conn = getConnection();
            Statement st = conn.createStatement();
            ResultSet rs = st.executeQuery(query);
            
            //HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("Cities");
            workbook.setSheetOrder("Cities",0);
            HSSFRow rowhead = sheet.createRow((short) 0);
            rowhead.createCell((short) 0).setCellValue("City ID");
            rowhead.createCell((short) 1).setCellValue("Name");
            rowhead.createCell((short) 2).setCellValue("Coutry code");
            rowhead.createCell((short) 3).setCellValue("District");
            int i = 1;
            
            while (rs.next()){
                HSSFRow row = sheet.createRow((short) i);
                row.createCell((short) 0).setCellValue(Integer.toString(rs.getInt("ID")));
                row.createCell((short) 1).setCellValue(rs.getString("Name"));
                row.createCell((short) 2).setCellValue(rs.getString("CountryCode"));
                row.createCell((short) 3).setCellValue(rs.getString("District"));
                i++;
            }
            
            String dest = "e:/test.xls";
            FileOutputStream fileOut = new FileOutputStream(dest);
            workbook.write(fileOut);
            fileOut.close();
        } catch (ClassNotFoundException | SQLException | IOException e1) {
            System.out.println("Error: " + e1);
        }
    }
    
    public void fillExcelSheetWithCountries(){
        try {
            Class.forName(DRIVER);
            String query = "Select * from country";
            Connection conn = getConnection();
            Statement st = conn.createStatement();
            ResultSet rs = st.executeQuery(query);
            
            //HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("Countries");
            workbook.setSheetOrder("Countries",1);
            HSSFRow rowhead = sheet.createRow((short) 0);
            rowhead.createCell((short) 0).setCellValue("Country code (3)");
            rowhead.createCell((short) 1).setCellValue("Name");
            rowhead.createCell((short) 2).setCellValue("Capital");
            rowhead.createCell((short) 3).setCellValue("Country code (2)");
            int i = 1;
            
            while (rs.next()){
                HSSFRow row = sheet.createRow((short) i);
                row.createCell((short) 0).setCellValue(rs.getString("Code"));
                row.createCell((short) 1).setCellValue(rs.getString("Name"));
                row.createCell((short) 2).setCellValue(Integer.toString(rs.getInt("Capital")));
                row.createCell((short) 3).setCellValue(rs.getString("Code2"));
                i++;
            }
            
            String dest = "e:/test.xls";
            FileOutputStream fileOut = new FileOutputStream(dest);
            workbook.write(fileOut);
            fileOut.close();
        } catch (ClassNotFoundException | SQLException | IOException e1) {
            System.out.println("Error: " + e1);
        }
    }
    
    public void fillExcelSheetWithLanguages(){
        try {
            Class.forName(DRIVER);
            String query = "Select * from countrylanguage";
            Connection conn = getConnection();
            Statement st = conn.createStatement();
            ResultSet rs = st.executeQuery(query);
            
            //HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("Language");
            workbook.setSheetOrder("Language",2);
            HSSFRow rowhead = sheet.createRow((short) 0);
            rowhead.createCell((short) 0).setCellValue("Country code (3)");
            rowhead.createCell((short) 1).setCellValue("Language");
            rowhead.createCell((short) 2).setCellValue("Is it official ?");
            rowhead.createCell((short) 3).setCellValue("Percentage");
            int i = 1;
            String enumResult;
            
            while (rs.next()){
                HSSFRow row = sheet.createRow((short) i);
                row.createCell((short) 0).setCellValue(rs.getString("CountryCode"));
                row.createCell((short) 1).setCellValue(rs.getString("Language"));
                MyEnum enumVal =  MyEnum.valueOf(rs.getString("IsOfficial"));
                if(enumVal == MyEnum.T){
                    enumResult = "Yes";
                }else enumResult = "No";
                row.createCell((short) 2).setCellValue(enumResult);
                row.createCell((short) 3).setCellValue(Float.toString(rs.getFloat("Percentage")));
                i++;
            }
            
            String dest = "e:/test.xls";
            FileOutputStream fileOut = new FileOutputStream(dest);
            workbook.write(fileOut);
            fileOut.close();
        } catch (ClassNotFoundException | SQLException | IOException e1) {
            System.out.println("Error: " + e1);
        }
    }
    
}
