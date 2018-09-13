/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package writepoiexcel;

import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.List;
import javax.swing.JFileChooser;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Kevin
 */
public class WriteListToExcelFile {
    public void writeListToFile(String fileName, List<POSreport> posList) throws Exception{
        Workbook workbook = null;
        if(fileName.toLowerCase().endsWith("xlsx")||fileName.toLowerCase().endsWith("xlsm")){
            workbook = new XSSFWorkbook();
        }else if(fileName.toLowerCase().endsWith("xls")||fileName.toLowerCase().endsWith("csv")){
            workbook = new HSSFWorkbook();
        }
        else throw new Exception("invalid file name, should be excel file");
        
        //create sheet
        Sheet sheet = workbook.createSheet("sheetName");
        
        //Create iterator to travel the the List of Pos
        Iterator<POSreport> posIterator = posList.iterator();
        
        int rowIndex = 0;
        
        while(posIterator.hasNext()){
            POSreport pos = posIterator.next();
            Row row = sheet.createRow(rowIndex++); 
                if(pos.getSN()=="")
                break;
               writeCell(pos,row);
            
        }
        FileOutputStream outputstream = new FileOutputStream(fileName);
        workbook.write(outputstream);
        
    }
    
   /* private void writeCell(POSreport pos,Row row){
        Cell cell = row.createCell(1);
        cell.setCellValue(pos.getSN());
        
        cell = row.createCell(2);
        cell.setCellValue(pos.getCompany());
        
        cell = row.createCell(3);
        cell.setCellValue(pos.getModel());
        
        cell = row.createCell(4);
        cell.setCellValue(pos.getDMIrev());
        
        cell = row.createCell(5);
        cell.setCellValue(pos.getDMIsn());
        
        cell = row.createCell(6);
        cell.setCellValue(pos.getSKU());
    }*/
    private void writeCell(POSreport pos,Row row){
        Cell cell = row.createCell(0);       
           cell.setCellValue(pos.getSN()+";"+pos.getCompany()+";"+pos.getModel()+";"+pos.getDMIrev()+";"+pos.getDMIsn());
        
            
    }
}
