/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package writepoiexcel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
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
public class ReadExcelFileToList {
    public List<POSreport> readExcelData(String fileName) throws IOException{
    List<POSreport> pos = new ArrayList<POSreport>();
        try {
            FileInputStream fis = new FileInputStream(fileName);
            //create instance workbook for reading both xlsx,xls
            Workbook workbook = null;
            if(fileName.toLowerCase().endsWith("xlsx")||fileName.toLowerCase().endsWith("xlsm")){
                workbook = new XSSFWorkbook(fis);
            }else if(fileName.toLowerCase().endsWith("xls")){
                workbook = new HSSFWorkbook(fis);
            }
            //get sheet position
            Sheet sheet = workbook.getSheetAt(3);
            //create iterator for row
            Iterator<Row> rowIterator = sheet.iterator();
            while(rowIterator.hasNext()){
                //get the row and travel through each line
                Row row = rowIterator.next();
                //skip the first line
               // if(row.getRowNum()==0)
                 //   continue;
                //create an POS object to store value
                POSreport posObj = new POSreport();
                //create iterator for cell
                Iterator<Cell> cellIterator = row.cellIterator();
                while(cellIterator.hasNext()){
                    //get the cell object and travel through each cell
                    Cell cell = cellIterator.next();
                  
                    //get the column index to assign appropriate value to Object
                    int columnIndex = cell.getColumnIndex();
                    
                    //using switch to assign the cells' value to POS object
                    switch (columnIndex-1){
                        case 1:
                            posObj.setSN(String.valueOf(getCellValue(cell)) );
                            break;
                        case 2:
                            posObj.setCompany(String.valueOf(getCellValue(cell)));
                            break;
                        case 3:
                            posObj.setModel(String.valueOf(getCellValue(cell)));
                            break;
                        case 4:
                            posObj.setDMIrev(String.valueOf(getCellValue(cell)));
                            break;
                        case 5:
                            posObj.setDMIsn(String.valueOf(getCellValue(cell)));
                            break;
                        case 6:
                            posObj.setSKU(String.valueOf(getCellValue(cell)));
                            break;
                    }//end of swich
                    
                }//end of cells iterator
                if(posObj.getSN()=="")
                    break;
                pos.add(posObj);
            }//end of rows
            fis.close();
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ReadExcelFileToList.class.getName()).log(Level.SEVERE, null, ex);
        }
        return pos;
    }
    private Object getCellValue(Cell cell){
        //check cell type and process accordingly
        switch (cell.getCellType()){
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            case Cell.CELL_TYPE_FORMULA:
                SimpleDateFormat fdate = new SimpleDateFormat();
                double numValue ;
                switch(cell.getCachedFormulaResultType()){
                    case Cell.CELL_TYPE_NUMERIC:
                        if(HSSFDateUtil.isCellDateFormatted(cell)){
                            return fdate.format(cell.getDateCellValue());
                        }else{
                             numValue = cell.getNumericCellValue();
                             if(numValue==0.0)
                             return "";    
                        }
        
                    case Cell.CELL_TYPE_STRING:
                        return cell.getRichStringCellValue();
                    case Cell.CELL_TYPE_BLANK:
                        return "";
                }
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            case Cell.CELL_TYPE_NUMERIC:
                return cell.getNumericCellValue();
            case Cell.CELL_TYPE_BLANK:
                return "";
            case Cell.CELL_TYPE_ERROR:
                return cell.getErrorCellValue();
        } 
        return null;
    }
}
