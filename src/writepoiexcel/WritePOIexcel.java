/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package writepoiexcel;

import java.awt.FlowLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.stage.FileChooser;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Lam Thai
 */
public class WritePOIexcel {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws FileNotFoundException, IOException {
        List<POSreport> list;
        ReadExcelFileToList reader =  new ReadExcelFileToList();
          
        String path =pathFile().toString();
        list = reader.readExcelData(path);
        System.out.println(list);
        SimpleDateFormat fdate = new SimpleDateFormat("yyyyMMdd");
        Date date = new Date();
        WriteListToExcelFile writer = new WriteListToExcelFile();
        String name = list.get(1).getSN()+"-"+list.get(1).getSKU()+"-"+(list.size()-1)+"pcs"+"-"+fdate.format(date);
        
        //String Desktop= System.getProperty("user.home")+"/Desktop";
        //String fileName=Desktop+"/"+"/PosReports/"+name+".xlsx";
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Select location to save your POS report");
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        chooser.showSaveDialog(null);
        System.out.println(chooser.getSelectedFile().getAbsolutePath());
        String fileName =chooser.getSelectedFile().getAbsolutePath().toString()+"/"+ name+".xlsx";
        try {
            writer.writeListToFile(fileName, list);
        
        } catch (Exception ex) {
            Logger.getLogger(WritePOIexcel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }      
    
    public static String pathFile(){
            
            JFileChooser chooser = new JFileChooser();
            String fileName="";
            chooser.setDialogTitle("Select 220-xxx file in QA lists");                       
            int result = chooser.showOpenDialog(null);
            if(result == JFileChooser.APPROVE_OPTION){
                File f = chooser.getSelectedFile();
                fileName = f.getAbsolutePath();
                chooser.setVisible(true);
        }
            return fileName;
    }
    
    private static HSSFCellStyle createStyleForTitle(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        return style;
    }
   
       
}

