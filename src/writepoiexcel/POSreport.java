/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package writepoiexcel;

import org.apache.poi.ss.usermodel.Cell;

/**
 *
 * @author Kevin
 */
public class POSreport {
    private String name;
    private String firstLine;
    private String SN;
    private String Model;
    private String Company;
    private String DMIrev;
    private String DMIsn;
    private String SKU;

    public POSreport(String name, String firstLine, String SN, String Model, String Company, String DMIrev, String DMIsn, String SKU) {
        this.name = name;
        this.firstLine = firstLine;
        this.SN = SN;
        this.Model = Model;
        this.Company = Company;
        this.DMIrev = DMIrev;
        this.DMIsn = DMIsn;
        this.SKU = SKU;
    }

    

    public POSreport() {
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getFirstLine() {
        return firstLine;
    }

    public void setFirstLine(String firstLine) {
        this.firstLine = firstLine;
    }

    public String getSN() {
        return SN;
    }

    public void setSN(String SN) {
        this.SN = SN;
    }

    public String getModel() {
        return Model;
    }

    public void setModel(String Model) {
        this.Model = Model;
    }

    public String getCompany() {
        return Company;
    }

    public void setCompany(String Company) {
        this.Company = Company;
    }

    public String getDMIrev() {
        return DMIrev;
    }

    public void setDMIrev(String DMIrev) {
        this.DMIrev = DMIrev;
    }

    public String getDMIsn() {
        return DMIsn;
    }

    public void setDMIsn(String DMIsn) {
        this.DMIsn = DMIsn;
    }

    public String getSKU() {
        return SKU;
    }

    public void setSKU(String SKU) {
        this.SKU = SKU;
    }
    public String toString(){
        return SN +";"+Company+";"+Model+";"+DMIrev+";"+DMIsn;
    }
    
}
