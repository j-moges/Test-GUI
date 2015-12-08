/*
 *  The whole panel for the application. Sets pass fail radio buttons and all other 
 *  parts of the application face.
 * 
 */
package bradsapp;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 *
 * @author Brandon Smith
 */
public class RadioPanel extends JPanel {
    
    public RadioPanel(){
        init();

    }
    
    JTextArea fName;
    JTextArea lName;
    JRadioButton EtGTrue;
    JRadioButton EtGFalse;
    JRadioButton AMPTrue;
    JRadioButton AMPFalse;
    JRadioButton OXYTrue;
    JRadioButton OXYFalse;
    JRadioButton BUPTrue;
    JRadioButton BUPFalse;
    JRadioButton MDMATrue;
    JRadioButton MDMAFalse;
    JRadioButton MTDTrue;
    JRadioButton MTDFalse;
    JRadioButton mAMPTrue;
    JRadioButton mAMPFalse;
    JRadioButton BZOTrue;
    JRadioButton BZOFalse;
    JRadioButton THCTrue;
    JRadioButton THCFalse;
    JRadioButton COTTrue;
    JRadioButton COTFalse;
    JRadioButton CocTrue;
    JRadioButton CocFalse;
    JRadioButton OPITrue;
    JRadioButton OPIFalse;
    JRadioButton TempTrue;
    JRadioButton TempFalse;
    JRadioButton k2True;
    JRadioButton k2False;
    JRadioButton k2NA;
    JTextArea initials;
    JTextArea staffSig1;
    JTextArea studentSig1;
    ButtonGroup EtGGrp;
    
    
    public void init(){
        
        
        JPanel panel = new JPanel(new GridLayout(0,4));
        
        
        JLabel name = new JLabel("   First Name: ");
        panel.add(name);
        fName = new JTextArea();
        panel.add(fName);
        JLabel name1 = new JLabel("   Last Name: ");
        panel.add(name1);
        lName = new JTextArea();
        panel.add(lName);
        
        //Pass/Fail
        JLabel space18 = new JLabel();
        panel.add(space18);
        JLabel pass = new JLabel("Pass   ");
        panel.add(pass);
        JLabel fail = new JLabel("  Fail");
        panel.add(fail);
        JLabel na = new JLabel("N/A");
        panel.add(na);     
        
        //EtG checkboxes
        JLabel EtG = new JLabel("EtG");
        panel.add(EtG);
        EtGTrue = new JRadioButton();
        panel.add(EtGTrue);
        EtGFalse = new JRadioButton();
        panel.add(EtGFalse);
        JLabel space4 = new JLabel();
        panel.add(space4);
        EtGGrp = new ButtonGroup();
        EtGGrp.add(EtGTrue); 
        EtGGrp.add(EtGFalse);
        EtGTrue.setSelected(true);
        
        //AMP checkboxes
        JLabel AMP = new JLabel("AMP");
        panel.add(AMP);
        AMPTrue = new JRadioButton();
        panel.add(AMPTrue);
        AMPFalse = new JRadioButton();
        panel.add(AMPFalse);
        JLabel space5 = new JLabel();
        panel.add(space5);
        ButtonGroup AMPGrp = new ButtonGroup();
        AMPGrp.add(AMPTrue); 
        AMPGrp.add(AMPFalse);
        AMPTrue.setSelected(true);
        
        //OXY checkboxes
        JLabel OXY = new JLabel("OXY");
        panel.add(OXY);
        OXYTrue = new JRadioButton();
        panel.add(OXYTrue);
        OXYFalse = new JRadioButton();
        panel.add(OXYFalse);
        JLabel space6 = new JLabel();
        panel.add(space6);
        ButtonGroup OXYGrp = new ButtonGroup();
        OXYGrp.add(OXYTrue); 
        OXYGrp.add(OXYFalse);
        OXYTrue.setSelected(true);
        
        //BUP checkboxes
        JLabel BUP = new JLabel("BUP");
        panel.add(BUP);
        BUPTrue = new JRadioButton();
        panel.add(BUPTrue);
        BUPFalse = new JRadioButton();
        panel.add(BUPFalse);
        JLabel space7 = new JLabel();
        panel.add(space7);
        ButtonGroup BUPGrp = new ButtonGroup();
        BUPGrp.add(BUPTrue); 
        BUPGrp.add(BUPFalse);
        BUPTrue.setSelected(true);
        
        //MDMA checkboxes
        JLabel MDMA = new JLabel("MDMA");
        panel.add(MDMA);
        MDMATrue = new JRadioButton();
        panel.add(MDMATrue);
        MDMAFalse = new JRadioButton();
        panel.add(MDMAFalse);
        JLabel space8 = new JLabel();
        panel.add(space8);
        ButtonGroup MDMAGrp = new ButtonGroup();
        MDMAGrp.add(MDMATrue); 
        MDMAGrp.add(MDMAFalse);
        MDMATrue.setSelected(true);
        
        //MTD checkboxes
        JLabel MTD =  new JLabel("MTD");
        panel.add(MTD);
        MTDTrue = new JRadioButton();
        panel.add(MTDTrue);
        MTDFalse = new JRadioButton();
        panel.add(MTDFalse);
        JLabel space9 = new JLabel();
        panel.add(space9);
        ButtonGroup MTDGrp = new ButtonGroup();
        MTDGrp.add(MTDTrue); 
        MTDGrp.add(MTDFalse);
        MTDTrue.setSelected(true);
        
        //mAMP checkboxes
        JLabel mAMP =  new JLabel("mAMP");
        panel.add(mAMP);
        mAMPTrue = new JRadioButton();
        panel.add(mAMPTrue);
        mAMPFalse = new JRadioButton();
        panel.add(mAMPFalse);
        JLabel space10 = new JLabel();
        panel.add(space10);
        ButtonGroup mAMPGrp = new ButtonGroup();
        mAMPGrp.add(mAMPTrue); 
        mAMPGrp.add(mAMPFalse);
        mAMPTrue.setSelected(true);
        
        //BZO checkboxes
        JLabel BZO =  new JLabel("BZO");
        panel.add(BZO);
        BZOTrue = new JRadioButton();
        panel.add(BZOTrue);
        BZOFalse = new JRadioButton();
        panel.add(BZOFalse);
        JLabel space11 = new JLabel();
        panel.add(space11);
        ButtonGroup BZOGrp = new ButtonGroup();
        BZOGrp.add(BZOTrue); 
        BZOGrp.add(BZOFalse);
        BZOTrue.setSelected(true);
        
        //THC checkboxes
        JLabel THC =  new JLabel("THC");
        panel.add(THC);
        THCTrue = new JRadioButton();
        panel.add(THCTrue);
        THCFalse = new JRadioButton();
        panel.add(THCFalse);
        JLabel space12 = new JLabel();
        panel.add(space12);
        ButtonGroup THCGrp = new ButtonGroup();
        THCGrp.add(THCTrue); 
        THCGrp.add(THCFalse);
        THCTrue.setSelected(true);
        
        //COT checkboxes
        JLabel COT =  new JLabel("COT");
        panel.add(COT);
        COTTrue = new JRadioButton();
        panel.add(COTTrue);
        COTFalse = new JRadioButton();
        panel.add(COTFalse);
        JLabel space13 = new JLabel();
        panel.add(space13);
        ButtonGroup COTGrp = new ButtonGroup();
        COTGrp.add(COTTrue); 
        COTGrp.add(COTFalse);
        COTTrue.setSelected(true);
        
        //Coc checkboxes
        JLabel Coc =  new JLabel("Coc");
        panel.add(Coc);
        CocTrue = new JRadioButton();
        panel.add(CocTrue);
        CocFalse = new JRadioButton();
        panel.add(CocFalse);
        JLabel space14 = new JLabel();
        panel.add(space14);
        ButtonGroup CocGrp = new ButtonGroup();
        CocGrp.add(CocTrue); 
        CocGrp.add(CocFalse);
        CocTrue.setSelected(true);
        
        //OPI checkboxes
        JLabel OPI =  new JLabel("OPI");
        panel.add(OPI);
        OPITrue = new JRadioButton();
        panel.add(OPITrue);
        OPIFalse = new JRadioButton();
        panel.add(OPIFalse);
        JLabel space15 = new JLabel();
        panel.add(space15);
        ButtonGroup OPIGrp = new ButtonGroup();
        OPIGrp.add(OPITrue); 
        OPIGrp.add(OPIFalse);
        OPITrue.setSelected(true);
        
        //Temp checkboxes
        JLabel Temp =  new JLabel("Temp");
        panel.add(Temp);
        TempTrue = new JRadioButton();
        panel.add(TempTrue);
        TempFalse = new JRadioButton();
        panel.add(TempFalse);
        JLabel space16 = new JLabel();
        panel.add(space16);
        ButtonGroup TempGrp = new ButtonGroup();
        TempGrp.add(TempTrue); 
        TempGrp.add(TempFalse);
        TempTrue.setSelected(true);
        
        //k2 checkboxes
        JLabel k2 =  new JLabel("k2");
        panel.add(k2);
        k2True = new JRadioButton();
        panel.add(k2True);
        k2False = new JRadioButton();
        panel.add(k2False);
        k2NA = new JRadioButton();
        panel.add(k2NA);
        ButtonGroup k2Grp = new ButtonGroup();
        k2Grp.add(k2True); 
        k2Grp.add(k2False);
        k2Grp.add(k2NA);
        k2True.setSelected(true);
        
        //text boxes for staff member initials
        JLabel staffInitials = new JLabel("Staff Member Initials: ");
        initials = new JTextArea();
        panel.add(staffInitials);
        panel.add(initials);
        
        //signature boxes
        JLabel staffSig = new JLabel("Staff Signature: ");
        staffSig1 = new JTextArea();
        panel.add(staffSig);
        panel.add(staffSig1);
        
        JLabel studentSig = new JLabel("Student Signature: ");
        studentSig1 = new JTextArea();
        panel.add(studentSig);
        panel.add(studentSig1);
        JLabel space24 = new JLabel();
        panel.add(space24);
        JLabel space25 = new JLabel();
        panel.add(space25);
        
        JButton submit = new JButton("Submit");
        panel.add(submit);
        
        submit.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                
                try {
                    checkForFiles();
                } catch (IOException ex) {
                    Logger.getLogger(RadioPanel.class.getName()).log(Level.SEVERE, null, ex);
                }
 
            }

            
        });

             
        add(panel);
        
        
    }
    
    String SName;
    
    
    public void write() throws IOException, FileNotFoundException{
        
        InsertIntoExcel excel = new InsertIntoExcel();
        SimpleDateFormat df = new SimpleDateFormat("MM/dd/yyyy hh:mm:ss");
        Date date = new Date();
        
        //sets date items for pulling the month and year for insert data
        SimpleDateFormat dfmonth = new SimpleDateFormat("MMMM");
        String month = dfmonth.format(date);
        SimpleDateFormat dfyear = new SimpleDateFormat("yyyy");
        String year = dfyear.format(date);
        
        //creates strings for data in text boxes
        String fname = fName.toString();
        String lname = lName.toString();
        String k2;
        String staffSignature = staffSig1.toString();
        String staffInit = initials.toString();
        String studentSignature = studentSig1.toString();
        
        //because k2 has three options. thanks brad....
        if(k2True.isSelected()){
            k2 = "True";
        }
        else if(k2False.isSelected()){
            k2 = "False";
        }
        else{
            k2 ="N/A";
        }
        
        try{
            excel.insertData(df.format(date), fname, lname, EtGTrue.isSelected(), AMPTrue.isSelected(), OXYTrue.isSelected(),
                    BUPTrue.isSelected(), MDMATrue.isSelected(), MTDTrue.isSelected(), mAMPTrue.isSelected(), 
                    BZOTrue.isSelected(), THCTrue.isSelected(), COTTrue.isSelected(), CocTrue.isSelected(), 
                    OPITrue.isSelected(), TempTrue.isSelected(), k2, staffSignature, staffInit, 
                    studentSignature, year, month);
        }
        catch(IOException ex){
            ex.printStackTrace();
        }
        
    }
    public void checkForFiles() throws IOException{
        
        Date date = new Date();
        SimpleDateFormat dfmonth = new SimpleDateFormat("MMMM");
        String month = dfmonth.format(date);
        SimpleDateFormat dfyear = new SimpleDateFormat("yyyy");
        String year = dfyear.format(date);
        
        boolean workbook = false;
        boolean sheet = false;
        
        InsertIntoExcel excel = new InsertIntoExcel();

        File file = new File(year + " Teen Challenge Urinalysis.xls");
        FileInputStream fis = new FileInputStream(year + " Teen Challenge Urinalysis.xls");
        XSSFWorkbook checkfile = new XSSFWorkbook(fis);
        if(file.exists()){
            workbook = true;
            if(checkfile.getSheetIndex(dfmonth.format(date))>= 0){
            sheet = true;
            write();
            }
        }
                
        else if(workbook == false){
            excel.newWb(year);
            excel.newSheet(month, year);
            write();
        }
        
    }
    
}
