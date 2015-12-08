/*
 * The purpose of this is to write all information to a file when submit is pressed.
 */


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.xssf.usermodel.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.*;
import java.util.Date;
import java.util.Map;
import java.util.TreeMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author Brandon Smith
 */
public class InsertIntoExcel {
    

    
    public void newWb(String year) throws IOException{
        
        XSSFWorkbook wb = new XSSFWorkbook(year + " Teen Challenge Urinalysis");
       
    }
    
    public void newSheet(String month, String year) throws IOException{
        
        InputStream input = new FileInputStream(year + " Teen Challenge Urinalysis"+ ".xls");
        try {
            
            Workbook wb = WorkbookFactory.create(input);
            XSSFSheet sheet = (XSSFSheet) wb.createSheet(month);
            Map<String, Object[]> data = new TreeMap<String, Object[]>();
            
             data.put("1", new Object[] {"Date", "First Name", "Last Name", "EtG", "AMP", "OXY", "BUP", "MDMA", "MTD",
             "mAMP", "BZO", "THC", "COT", "Coc", "OPI", "Temp", "k2", "Staff Initials", "Student Signature",
             "Staff Signature", "/n"});
            
        } catch (InvalidFormatException | EncryptedDocumentException ex) {
            Logger.getLogger(InsertIntoExcel.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }
    
    public void insertData(String date, String firstName, String lastName, boolean EtG, boolean AMP, boolean OXY, boolean BUP, boolean MDMA,
            boolean MTD, boolean mAMP, boolean BZO, boolean THC, boolean COT, boolean Coc,
            boolean OPI, boolean Temp, String k2, String staffSig, String staffInitials, String studentSig, 
            String year, String month) throws FileNotFoundException, IOException{
        
        InputStream input = new FileInputStream(year+ " Teen Challenge Urinalysis" + ".xls");
        try {
            
            Workbook wb = WorkbookFactory.create(input);
            XSSFSheet sheet = (XSSFSheet) wb.createSheet(month);
            Map<String, Object[]> data = new TreeMap<String, Object[]>();
            
             data.put("1", new Object[] {date, firstName, lastName, EtG, AMP, OXY, BUP, MDMA, MTD, mAMP, BZO, THC, COT, 
                Coc, OPI, Temp, k2, staffInitials, studentSig, staffSig});
            
        } catch (InvalidFormatException | EncryptedDocumentException ex) {
            Logger.getLogger(InsertIntoExcel.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }
    
}
