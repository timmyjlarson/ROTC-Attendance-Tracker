import java.util.*;
import java.io.File;  
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException; 
import java.util.Iterator;  
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.ss.usermodel.Cell;


/** @author Timothy J Larson
 * 
 * This project is intended for use by AFROTC Detachment 410 for the purpose of creating a automated process for calculating individual cadets objective status
 * 
 * Build requirements: an Java IDE must be downloaded to run this program. For development I used VSCode, but for the purpose of soley running, Eclipse may
 * be more  lightwieght and user friendly. Further, the most recent Java SDK/JRE kit must be downloaded from the Oracle website. All the files included in my 
 * instances lib folder must also be in your projects lib folder, and the spreadsheet either needs to be located in the project folder, or be referenced with an
 * absolute path (I am going to see if I can reference a SharePoint page, but no promises there). Some data sanitation may need to happen on future attendance
 * trackers to make the spreadsheet play nicer with some functionality.
 * 
 * I am willing to update and improve things as needed, but will be presenting code in current state as a proof of concept. If used in any widespread manner,
 * I would seperate these out into different classes and refactor so everything is nicer.
 * 
 * Program restrictions: Cell values must come from the uploaded spreadsheet, not from a link to a different sheet. Multiple sheets may be used via different fis,
 * probably the best COA for scalability
 * 
 * Ideas for future growth: Look into creating front end where current XO could access a remote database for logging rather than a spreadsheet, host this program on
 * a server, and update objectives in real time. Could create a nice dashboard view of objective trends for the OFC, or automate emails out to supervisors regarding
 * one of their cadets missing a PMT event
 */
public class AttendanceDriver{
    public static void main(String[] args) throws IOException{
        //first read in the spreadsheet -> this is dyanmic, so the path can be updated by anyone with ease
        try {  

            //This chunk is for the objective tracker
            File objectiveTracker = new File("Individual Cadet Objective Tracker.xlsx");   //creating a new file instance  ->update this with future paths
            FileInputStream objIn = new FileInputStream(objectiveTracker);   //obtaining bytes from the file  
            //creating Workbook instance that refers to .xlsx file  
            XSSFWorkbook objWorkbook = new XSSFWorkbook(objIn);   
            XSSFSheet llabSchedule = objWorkbook.getSheetAt(1);     //creating a Sheet object to retrieve object 
            Iterator<Row> itr = llabSchedule.iterator();    //iterating over excel file  

            //This chunk is for the attendance tracker
            File attendanceTracker = new File("AttendanceTracker.xlsx"); //->update this with future paths
            FileInputStream attIn = new FileInputStream(attendanceTracker);
            XSSFWorkbook attWorkbook = new XSSFWorkbook(attIn);
            XSSFSheet attSheet = attWorkbook.getSheetAt(0);

            //This chunk is for the output spreadsheet
            /** 
            File cot = new File("CadetObjectTracker.xlsx"); //this is a blank spreadsheet, will be filled in with program output
            FileInputStream cotIn = new FileInputStream(cot);
            XSSFWorkbook cadetObjectiveTracker = new XSSFWorkbook(cot);
            XSSFSheet cadetObjSheet = cadetObjectiveTracker.createSheet("Cadet Objectives");
            */

            //FormulaEvaluator objEvaluator = objWorkbook.getCreationHelper().createFormulaEvaluator(); //this line only necessary if reading formulas, do not recommend

            //Begin IMT Section

            //attempt at adding to a spreadsheet, did not work, revisit later
            /*XSSFRow createRow;
            int addRow = 0;
            createRow = cadetObjSheet.createRow(addRow++);
            int addCell=0;
            Cell cell = createRow.createCell(addCell++);
            cell.setCellValue("IMT Objectives");
            */
            Stack<String> imtObj = new Stack<String>();
            String currentWeekObjectives = "";
            for(int i =1; i< 15; i++){  //these numbers come from the number of LLAbs listed on the spreadsheet
                currentWeekObjectives = llabSchedule.getRow(i).getCell(2).getStringCellValue();
                String[] weeklyObj = currentWeekObjectives.split("," , 0);
                int j = 0;
                while(j< weeklyObj.length){
                    if(!imtObj.contains(weeklyObj[j].trim())){
                        imtObj.push(weeklyObj[j].trim());
                    }
                    j++;
                }
            }
            System.out.println("------ ALL IMT OBJECTIVES ---------");
            System.out.println(imtObj);
            System.out.println();
            System.out.println("------ IMT and their objectives --------");
            currentWeekObjectives = "";
            String currentCadet = "";
            Stack<String> cadetObj = new Stack<String>();
            Stack<String> cadetMissedObj = new Stack<String>();
            for(int i = 45; i< 55; i++){ //these numbers come from start and end indicies of the cadets on the attendence tracker
                currentCadet = attSheet.getRow(i).getCell(0).getStringCellValue();
                cadetObj.clear();
                cadetMissedObj.clear();
                System.out.println(currentCadet);
                for(int j =1; j< 13; j++){ //the number of llabs represented on the attendence sheet
                    currentWeekObjectives = llabSchedule.getRow(j).getCell(2).getStringCellValue();
                    String[] weeklyObj = currentWeekObjectives.split("," , 0);
                    int k = 0;
                    while(k< weeklyObj.length){
                        if((!cadetObj.contains(weeklyObj[k].trim())) && (attSheet.getRow(i).getCell(j+1).getStringCellValue().toUpperCase().equals("X"))){
                            cadetObj.push(weeklyObj[k].trim());
                        }
                        k++;
                    }
                } 
                for(int j = 1; j < imtObj.size(); j++){
                    if(!cadetObj.contains(imtObj.get(j).trim()) && !cadetMissedObj.contains(imtObj.get(j).trim())){
                        cadetMissedObj.push(imtObj.get(j).trim());
                    }
                }
                System.out.println("Objectives Met: ");
                System.out.println(cadetObj); 
                System.out.println(); 
                System.out.println("Objectives Missed: ");
                System.out.println(cadetMissedObj); 
                System.out.println(); 
            }

            //Begin FTP Section

            /*
             * Reference comments on IMT Section for maintence purposes
             */

            Stack<String> ftpObj = new Stack<String>();
            currentWeekObjectives = "";
            for(int i =1; i< 15; i++){
                currentWeekObjectives = llabSchedule.getRow(i).getCell(3).getStringCellValue();
                String[] weeklyObj = currentWeekObjectives.split("," , 0);
                int j = 0;
                while(j< weeklyObj.length){
                    if(!ftpObj.contains(weeklyObj[j].trim())){
                        ftpObj.push(weeklyObj[j].trim());
                    }
                    j++;
                }
            }
            System.out.println("------ ALL FTP OBJECTIVES ---------");
            System.out.println(ftpObj); 
            System.out.println();
            System.out.println("------ FTP and their objectives --------");
            for(int i = 30; i< 44; i++){
                currentCadet = attSheet.getRow(i).getCell(0).getStringCellValue();
                cadetObj.clear();
                cadetMissedObj.clear();
                System.out.println(currentCadet);
                for(int j =1; j< 13; j++){
                    currentWeekObjectives = llabSchedule.getRow(j).getCell(3).getStringCellValue();
                    String[] weeklyObj = currentWeekObjectives.split("," , 0);
                    int k = 0;
                    while(k< weeklyObj.length){
                        if((!cadetObj.contains(weeklyObj[k].trim())) && (attSheet.getRow(i).getCell(j+1).getStringCellValue().toUpperCase().equals("X"))){
                            cadetObj.push(weeklyObj[k].trim());
                        }
                        k++;
                    }
                } 
                for(int j = 1; j < ftpObj.size(); j++){
                    if(!cadetObj.contains(ftpObj.get(j).trim()) && !cadetMissedObj.contains(ftpObj.get(j).trim())){
                        cadetMissedObj.push(ftpObj.get(j).trim());
                    }
                }
                System.out.println("Objectives Met: ");
                System.out.println(cadetObj); 
                System.out.println(); 
                System.out.println("Objectives Missed: ");
                System.out.println(cadetMissedObj); 
                System.out.println(); 
            } 
            
            //Begin POC Section

            /*
             * Reference comments on IMT Section for maintence purposes
             */

            Stack<String> pocObj = new Stack<String>();
            currentWeekObjectives = "";
            for(int i =1; i< 15; i++){
                currentWeekObjectives = llabSchedule.getRow(i).getCell(4).getStringCellValue();
                String[] weeklyObj = currentWeekObjectives.split("," , 0);
                int j = 0;
                while(j< weeklyObj.length){
                    if(!pocObj.contains(weeklyObj[j].trim())){
                        pocObj.push(weeklyObj[j].trim());
                    }
                    j++;
                }
            }
            System.out.println();
            System.out.println("------ ALL POC OBJECTIVES ---------");
            System.out.println(pocObj); 
            System.out.println();
            System.out.println("------ POC and their objectives --------");
            for(int i = 3; i< 29; i++){
                currentCadet = attSheet.getRow(i).getCell(0).getStringCellValue();
                cadetObj.clear();
                cadetMissedObj.clear();
                System.out.println(currentCadet);
                for(int j =1; j< 13; j++){
                    currentWeekObjectives = llabSchedule.getRow(j).getCell(4).getStringCellValue();
                    String[] weeklyObj = currentWeekObjectives.split("," , 0);
                    int k = 0;
                    while(k< weeklyObj.length){
                        if((!cadetObj.contains(weeklyObj[k].trim())) && (attSheet.getRow(i).getCell(j).getStringCellValue().toUpperCase().equals("X"))){
                            cadetObj.push(weeklyObj[k].trim());
                        }
                        k++;
                    }
                } 
                for(int j = 1; j < pocObj.size(); j++){
                    if(!cadetObj.contains(pocObj.get(j).trim()) && !cadetMissedObj.contains(pocObj.get(j).trim())){
                        cadetMissedObj.push(pocObj.get(j).trim());
                    }
                }
                System.out.println("Objectives Met: ");
                System.out.println(cadetObj); 
                System.out.println(); 
                System.out.println("Objectives Missed: ");
                System.out.println(cadetMissedObj); 
                System.out.println(); 
            }             
        }  
        catch(Exception e){  
            e.printStackTrace();  
        }  

    }
}