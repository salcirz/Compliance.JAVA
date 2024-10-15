package apch;

// import all necessary libraries 
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.ArrayList;
import java.util.HashMap;

public class read {
    

// we have out master sheet seuriteis sheet 5 percent identifier, column indexes, total assets and master
//book variables for each read 
private Sheet masterSheet;
private Sheet securitiesSheet;
private double f5percent;
private int nameColumnIndex;
private int pricesColumnIndex; 
private double totalAssets;
private Workbook masterBook;


// I created two array lists to hold the totals and their names 
private ArrayList<Double> totals = new ArrayList<Double>();
private ArrayList<String> names = new ArrayList<String>();


//constructor 

// each read takes in the two sheets and two columns for name and price
public read (String master, String securities, String description, String prices){

    try{

        // create the file stram and catch the exception if there is none 
        FileInputStream file = new FileInputStream(new File(master));
        this.masterBook = WorkbookFactory.create(file);

        // set the master sheet 
        this.masterSheet = this.masterBook.getSheetAt(0);
    }catch (IOException m){
        m.printStackTrace();
    }
    

    //set securiteis sheet
    this.securitiesSheet = fileInputStream(securities, 0);

    //find the total assets and also how much 5% of that is 
    this.totalAssets = getValue(4, 1, this.masterSheet);
    this.f5percent = totalAssets *.05; 

    //set it on the master sheet
    Cell five = this.masterSheet.getRow(5).createCell(1);
    five.setCellValue(this.f5percent);

    //set the name and prices columns 
    this.nameColumnIndex = getCellorColumnNumber(description, 0);
    this.pricesColumnIndex = getCellorColumnNumber(prices, 0);



}
// return constructors 
public int getcolumnindexprices(){
return this.pricesColumnIndex;
}
public int getColumninname(){
    return this.nameColumnIndex;
}
public Sheet getSecuritiesSheet(){
    return this.securitiesSheet;
}
public double get5(){
    return this.f5percent;
}
//


//function to print the names then the totals for each based on the array lists 
public void printString(){

    for(int i= 0; i < this.totals.size(); i ++){
        System.out.print(this.names.get(i) + "\t");
        System.out.println(this.totals.get(i));
    }
}


// function to print our answers to the master excel sheet 
public void printToMaster(String mast){

    
    // starting at row 11 that is just how i formatted the master sheet which can be
    // copied for multiple firms. also storing the total value and percent of fund.
    int row = 11;
    double totalpercent = 0;
    double totalvalue = 0;

    //for the size of array list 
    for(int i= 0; i < this.totals.size(); i ++){


        //get that row 
        Row cellRow = this.masterSheet.createRow(row);
         
        // set the name total and percent cells
        Cell name = cellRow.createCell(1);
        Cell total = cellRow.createCell(3);
        Cell percent = cellRow.createCell(5);

        // set each corresponding to the array list values 
        name.setCellValue(this.names.get(i));
        total.setCellValue(this.totals.get(i));
        percent.setCellValue(this.totals.get(i)/ this.totalAssets * 100.00);


        //update the total and total percent
        totalvalue += this.totals.get(i);
        totalpercent += percent.getNumericCellValue();

        //we want to skip two rows each time for formatting 
        row+=2;


    }

    // get the total row which is the last row after all is printed
    Row totalRow = this.masterSheet.createRow(row);
         

    //get total and total percent cells and set them to our values 
    Cell TT = totalRow.createCell(3);
    Cell TP = totalRow.createCell(5);
    TT.setCellValue(totalvalue);
    TP.setCellValue(totalpercent);
    
    //make the outsteam and write it to our master file or print error
    try(FileOutputStream outFile = new FileOutputStream(new File(mast))){
        this.masterBook.write(outFile);
    }catch(IOException p){
        p.printStackTrace();
    }
       

}


//returns the index of a column with the identifier value based on row and sheet
public int getCellorColumnNumber(String identifier , int row){

        //gets the specified row based on index 
        Row theRow = this.securitiesSheet.getRow(row);

        //for the whole row
        for(int i = 0; i <theRow.getLastCellNum(); i++){
            
            //get each cell and check if theres a value
            Cell theCell = theRow.getCell(i);
            if(theCell ==null){

                continue;
            }

            //if its equal to what column we are looking for we save that value. we are trying
            //to find the column index of the names and security value sheets.
            if(theCell.getStringCellValue().equals(identifier)){
                return i;
            }
        }
        //if we cant find it return -1 to get exception
         return -1;

}


// function to find and save non compliant 
public void findNonCompliant(int nameColumnIndex, int pricesColumnIndex){


    HashMap <String, String> identifiers = new HashMap<>();
    HashMap <Double, String> TotalsNames= new HashMap<>();
    String currentIdentifier; 

    //for the totals column: 
    for(int i= 1; i< this.securitiesSheet.getLastRowNum()+1; i ++ ){
       
        //we save the total as 0 and our current identifier to what cell were on
        double total =0;
        currentIdentifier = getString(i, nameColumnIndex, this.securitiesSheet);
        
        // if the hash map already contains this value we know to skip it because we already checked
        //it, if not we now check the rest of the list against this identifier value
        if(identifiers.containsValue(currentIdentifier)){
            continue;

        } else {

            //set the current index to the current value and check the rest of the row 
            for(int j = i; j < this.securitiesSheet.getLastRowNum()+1; j ++){

                //if its empty continue 
                if(getString(j, nameColumnIndex, this.securitiesSheet)== null){
                    continue;
                }
            
                //if its equal to our current identifier then we add its value to the total.
                //there can be multiple securities like google a and google but it still
                //falls under google which is why its complicated 
                if(getString(j, nameColumnIndex, this.securitiesSheet).equals(currentIdentifier)){
                    
                    total+=getValue(j,pricesColumnIndex,this.securitiesSheet);
                }

             }
             //then we put our identifier in the hash map to save which names we already checked in
             //case there are dupliactes
             identifiers.put(currentIdentifier, currentIdentifier);

             //we now check if the total is over the threshold 
             if(total > this.get5()){

            // if so we add the total to our arraylist and hash map
             this.totals.add(total);
             TotalsNames.put(total,currentIdentifier);

             }

             
        

        }




    }


    //now we can sort the totals using the array list sorts class
    Collections.sort(totals, Collections.reverseOrder());
    
    
    //because we have the names and totals of the ones over the 5% stored in our hash map,
    // we can go through the now sorted totals list and find the corresponding name to the total
    //so the names list is now also sorted for easier organization.
    for(int k = 0; k< totals.size(); k++){
        this.names.add(TotalsNames.get(totals.get(k)));

    }

}




//check if number is less than 5%
public boolean lessThan(double TotalHeld){

    if(TotalHeld> this.f5percent){
        return true;
    }else {
        return false;
    }
}


// initialize the master sheet and securitites sheet 
public Sheet fileInputStream(String bookname, int index){

    try{
        //broke down this code into a function so the calls are easier 
        FileInputStream excelInput =  new FileInputStream(new File(bookname));
        Workbook excelBook = WorkbookFactory.create(excelInput);
        return excelBook.getSheetAt(index);

    }catch(Exception inputExcel){
    
        inputExcel.printStackTrace();
    }
    return null;

}


//get a numeric value within a cell 
public double getValue(int row, int column, Sheet mySheet){
    

    //get the row and cell and return its value
    Row myRow = mySheet.getRow(row);
    Cell cell = myRow.getCell(column);

    double value = cell.getNumericCellValue();

    return value;


} 

// get a string cell value 
public String getString(int row, int column, Sheet mySheet){


// get the row and cell and check if its null
Row myRow = mySheet.getRow(row);
if(myRow== null){
    return null;
}
Cell cell = myRow.getCell(column);

if(cell == null){
    return null;
}

//get the value and retrn 
String myValue = cell.getStringCellValue();

return myValue;

}

//get a cell, check if null and return it if not 
//used to get any cell during the program
public Cell getCell(int row, int column, Sheet mySheet){
    
    Row myRow = mySheet.getRow(row);
    if(myRow== null){
        return null;
    }
    Cell cell = myRow.getCell(column);
    return cell; 
    

}
//

}
