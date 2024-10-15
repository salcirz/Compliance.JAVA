package apch;

// import all of the important xcel file libraries 
import java.io.File;
import java.io.FileInputStream;
import org.apache.logging.log4j.core.tools.picocli.CommandLine.Help.Column;
import org.apache.poi.ss.usermodel.*;
import java.io.FileNotFoundException;
import javax.xml.xpath.XPathExpressionException;
import java.util.Scanner;


//driver class 
public class App 
{
    public static void main(String[] args)
    {
        //call get input
         getInput();

    }



    //this reads all the input
    public static void getInput(){


        //prompt the user to enter the master sheet and the securities sheet as well as the columns
        //needed (prices and securities decription)
        Scanner scan = new Scanner(System.in);
        System.out.println("1:Enter the master sheet name ex: book2.xlsx \n2: press enter \n3:enter the securities sheet name ex: book2.xlsx \n4: press enter \n5: enter the name of the security description column \n6 press enter");
        System.out.println("6: enter name of holding amount column ");
        String mast = scan.nextLine();
        String sec = scan.nextLine();
        String columName = scan.nextLine();
        String pricesName = scan.nextLine();


        //call a new read function for each sheet we want to use
        read currentread = new read(mast,sec, columName, pricesName);


        // find its non compliant securities 
        currentread.findNonCompliant(currentread.getCellorColumnNumber(columName,0), currentread.getCellorColumnNumber(pricesName,0));

        //print them to console and the master sheet 
        currentread.printString();
        currentread.printToMaster(mast);


    }
}