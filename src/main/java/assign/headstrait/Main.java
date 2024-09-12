package assign.headstrait;

import java.io.File;
import java.util.Scanner;

public class Main {
	public static void main(String[] args) throws Exception {
		   Scanner sc;
		   int ch =0 ;
		do {
	            System.out.println("Select the Choice you want :- ");
	            System.out.println("1 : Start Process File ");
	            System.out.println("2 : Delete Processd Files ");
	            System.out.println("3 : Enter 3 for Exit >>");
	            
	            sc = new Scanner(System.in);
	            ch = sc.nextInt();

	            switch (ch) {
	                
	                case 1:
	                   
	                    long startTime =  System.currentTimeMillis();
	                    synchronized (sc) {
	                    	OperationsOfWord.excelConnection();
						} 
	                    OperationsOfWord.executionStart();
	                    System.out.println("Time taken : " + (System.currentTimeMillis() - startTime) + "ms");
	                    break;

	                case 2:
	                   
	                	 File files = new File("./Reprocessed/");
	                     File[] list = files.listFiles();
	                     if(list.length >0 ) {
	                    	 for (File file : list) {
	                    			 file.delete();
	                    	 }	                    	 
	                    	 System.out.println("\n------------------------------ \n Folder Empty successful !! \n ------------------------------ \n");
	                     }else {
	                    	 System.out.println("\n------------------------------ \n Folder IS Empty !!\n ------------------------------ \n");
	                     }
	                    break;
	                 
	                case 3 :
	                	System.out.println("\n------------------------------\n Exited \n ------------------------------\n");
	                	break;

	                default:
	                	System.err.println("\n------------------------------ \n You Enter Wrong Option Try Again!!! \n -----------------------------------\n");
	            }
	        } while ( ch != 3);
		sc.close();
	}
}
