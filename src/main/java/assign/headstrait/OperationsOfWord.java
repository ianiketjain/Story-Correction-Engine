package assign.headstrait;


import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class OperationsOfWord {
static HashMap<String, String> hashMap = new HashMap<String, String>(110);
	
	synchronized public static void excelConnection() throws Exception {
		
		FileInputStream file = new FileInputStream(new File("./ExcelFile/Word Substitutions.xlsx"));
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);

		int i = 0;
		Iterator<Row> rowItr = sheet.iterator();
		while (rowItr.hasNext()) {
			Row row = rowItr.next();

			String key = row.getCell(i).toString();
			String value = row.getCell(i + 1).toString();
			hashMap.put(key, value);
		}
		
		file.close();
		workbook.close();
	}
	
	public static void executionStart() throws Exception {

		//Get the list files in our folder of unprocessed
		File fObj = new File("./Unprocessed");
		File[] files = fObj.listFiles();
		
		 if (files.length != 0) {
	            for (File inputFile : files) {
	                if (inputFile.exists()) {

	                	//creating thread, for each process
	                    Thread processThread = new Thread(new Runnable() {
	                        @Override
	                        public void run() {
	                            try {
										 processFile(inputFile);	
								} catch (Exception e) { e.printStackTrace(); }
	                        }
	                    });
	                    processThread.start();
	                    
	                } else {
	                    System.out.println("FILE DOES NOT EXIST : " + inputFile);
	                }
	            }
	      } else {
	           System.out.println("\t\t FOLDER IS EMPTY \n");
	      }
	}

	private static void processFile(File ss) throws FileNotFoundException, IOException {
		BufferedReader br = new BufferedReader(new FileReader(new File(ss.getAbsolutePath())));
		   StringBuilder builder = new StringBuilder();
				   
		String line;
		while ((line = br.readLine()) != null) {
			
			String[] words = line.split("\\s+");
			String xx="";
			for (int j = 0; j < words.length; j++) {
				if( hashMap.get(words[j].trim()) != null ) {
					xx +=" " + words[j].replace(words[j],hashMap.get(words[j].trim()));
				}else {
					xx += " " +words[j];
				}
			}
			
			builder.append(xx+"\n");
		}
		
		try (FileWriter fileWriter = new FileWriter("./Reprocessed/"+ss.getName())) {
			fileWriter.write(builder.toString());
			fileWriter.flush();
		}
		
	}
}
