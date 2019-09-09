package PDFRenamer;


import java.io.IOException;
import java.util.List;
import java.util.Iterator;
import java.io.File;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.text.WordUtils;
import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.multipdf.Splitter; 

@SuppressWarnings("deprecation")
public class Rename {
	
		@SuppressWarnings({ "finally" })
		public static void main(String[] args) throws IOException {
			String x = "C:\\Users\\Rupak\\Downloads";
		
			
				int i = 0, j =0;   
			  //Loading an existing PDF document
				File file6 = new File(x+ "//AllBills.pdf");
			   	PDDocument document2 = PDDocument.load(file6);
			   	Splitter splitter2 = new Splitter();		  
			   	List<PDDocument> Pages2 = splitter2.split(document2);
			   	Iterator<PDDocument> iterator2 = Pages2.listIterator();
			    int h = 0;
			   while(iterator2.hasNext()) {
			   	PDDocument pd = iterator2.next();
			   	 pd.save(x + "//Bills//A"+ h++ +".pdf");
			   	}
			   	 document2.close();
			   	 
			   	File file10 = new File(x+ "//AllBills2.pdf");
			   	
			   	PDDocument document10 = PDDocument.load(file10);
			   	Splitter splitter10 = new Splitter();		  
			   	List<PDDocument> Pages10 = splitter10.split(document10);
			   	Iterator<PDDocument> iterator10 = Pages10.listIterator();
			    int r = 0;
			   while(iterator10.hasNext()) {
			   	PDDocument pd = iterator10.next();
			   	 pd.save(x + "//Bills//AA"+ r +".pdf");
			   	r++;
			   	}
			   	 document10.close();
			   	 
			    File file = new File(x + "\\AllCovers.pdf");
			    
			    PDDocument document = PDDocument.load(file); 
			
			    //Instantiating Splitter class
			    Splitter splitter = new Splitter();
			
			    //splitting the pages of a PDF document
			    List<PDDocument> Pages = splitter.split(document);
			
			    //Creating an iterator 
			    Iterator<PDDocument> iterator = Pages.listIterator();
			
			    //Loading an existing PDF document
			    File file3 = new File(x + "\\AllCovers2.pdf");
			    if (file3.exists()) {
			    	 PDDocument document3 = PDDocument.load(file3); 
						
					    //Instantiating Splitter class
					    Splitter splitter3 = new Splitter();
					
					    //splitting the pages of a PDF document
					    List<PDDocument> Pages3 = splitter3.split(document3);
					
					    //Creating an iterator 
					    Iterator<PDDocument> iterator3 = Pages3.listIterator();
					
					    //Saving each page as an individual document
					    while(iterator3.hasNext() ) {
						       PDDocument pd2 = iterator3.next();
						       pd2.save(x + "\\BillsAndCoverLetters\\A"+ j +".pdf");
						       j++;
						   }
					    document3.close();
				
			    }
			   
			    
			    //Saving each page as an individual document
			    
			    while(iterator.hasNext()) {
			       PDDocument pd = iterator.next();
			       pd.save(x + "\\BillsAndCoverLetters\\"+ i +".pdf");
			       i++;
			    }
			   
			    
			    document.close();
			    int v = 0;
			    int n =0;
			    for ( v=0; v<=100;v += 2) {
			    	try {
			    File file9 = new File(x+ "//Bills//AA" + v +".pdf");
			    PDFMergerUtility ut = new PDFMergerUtility();
		    	 ut.addSource(file9);
		    	 ut.addSource(new File(x+ "//Bills//AA" + (v+1) +".pdf"));
		    	 ut.setDestinationFileName(x + "//Bills//AAA" + n +".pdf");
		    	 ut.mergeDocuments(null);
		    	 n++;
			    	}finally {
			    		continue;
			    	}
			    }
			    
					   for (int f = 0; f<=1000;f++){
						   try {
						   		File file5 = new File(x+ "//Bills//A" + f +".pdf");
						   	 PDDocument document1 = PDDocument.load(file5);
								      //Instantiate PDFTextStripper class
								      PDFTextStripper pdfStripper = new PDFTextStripper();
								      
								      //Retrieving text from PDF document
								      String text = pdfStripper.getText(document1);
								      
								      //Name Extraction
								      String name = StringUtils.substringBetween(text,"Name:", "Clinic Name");
								      name = WordUtils.capitalizeFully(name.trim());
								      
								      
								      //Accession Extraction
								      String aa = StringUtils.substringBetween(text, "Accession#: ", "Date of Collection");
								      aa = aa.replaceAll("\\s","");
								      int acc = Integer.parseInt(aa);
								      
								      for (int k = 0; k<=250; k++) {
								    	  try {
									    	  String path1 = x + "\\BillsAndCoverLetters\\" + k + ".pdf";
									    	  File file2 = new File(path1);
									    	  document1 = PDDocument.load(file2);
									    	  //Instantiate PDFTextStripper class
									    	  PDFTextStripper pdfStripper2 = new PDFTextStripper();
									      
									    	  //Retrieving text from PDF document
									    	  String text2 = pdfStripper2.getText(document1);
									    	  //Name Extraction
									    	  String name2 = StringUtils.substringBetween(text2,"Dear", ",");
									    	  name2 = WordUtils.capitalizeFully(name2.trim());
									    	  //Accession Extraction
									    	  String aaa = StringUtils.substringBetween(text2, "Accession# ", "Dear");
									    	  aaa = aaa.replaceAll("\\s","");
									    	  int acc2 = Integer.parseInt(aaa);
									    	  
									    	  if (name.equals(name2) && acc == acc2) {
									    		  PDFMergerUtility PDFmerger = new PDFMergerUtility();  
									    		      //adding the source files 
									    		  PDFmerger.addSource(file5); 
									    		  PDFmerger.addSource(file2); 
									    		      
									    		      //Setting the destination file 
									    		  
									    		  PDFmerger.setDestinationFileName(x + "\\" + name + " " + acc + ".pdf"); 
									    		
									    		      //Merging the two documents 
									    		  PDFmerger.mergeDocuments(null);
									    		  break;  
									    	  }
									    	 
									    	  
										    }finally {
										    	
										    	continue;
										    }  
								      	} 	
						      		} finally {
						      			continue;
						      		}
						       
					   }
										   
						   
						   for (int d = 0; d<=1000;d++){
						   try {
						   		File file7 = new File(x+ "//Bills//AAA" + d +".pdf");
						   		PDDocument document1 = PDDocument.load(file7);
						   		
								      //Instantiate PDFTextStripper class
								      PDFTextStripper pdfStripper = new PDFTextStripper();
								      
								      //Retrieving text from PDF document
								      String text = pdfStripper.getText(document1);
								      
								      //Name Extraction
								      String name = StringUtils.substringBetween(text,"Name:", "Clinic Name");
								      name = WordUtils.capitalizeFully(name.trim());
								      
								      
								      //Accession Extraction
								      String aa = StringUtils.substringBetween(text, "Accession#: ", "Date of Collection");
								      aa = aa.replaceAll("\\s","");
								      int acc = Integer.parseInt(aa);
								      
								      for (int k = 0; k<=250; k++) {
								    	  try {
									    	  String path1 = x + "\\BillsAndCoverLetters\\A" + k + ".pdf";
									    	  File file2 = new File(path1);
									    	  document1 = PDDocument.load(file2);
									    	  //Instantiate PDFTextStripper class
									    	  PDFTextStripper pdfStripper2 = new PDFTextStripper();
									      
									    	  //Retrieving text from PDF document
									    	  String text2 = pdfStripper2.getText(document1);
									    	  //Name Extraction
									    	  String name2 = StringUtils.substringBetween(text2,"Dear", ",");
									    	  name2 = WordUtils.capitalizeFully(name2.trim());
									    	  //Accession Extraction
									    	  String aaa = StringUtils.substringBetween(text2, "Accession# ", "Dear");
									    	  aaa = aaa.replaceAll("\\s","");
									    	  int acc2 = Integer.parseInt(aaa);
									    	  
									    	  if (name.equals(name2) && acc == acc2) {
									    		  PDFMergerUtility PDFmerger = new PDFMergerUtility();  
									    		      //adding the source files 
									    		  PDFmerger.addSource(file7); 
									    		  PDFmerger.addSource(file2); 
									    		      
									    		      //Setting the destination file 
									    		  
									    		  PDFmerger.setDestinationFileName(x + "\\" + name + " " + acc + ".pdf"); 
									    		
									    		      //Merging the two documents 
									    		  PDFmerger.mergeDocuments(null);
									    		  break;  
									    	  }
									    	 
									    	  
										    }finally {
										    	
										    	continue;
										    }  
								      	} 			       
						   }finally {
							   continue;
						   }
						   }		   
						   }
		}		    
