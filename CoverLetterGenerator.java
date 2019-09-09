package PatientResponsibility;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.io.File;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.text.WordUtils;
import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.multipdf.Splitter;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

@SuppressWarnings("deprecation")
public class PatientRespCoverLetter {
	public static void Body(XWPFDocument word,String name, String dos, int acc, double bill) throws InvalidFormatException, FileNotFoundException, IOException{
		 
		XWPFParagraph paragraph = word.createParagraph();
			XWPFRun run = paragraph.createRun(); 
			    paragraph = word.createParagraph();
			    
				paragraph.setAlignment(ParagraphAlignment.CENTER);
				
				run = paragraph.createRun();
				run.addBreak();
				run.addBreak();
				run.setBold(true);
				run.setUnderline(UnderlinePatterns.SINGLE);
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setText("Accession# " + Integer.toString(acc));
		
				run.addBreak();
				run.addBreak();
				paragraph = word.createParagraph();
				paragraph.setAlignment(ParagraphAlignment.LEFT);
				run = paragraph.createRun();
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setText("Dear ");
				run = paragraph.createRun();
				run.setBold(true);
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setUnderline(UnderlinePatterns.SINGLE);
				run.setText(WordUtils.capitalizeFully(name) + ",");
				run.addBreak(); 
				run.addBreak(); 
				run = paragraph.createRun();
				run.addTab();
				run = paragraph.createRun();
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setText("On the date of ");
				run = paragraph.createRun();
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setBold(true);
				run.setUnderline(UnderlinePatterns.SINGLE);
				run.setText(dos); 
				run = paragraph.createRun();
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setText(", you visited your doctor and your specimens were sent to us for processing.");
				run.addBreak();
				run.addBreak();
				run = paragraph.createRun();
				run.addTab();
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setText("After billing your insurance company, we have received notification that your insurance company applied/partially applied this claim to your");
				run = paragraph.createRun();
				run.setBold(true);
				run.setUnderline(UnderlinePatterns.SINGLE);
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setText(" annual deductible, co-insurance, and/or co-pay");
				run = paragraph.createRun();
				run.setText(" in the amount of ");
				run = paragraph.createRun();
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setBold(true);
				run.setUnderline(UnderlinePatterns.SINGLE);
				run.setText("$");
				run.setText(String.format("%,.2f",bill) + ".");
				run.addBreak();
				run.addBreak();
				run = paragraph.createRun();
				run.addTab();
				run.setBold(true);
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setText("Deductible/Co-insurance/Co-pay ");
				run = paragraph.createRun();
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setText("is the amount ");
				run = paragraph.createRun();
				run.setItalic(true);
				run.setUnderline(UnderlinePatterns.SINGLE);
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setText("you're responsible for paying ");
				run = paragraph.createRun();
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setText("for covered medical expenses before your health insurance plan begins to pay for covered medical expenses each year, and your responsibility amount is determined by your insurance company. ");
				run.addTab();
				run.addBreak();
				run.addBreak();
				run = paragraph.createRun();
				run.setBold(true);
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setText("Federal and State Law requires that all healthcare providers charge patients for their deductible, co-insurance, and co-pay amounts.");
				run.addTab();
				run.addBreak();
				run.addBreak();
				run = paragraph.createRun();
				run.addTab();
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setText("If you have any questions, please feel free to call us at the number below.");
				run.addBreak();
				run.addBreak();
				paragraph = word.createParagraph();
				paragraph.setAlignment(ParagraphAlignment.RIGHT);
				run = paragraph.createRun();
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setText("Thank You,");
				run.addBreak();
				run = paragraph.createRun();
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setText("Rupak Billing Lab Department");
				run.addBreak();
				run = paragraph.createRun();
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setText("Phone: 347-242-8104");
				run.addBreak();
				run.addBreak();
				run = paragraph.createRun();
				run.setBold(true);
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setUnderline(UnderlinePatterns.SINGLE);
				run.setText("PLEASE MAIL PAYMENTS TO");
				run.addBreak();
				run = paragraph.createRun();
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setText("ATTN: Billing Department");
				run.addBreak();
				run.setText("9021 N Indiana Ave");
				run.addBreak();
				run.setText("Fifth Floor");
				run.addBreak();
				run.setText("Lalitpur, Nepal");
				run.addBreak();
				run.setText("Phone: 9843422998");
				run.addCarriageReturn();
				run.addBreak(BreakType.PAGE);	
	}
	
		public static void Extract(XWPFDocument word,FileOutputStream out,PDDocument document) throws InvalidFormatException, FileNotFoundException, IOException {

		      //Instantiate PDFTextStripper class
		      PDFTextStripper pdfStripper = new PDFTextStripper();
		      
		      //Retrieving text from PDF document
		      String text = pdfStripper.getText(document);
		      
		      //Name Extraction
		      String name = StringUtils.substringBetween(text,"Name:", "Clinic Name");
		      name = name.trim();
		      
		      //Accession Extraction
		      String aa = StringUtils.substringBetween(text, "Accession#: ", "Date of Collection");
		      aa = aa.replaceAll("\\s","");
		      int acc = Integer.parseInt(aa);
		      
		      //Date of Service Extraction
		      String dos = StringUtils.substringBetween(text, "Date of Collection: ", "Billing");
		      dos = dos.replaceAll("\\s","");
		      
		      //Bill Amount Extraction
		      String b = StringUtils.substringBetween(text, "Balance Due", "Collection");
		      b = b.replaceAll("\\s","");
		      double bill = Double.parseDouble(b);
		      document.close();
		      Body(word,name, dos, acc, bill);
		      
		}

	public static void HeadFoot(XWPFDocument word, FileOutputStream out) throws InvalidFormatException, FileNotFoundException, IOException {
			
		    XWPFHeaderFooterPolicy headerFooterPolicy = word.getHeaderFooterPolicy();
			  if (headerFooterPolicy == null) headerFooterPolicy = word.createHeaderFooterPolicy();

			  // create header start
			  XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);

			  XWPFParagraph paragraph = header.createParagraph();
			  
			  XWPFRun run = paragraph.createRun(); 
			  paragraph.setAlignment(ParagraphAlignment.CENTER);
			  String imgFile="C:\\Users\\Rupak\\Desktop\\Programs\\PDF\\src\\Rupak.jpg";
			  run.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(380), Units.toEMU(65));
			
			  // create footer start
			  XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);

			  paragraph = footer.createParagraph();
			  paragraph.setAlignment(ParagraphAlignment.CENTER);

			  run = paragraph.createRun();  
			  run.setText("9021 N Indiana Ave • Fifth Floor • Lalitpur, Nepal");
				
	}
	@SuppressWarnings("finally")
	public static void main(String args[]) throws IOException, InvalidFormatException {
		      //Loading an existing document
		int z =0;
		
			String x = "C:\\Users\\Rupak\\";
			new File(x + "Downloads\\BillsAndCoverLetters").mkdirs();
		    String folder_path = 
		    		x + "Downloads\\Bills";
		    File myfolder = new File(folder_path);
		    File[] file_array = myfolder.listFiles(); 
		    
		    XWPFDocument word = new XWPFDocument();
		    
            String wf = x + "Downloads\\AllCovers";
            FileOutputStream out = new FileOutputStream(new File(wf + ".docx"));
            HeadFoot(word,out);
            
		    XWPFDocument word2 = new XWPFDocument();
	   		String wf2 = x + "Downloads\\AllCovers2";
	   		FileOutputStream out2 = new FileOutputStream(new File(wf2 + ".docx"));
	   		HeadFoot(word2,out2);
	   		PDFMergerUtility PDF = new PDFMergerUtility();  
	   		PDFMergerUtility PDF2 = new PDFMergerUtility(); 
		    for (int h = 0; h < file_array.length; h++) 
		    { 

		    	 if (file_array[h].isFile()) 
			        { 
		            File file = new File(folder_path + 
		                     "\\" + file_array[h].getName()); 
		  
		           
		          
		            
					
					  PDDocument document = PDDocument.load(file);
					  if (document.getNumberOfPages() == 1) {
						  
			    		      //adding the source files 
			    		  PDF.addSource(file);  
			    		      
			    		      //Setting the destination file 
			    		  PDF.setDestinationFileName(x + "Downloads\\AllBills.pdf"); 
			    		
			    		      //Merging the two documents 
			    //		  PDF.mergeDocuments(null);
					      
					      
					
					     
					      Extract(word,out, document);
					      
					  }   
			        }
		    }
		    for (int h = 0; h < file_array.length; h++) 
		    { 

		    	 if (file_array[h].isFile()) 
			        { 
		            File file = new File(folder_path + 
		                     "\\" + file_array[h].getName()); 
		  
			   		
			   		
			   		 
					
						  PDDocument document2 = PDDocument.load(file);
						  if (document2.getNumberOfPages() == 2) {
							  
				    		      //adding the source files 
				    		  PDF2.addSource(file);  
				    		      
				    		      //Setting the destination file 
				    		  PDF2.setDestinationFileName(x + "\\AllBills2.pdf"); 
				    		
				    		 
						    
						     Extract(word2, out2 ,document2);
						  }
						  }		 
		    	 }
				   	
		    for (int h = 0; h < file_array.length; h++) 
		    { 

		    	 if (file_array[h].isFile()) 
			        { 
		            File file = new File(folder_path + 
		                     "\\" + file_array[h].getName()); 
		            PDDocument document2 = PDDocument.load(file);
							  if (document2.getNumberOfPages() >= 3) {
							 PDDocument document3 = PDDocument.load(file); 

					            //Instantiating Splitter class
					            Splitter splitter = new Splitter();

					            //splitting the pages of a PDF document
					            List<PDDocument> Pages = splitter.split(document3);

					            //Creating an iterator 
					            Iterator<PDDocument> iterator = Pages.listIterator();

					            //Saving each page as an individual document
					            int i = 0;
					            while(iterator.hasNext()) {
					               PDDocument pd = iterator.next();
					              
					               pd.save(x + "Downloads//BillsAndCoverLetters//BB"+ i++ +".pdf");
					            }
					            
					            document3.close();
					            
					            
					            for (int j = 0; j<=250;j++) {
					            	try {
							            File file5 = new File(x+ "Downloads//BillsAndCoverLetters//BB" + j +".pdf");
							            //adding the source files
							            PDDocument pd = PDDocument.load(file5);
							            PDFTextStripper pdfStripper = new PDFTextStripper();
									      
									      //Retrieving text from PDF document
									     String text = pdfStripper.getText(pd);
									     if (text.contains("Accession#:") && text.contains("Billing Date")) {
									    	 
									    	 PDF.addSource(file5);  
							   		      
									    	 //Setting the destination file 
									    	 PDF.setDestinationFileName(x + "Downloads//AllBills.pdf"); 
							   		
							   		      
									    	 Extract(word,out, pd);
									    	
									     }else {
									    	 PDFMergerUtility ut = new PDFMergerUtility();
									    	 ut.addSource(file5);
									    	 ut.addSource(new File(x+ "Downloads//BillsAndCoverLetters//BB" + (j+1) +".pdf"));
									    	 ut.setDestinationFileName(x + "Downloads//Bills//BB" + z +".pdf");
									    	 ut.mergeDocuments(null);
									         //adding the source files 
									    	 File file6 = new File(x + "Downloads//Bills//BB" + z +".pdf");
									    	 PDDocument pdf = PDDocument.load(file6);
								    		  PDF2.addSource(x + "Downloads//Bills//BB" + z +".pdf");  
								    		      
								    		      //Setting the destination file 
								    		  PDF2.setDestinationFileName(x + "Downloads//AllBills2.pdf"); 
								    		
								    		  Extract(word2,out2, pdf);
								    		  z++;
								    		  j++;
									     }
						            }finally {
						            	continue;
					             }	
							  
					            	}}}}
									     //Merging the two documents 
							    		  PDF2.mergeDocuments(null);
							    		//Merging the two documents 
									   PDF.mergeDocuments(null);
			


		    word.write(out);   
		    word.close(); 
		    out.close();
		    word2.write(out2);
		    out2.close();
			word2.close();       
		    
}}
