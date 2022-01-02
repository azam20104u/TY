package wordformat.com.pooja;
//Java Program to format the text
		// in a word document

		// Importing input output java files
		import java.io.*;

import org.apache.poi.xwpf.usermodel.*;

		// Class to format text in Word file
public class App {
	public static void main(String[] args) throws IOException {
				// Step 1: Creating blank document
				XWPFDocument document = new XWPFDocument();

				// Step 2: Getting path of current working directory
				// to create the pdf file in the same directory
				String path = System.getProperty("D:\\test\\Doc1.docx");

				// Step 3: Creating a file object with the path
				// specified
				FileOutputStream out
					= new FileOutputStream(new File("D:\\test\\Doc1.docx"));

				// Step 4: Create paragraph
				// using createParagrapth() method
				XWPFParagraph paragraph
					= document.createParagraph();

				// Step 5: Formatting lines

				// Line 1
				// Creating object for line 1
				XWPFRun line1 = paragraph.createRun();

				// Formating line1 by setting bold
				line1.setBold(true);
				line1.setText("Formatted with Bold");
				line1.addBreak();

				// Line 2
				// Creating object for line 2
				XWPFRun line2 = paragraph.createRun();

				// Formating line1 by setting italic
				line2.setText("Formatted with Italics");
				line2.setItalic(true);
				line2.addBreak();

				// Line 3
				// Creating object for line 3
				XWPFRun line3 = paragraph.createRun();

				// Formatting line3 by setting
				// color & font size
				line3.setColor("73fc03");
				line3.setFontSize(20);
				line3.setText(" Formatted with Color");

				// Step 6: Saving changes to document
				document.write(out);

				// Step 7: Closing the connections
				out.close();
				document.close();

				// Print message after program has compiled
				// successfully showcasing formatting text in file
				// successfully.
				System.out.println("Word Document with Formatted Text created successfully!");
	}
}
