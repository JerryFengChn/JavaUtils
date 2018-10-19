package com.yumaolin.util.PDFBoxUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.pdfbox.contentstream.operator.Operator;
import org.apache.pdfbox.cos.COSArray;
import org.apache.pdfbox.cos.COSString;
import org.apache.pdfbox.pdfparser.PDFStreamParser;
import org.apache.pdfbox.pdfwriter.ContentStreamWriter;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.pdmodel.PDResources;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.common.PDStream;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotationWidget;
import org.apache.pdfbox.pdmodel.interactive.form.PDAcroForm;
import org.apache.pdfbox.pdmodel.interactive.form.PDTextField;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.util.Charsets;


public class PDFUtil {

	/**
	 * <p>创建PDF文件</p>
	 * @param filePath
	 */
	public void createPDFFile(String filePath) {
		try(PDDocument document = new PDDocument()){
			PDPage page = new PDPage();
			document.addPage(page);
			PDAcroForm acroForm = new PDAcroForm(document);
			document.getDocumentCatalog().setAcroForm(acroForm);
			 // Add and set the resources and default appearance at the form level
			 PDFont formFont = PDType0Font.load(document, new FileInputStream(
	                    "C:\\Windows\\Fonts\\MSYHMONO.ttf"), false);
            final PDResources resources = new PDResources();
            acroForm.setDefaultResources(resources);
            
            final String fontName = resources.add(formFont).getName();
            
            // Acrobat sets the font size on the form level to be
            // auto sized as default. This is done by setting the font size to '0'
            acroForm.setDefaultResources(resources);
            String defaultAppearanceString = "/" + fontName + " 0 Tf 0 g";

            PDTextField textBox = new PDTextField(acroForm);
            textBox.setPartialName("SampleField");
            textBox.setDefaultAppearance(defaultAppearanceString);
            acroForm.getFields().add(textBox);

            // Specify the widget annotation associated with the field
            PDAnnotationWidget widget = textBox.getWidgets().get(0);
            PDRectangle rect = new PDRectangle(50, 700, 200, 50);
            widget.setRectangle(rect);
            widget.setPage(page);
            page.getAnnotations().add(widget);

            // set the field value. Note that the last character is a turkish capital I with a dot,
            // which is not part of WinAnsiEncoding
            textBox.setValue("protocolNo");

            document.save(filePath);
			/*PDFont font = PDType0Font.load(document,new FileInputStream("C:\\Windows\\Fonts\\MSYHMONO.ttf"),false);
			//PDFont font = PDType1Font.COURIER;
			PDPageContentStream content = new PDPageContentStream(document, page);
			content.beginText();
			content.setFont(font, 12);
			content.newLineAtOffset(100, 700);
			content.showText("protocolNo");
			content.endText();
			content.close();*/
			document.save(filePath);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * <p>读取PDF文件</p>
	 * @param filePath
	 */
	public void readPDF(String filePath) {
		PDDocument helloDocument = null;
		try {
			helloDocument = PDDocument.load(new File(filePath));
			PDFTextStripper textStripper = new PDFTextStripper();
			System.out.println(textStripper.getText(helloDocument));
			helloDocument.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * <p>替换PDF文件中的某个字符串</p>
	 * @param inputFile
	 * @param outputFile
	 * @param strToFind
	 * @param message
	 * @throws IOException
	 */
	public void replaceText(String inputFile, String outputFile, String strToFind, String message)
			throws IOException {
		// the document
		 PDDocument document = null;
		try{
			document = PDDocument.load(new File(inputFile));
			//PDFTextStripper stripper=new PDFTextStripper("ISO-8859-1");
			PDPageTree pages = document.getPages();
			PDResources resources = new PDResources();
			PDFont font = PDType0Font.load(document,new FileInputStream("C:\\Windows\\Fonts\\MSYHMONO.ttf"),false);
			resources.add(font);
			for (int i = 0; i < document.getNumberOfPages(); i++) {
				PDPage page = (PDPage) pages.get(i);
				page.setResources(resources);
				PDFStreamParser parser = new PDFStreamParser(page);
				parser.parse();
				List<Object> tokens = parser.getTokens();
				System.out.println(tokens);
				for (int j = 0; j < tokens.size(); j++) {
					Object next = tokens.get(j);
					System.out.println(next);
					if (next instanceof Operator) {
						// Tj and TJ are the two operators that display
						// strings in a PDF
						String operatorName = ((Operator) next).getName();
						if (operatorName.equals("Tj")) {
							// Tj takes one operator and that is the string
							// to display so lets update that operator
							COSString previous = (COSString) tokens.get(j-1);
							byte[] bytes = previous.getBytes();
							String string = new String(bytes,Charsets.UTF_8);
							System.out.println(string);
							System.out.println(previous.getASCII());
							System.out.println(previous.getForceHexForm());
							System.out.println(previous.getCOSObject());
							System.out.println(previous.getString());
							System.out.println(previous.getClass());
							string = string.replaceFirst(strToFind, message);
							previous.setValue(string.getBytes(Charsets.UTF_8));
						} else if (operatorName.equals("TJ")) {
							COSArray previous = (COSArray) tokens.get(j - 1);
							for (int k = 0; k < previous.size(); k++) {
								Object arrElement = previous.getObject(k);
								if (arrElement instanceof COSString) {
									COSString cosString = (COSString) arrElement;
									String string = cosString.getString();
									string = string.replaceFirst(strToFind, message);
									cosString.setValue(string.getBytes(Charsets.UTF_8));
								}
							}
						}
					}
				}
				// now that the tokens are updated we will replace the
				// page content stream.
				PDStream updatedStream = new PDStream(document);
				OutputStream out = updatedStream.createOutputStream();
				ContentStreamWriter tokenWriter = new ContentStreamWriter(out);
				tokenWriter.writeTokens(tokens);
				page.setContents(updatedStream);
				out.close();
			}
			document.save(outputFile);
		}finally{
			if(document != null){
				document.close();
			 }
		}
	}
	
	public static void main(String[] args) throws IOException {
		PDFUtil pdfUtil = new PDFUtil();
		pdfUtil.createPDFFile("D:\\test.PDF");
		pdfUtil.replaceText("D:\\test.PDF","D:\\newTest.PDF","protocolNo","合同编号");
	}
}
