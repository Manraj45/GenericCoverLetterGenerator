package CoverLetter;

import java.io.*;
import java.io.FileInputStream;
import java.util.List;
import java.util.Scanner;

import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class CoverLetterGenerator
{
	private static String hiringManager;
	private static String position;
	private static String company;
	private static String name;
	private static String outFilePath;
	private static String inFilePath;
	private static String inFileName;
	private static String docGeneratedPath;

	public static void createCoverLetter() throws FileNotFoundException, IOException
	{
		Scanner scanner = new Scanner(System.in);
		System.out.print("Enter Hiring Manager Name: ");
		hiringManager = scanner.nextLine();
		System.out.print("Enter Position: ");
		position = scanner.nextLine();
		System.out.print("Enter company name: ");
		company = scanner.nextLine();

		name = System.getenv("NAME");
		outFilePath = System.getenv("OUTPUT_FILE_PATH");
		inFilePath = System.getenv("INPUT_FILE_PATH");
		inFileName = System.getenv("INPUT_FILE_NAME");
		docGeneratedPath = outFilePath + "/" + name +"-Cover_Letter-" + company + ".docx";

		XWPFDocument newCoverLetter = new XWPFDocument();
		XWPFDocument genericCoverLetter = new XWPFDocument(new FileInputStream(inFilePath + "/" + inFileName + ".docx"));

		List<XWPFParagraph> paras = genericCoverLetter.getParagraphs();
		for (XWPFParagraph para : paras)
		{
			if (!para.getParagraphText().isEmpty())
			{
				XWPFParagraph newParagraph = newCoverLetter.createParagraph();
				setHiringManagerAndPosition(para, newParagraph);
			}
		}

		newCoverLetter.write(new FileOutputStream(docGeneratedPath));
		convertDOCXToPDF();
	}

	private static void setHiringManagerAndPosition(XWPFParagraph oldParagraph, XWPFParagraph newParagraph)
	{
		for (XWPFRun run : oldParagraph.getRuns())
		{
			String textInRun = run.getText(0);
			System.out.println(textInRun);
			if (textInRun == null || textInRun.isEmpty())
			{
				continue;
			}

			if (textInRun.contains("<hiring manager>"))
			{
				textInRun = textInRun.replace("<hiring manager>", hiringManager);
			}

			if (textInRun.contains("<position>"))
			{
				textInRun = textInRun.replace("<position>", position);
			}
			newParagraph.setAlignment(oldParagraph.getAlignment());
			XWPFRun newRun = newParagraph.createRun();

			// Copy text
			newRun.setText(textInRun);

			// Apply the same style
			newRun.setFontSize(run.getFontSize());
			newRun.setFontFamily(run.getFontFamily());
			newRun.setBold(run.isBold());
			newRun.setItalic(run.isItalic());
			newRun.setStrike(run.isStrike());
			newRun.setColor(run.getColor());
		}
	}

	private static void convertDOCXToPDF()
	{
		try  {
			InputStream docxInputStream = new FileInputStream(docGeneratedPath);
			String pdfGeneratedPath = docGeneratedPath.substring(0,docGeneratedPath.lastIndexOf(".")) + ".pdf";
			OutputStream outputStream = new FileOutputStream(pdfGeneratedPath);
			IConverter converter = LocalConverter.builder().build();
			converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
			outputStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
