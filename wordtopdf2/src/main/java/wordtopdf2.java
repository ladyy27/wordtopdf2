import java.io.*;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
import com.lowagie.text.Row;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import org.apache.poi.ss.usermodel.*;


public class wordtopdf2 {
    public static void main(String[] args) {
        //1 paso: 1 ARCHIVO - convertir a PDF
        wordtopdf2 wordtopdf = new wordtopdf2();
        System.out.println("Inicio");
        wordtopdf.ConvertToPDF("/Users/ela/Desktop/Fgdf.docx", "/Users/ela/Desktop/pdf1.pdf");

        //3 paso: convertir IMG to PDF
        try {
            generatePDFFromImage("/Users/ela/Desktop/maxresdefault.jpg", "/Users/ela/Desktop/maxresdefault.pdf");
            generatePDFFromImage("/Users/ela/Desktop/1.png", "/Users/ela/Desktop/1.pdf");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (DocumentException e) {
            e.printStackTrace();
        }

        //2 paso: JOIN 1ER PDF A 2DO PDF
        List<InputStream> listapdf = new ArrayList<InputStream>();
        try {
            // lista de archivos a unir
            listapdf.add(new FileInputStream(new File("/Users/ela/Desktop/pdf1.pdf")));
            listapdf.add(new FileInputStream(new File("/Users/ela/Desktop/maxresdefault.pdf")));
            listapdf.add(new FileInputStream(new File("/Users/ela/Desktop/1.pdf")));

            // Resulting pdf
            OutputStream salidapdf = new FileOutputStream(new File("/Users/ela/Desktop/pdf_total.pdf"));

            MergePDF(listapdf, salidapdf);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Fin");

    }

    public static void ConvertToPDF(String docPath, String pdfPath) {
        try {
            InputStream doc = new FileInputStream(new File(docPath));
            XWPFDocument documento = new XWPFDocument(doc);
            PdfOptions options = PdfOptions.create();
            OutputStream salida = new FileOutputStream(new File(pdfPath));
            PdfConverter.getInstance().convert(documento, salida, options);

        } catch (FileNotFoundException ex) {
            System.out.println(ex.getMessage());
        } catch (IOException ex) {

            System.out.println(ex.getMessage());
        }
    }

    public static void MergePDF(List<InputStream> list, OutputStream outputStream)
            throws DocumentException, IOException {
        Document document = new Document();
        PdfWriter writer = PdfWriter.getInstance(document, outputStream);
        document.open();
        PdfContentByte cb = writer.getDirectContent();

        for (InputStream in : list) {
            PdfReader reader = new PdfReader(in);
            for (int i = 1; i <= reader.getNumberOfPages(); i++) {
                document.newPage();
                PdfImportedPage page = writer.getImportedPage(reader, i);
                cb.addTemplate(page, 0, 0);
            }
        }
        outputStream.flush();
        document.close();
        outputStream.close();
    }

    public static void generatePDFFromImage(String imgpath, String pdfpath)
            throws IOException, BadElementException, DocumentException {
        Document document = new Document();
        try {
            FileOutputStream fos = new FileOutputStream(pdfpath);
            PdfWriter writer = PdfWriter.getInstance(document, fos);
            writer.open();
            document.open();
            document.newPage();
            Image image = Image.getInstance(Image.getInstance(imgpath));
            image.scalePercent(35);
            document.add(image);
            document.close();
            writer.close();
        }
        catch (Exception e) {
            e.printStackTrace();
        }
    }

}


