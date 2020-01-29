import java.io.*;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfWriter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class wordtopdf2 {
    public static void main(String[] args) {
        //1 paso: 1 ARCHIVO - convertir a PDF
        wordtopdf2 wordtopdf = new wordtopdf2();
        System.out.println("Inicio");
        wordtopdf.ConvertToPDF("/Users/ela/Desktop/libro1.xlsx", "/Users/ela/Desktop/libro1.pdf");


        //2 paso: JOIN 1ER PDF A 2DO PDF
        List<InputStream> listapdf = new ArrayList<InputStream>();
        try {
            // lista de archivos a unir
            listapdf.add(new FileInputStream(new File("/Users/ela/Desktop/ANEXO4-PAGO-FACTURAS-HONORARIOS-DICIEMBRE-2.pdf")));
            listapdf.add(new FileInputStream(new File("/Users/ela/Desktop/pdf2.pdf")));
            listapdf.add(new FileInputStream(new File("/Users/ela/Desktop/pdf3.pdf")));

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
}


