import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;

//needed jars: fr.opensagres.poi.xwpf.converter.core-2.0.1.jar,
//             fr.opensagres.poi.xwpf.converter.pdf-2.0.1.jar,
//             fr.opensagres.xdocreport.itext.extension-2.0.1.jar,
//             itext-2.1.7.jar
//needed jars: apache poi and it's dependencies

public class Converter {
    public static void ConvertToPDF(String FileName, String newFileName) throws Exception {

        InputStream in = new FileInputStream(new File(FileName));
        XWPFDocument document = new XWPFDocument(in);
        PdfOptions options = PdfOptions.create();
        OutputStream out = new FileOutputStream(new File(newFileName));
        PdfConverter.getInstance().convert(document, out, options);

        document.close();
        out.close();

    }
}
