import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

//import com.spire.doc.*;

public class POI {
    public static void process(String data[][], String fileName, String newFileName) throws Exception {
        XWPFDocument doc = new XWPFDocument(OPCPackage.open(fileName));
        for (XWPFParagraph p : doc.getParagraphs()) {
            for (XWPFRun r : p.getRuns()) {
                String textBuffer = r.getText(0);

                if(textBuffer == null){
                        //System.out.println("=empty=");
                        continue;
                }

                for(int i = 0;i < data.length;i++){
                    if(textBuffer.isEmpty()){
                        continue;
                    }

                    if (textBuffer.contains(data[i][0]))
                    {
                        textBuffer = textBuffer.replaceAll(data[i][0], data[i][1]);

                        r.setText(textBuffer, 0);
                        //System.out.println(textBuffer);
                    }
                }
            }
        }
        doc.write(new FileOutputStream(newFileName));

        //Converter.ConvertToPDF(newFileName, newFileName.replaceAll("docx","pdf"));
    }

}