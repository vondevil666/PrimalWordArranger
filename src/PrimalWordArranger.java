import org.apache.poi.POITextExtractor;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by dongz on 2017/5/16.
 */
public class PrimalWordArranger {
    private static String inputFilePath = "aWord.docx";

    public static void main(String[] args) {
        PrimalWordArranger arranger=new PrimalWordArranger();
        arranger.startProcess();
    }

    private void startProcess(){
        File inputFile=checkFileExistence(inputFilePath);
        XWPFDocument document = getExtractorInstance(inputFile);
        List<XWPFParagraph> paraList = document.getParagraphs();
        System.out.println(paraList.size()+'\n');
        ArrayList<String> arrayList = new ArrayList<String>();
        StringBuilder sb=new StringBuilder();

        for (XWPFParagraph para : paraList) {
            List<XWPFRun> runList=para.getRuns();
            System.out.println(runList.size()+'\n');
            for (XWPFRun run : runList) {
                if(run.isBold()){
                    arrayList.add(sb.toString());
                    sb=new StringBuilder();
                }
                sb.append(run.text());
            }
        }

        for (String s : arrayList) {
            System.out.println(s+'\n');
        }
    }

    private File checkFileExistence(String filePath){
        File file = new File(filePath);
        if(!file.exists()){
            System.out.println("File Not Found,Check inputFilePath");
            System.exit(1);
        }
        return file;
    }

    private XWPFDocument getExtractorInstance(File inputFile) {
        try {
            return new XWPFDocument(new FileInputStream(inputFile));
        } catch (Exception e) {
            System.out.println("Wrong File Format");
            e.printStackTrace();
        }
        return null;
    }
}
