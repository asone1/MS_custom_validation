package msword;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

public class ManageText {
    public static XWPFParagraph getParagraph(XWPFTableCell cell){
        if(cell.getParagraphs().size()>0){
            return cell.getParagraphs().get(0);
        }
        else return cell.addParagraph();
    }
    public static XWPFRun getRun(XWPFParagraph paragraph){
        if(paragraph.getRuns().size()>0){
            return paragraph.getRuns().get(0);
        }
        else return paragraph.createRun();
    }

    public static void addNewLines(XWPFDocument doc, int numberOfNewLine) {
        XWPFParagraph newParagraph = doc.createParagraph();
        XWPFRun run = newParagraph.createRun();
        for (int i = 0; i < numberOfNewLine; ++i) {
            run.addCarriageReturn();
        }
    }
    public static void addNewLines(XWPFDocument doc) {
        addNewLines(doc, 1);
    }

}
