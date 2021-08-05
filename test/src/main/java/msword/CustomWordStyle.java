package msword;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

import static msword.ManageText.*;
public class CustomWordStyle {
    public enum types {
        Title("Title"),Content("Content");
       private String value;
        types(String value){
            this.value = value;
        }
        public String getValue(){
            return this.value;
        }
    }
    public static void addTitle_Style(XWPFParagraph paragraph, String text){
        XWPFRun run = getRun(paragraph);
        getTableTitleStyle(run);
        run.setText(text);
    }
    public static void addContent_Style(XWPFParagraph paragraph, String text){
        XWPFRun run = getRun(paragraph);
        getTableContentStyle(run);
        run.setText(text);
    }
    public static void getTablePrefaceStyle(XWPFRun run, int size){

    }
    public static XWPFRun getTableTitleStyle(XWPFRun run){
        getTableContentStyle(run);
        run.setBold(true);
        return run;
    }
    public static XWPFRun getTableContentStyle(XWPFRun run){
        run.setFontFamily("Arial");
        run.setFontSize(10);
        return run;
    }
}
