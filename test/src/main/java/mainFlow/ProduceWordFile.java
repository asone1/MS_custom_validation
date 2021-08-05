package mainFlow;

import dataStructure.ValidGoal;
import msword.CustomTableStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;

import static msword.ManageTable.*;
import static msword.ManageText.*;

public class ProduceWordFile {
    public static void writeToWord(HashMap<String, ValidGoal> goals) throws IOException {
        XWPFDocument doc = new XWPFDocument();
        CustomTableStyle.getTable_Style1(doc,"","",goals);
        addNewLines(doc);

        CustomTableStyle.getTable_Style2(doc,goals);
        addNewLines(doc);

        CustomTableStyle.getTable_Style3(doc,goals);
        addNewLines(doc);

        CustomTableStyle.getTable_Style4(doc,goals);
        FileOutputStream out = new FileOutputStream("final result.docx");
        doc.write(out);
        out.close();
    }
}
