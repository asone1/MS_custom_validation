package mainFlow;

import dataStructure.ValidGoal;
import msexcel.Excel;
import msexcel.ExcelCell;

import javax.swing.filechooser.FileSystemView;
import java.io.*;
import java.util.HashMap;
import java.util.Map;

import static msexcel.Excel.*;
import static test.ExcelForRu.proj_path;

public class VBS {
    public final static String vbsFileName = "_vbsFile";
    public final static String vbsExcelName = "_result";

    public static void execVBSFile(String vbsfileNameWithPath) throws IOException {
        Runtime.getRuntime().exec( "cscript "+ vbsfileNameWithPath );
    }

    public static String saveVBSFile(int index, String inputIdx, String outputIdx, String tgtIdx) {
        return "Dim changingCell" + index + System.getProperty("line.separator") +
                "Dim goal" + index + System.getProperty("line.separator") +
                "Set changingCell" + index + " = objExcel.Range(\"" + inputIdx + "\")" + System.getProperty("line.separator") +
                "Set goal" + index + " = objExcel.Range(\"" + tgtIdx + "\")" + System.getProperty("line.separator") +
                "Call objExcel.Range(\"" + outputIdx + "\").GoalSeek(goal" + index + ", changingCell" + index + ")" + System.getProperty("line.separator");

    }

    public static String produceVBSFile(String FileName, HashMap<String, ValidGoal> TobeProcessed) {

        String fileName_noExt = FileName.split("\\.")[0];

        Excel excel = Excel.loadExcel(proj_path + FileName);
        excel.assignSheet(0);
        int lastRowIdx = excel.getLastRowNum();
        excel.assignRow(lastRowIdx + 1);
        StringBuilder vbsFileContent = new StringBuilder();
        int tgtCellIdx = -1;
        //output r1c1 , temp target cell
        HashMap<String, ExcelCell> allTgt = new HashMap<String, ExcelCell>();
        for (Map.Entry<String, ValidGoal> e : TobeProcessed.entrySet()) {
            ValidGoal data = e.getValue();
            if (data != null) {
                //每個output都要另存目標值在最下面一列
                excel.assignCell(++tgtCellIdx);
                excel.setCellValue(data.getTargetOutput());
//                allTgt.put(getR1C1Idx(data.getOutput().getCell()),
//                        new ExcelCell(
//                                getR1C1Idx(excel.getCurCell()),
//                                String.valueOf(data.getTargetOutput()),
//                                excel.getCurCell()));
                vbsFileContent.append(saveVBSFile(tgtCellIdx, data.getInput().getR1c1(), data.getOutput().getR1c1(), getR1C1Idx(excel.getCurCell())));
//                System.out.println(e.getValue());
            }
        }
        vbsFileContent.insert(0, "Dim objExcel" + System.getProperty("line.separator") +
                "Dim xlSheet" + System.getProperty("line.separator") +
                "Set objExcel = CreateObject(\"Excel.Application\")" + System.getProperty("line.separator") +
                "Set objWorkbook = objExcel.Workbooks.Open(\"" + proj_path + FileName + "\")" + System.getProperty("line.separator")
        );

        vbsFileContent.append("objExcel.ActiveWorkbook.SaveAs \"" + proj_path + fileName_noExt + vbsExcelName + excel.getExcelType().getValue() + "\"\n" +
                "objExcel.ActiveWorkbook.Close\n" +
                "objExcel.Application.Quit");

        try {
            File VBSfile = new File(proj_path+ fileName_noExt + vbsFileName + ".vbs");

            if (VBSfile.createNewFile()) {
                System.out.println("File created: " + VBSfile.getName());
            } else {
                System.out.println("File already exists.");
            }
            FileWriter myWriter = new FileWriter(VBSfile);
            myWriter.write(vbsFileContent.toString());
            myWriter.close();
            System.out.println("Successfully wrote to the file.");
        } catch (
                IOException e) {
            System.out.println("An error occurred.");
            e.printStackTrace();
        }

        excel.saveToFile(FileName);
        return proj_path + fileName_noExt + vbsFileName + ".vbs";
    }

}
