package test;

import dataStructure.ValidGoal;
import mainFlow.ProduceWordFile;
import mainFlow.VBS;
import mainFlow.extract;
import msexcel.Excel;
import com.gembox.spreadsheet.*;
import org.apache.poi.ss.usermodel.Cell;
import validate.excelFormulaValidator;

import javax.swing.filechooser.FileSystemView;
import java.io.IOException;
import java.util.*;


/*
[Extract]
    1.找規格(=目標值)
    2.找輸入格
    3.找輸出格
[VBS]
    4.將目標值存入資料表
    5.產出vbs檔案
    6.執行vbs (會執行goal seek產出理想輸入格--->存成新資料表)
[Word]
    7.從新資料表讀入資料表
    8.產生word檔案
 */

/*
重複?
more or less尚未寫
description產生器
最後的word調整
*/

public class ExcelForRu {
    public static final String proj_path = FileSystemView.getFileSystemView().getHomeDirectory().getPath() + "/";

    public static void main(String[] args) throws IOException, InterruptedException {
        List<String> SpecificationList = null;
        HashMap<String, ValidGoal> TobeProcessed = null;
        String fileName = "test1.xlsx";
        Object result = extract.extractData(fileName);
        SpecificationList = (List<String>) ((Object[]) result)[0];
        TobeProcessed = (HashMap<String, ValidGoal>) ((Object[]) result)[1];

//        for(String specification : SpecificationList){
//            System.out.println(specification);
//        }

//                for (Map.Entry<String, ValidGoal> e : TobeProcessed.entrySet()){
//            HashSet<Cell> allin = e.getValue().getAllInputs();
//
//            for(Cell c:allin){
//                System.out.print(Excel.getR1C1Idx(c)+",");
//            }
//                    e.getValue().getInput();
//            System.out.println("____"+e.getValue().getInput());
//        }

        VBS.execVBSFile(VBS.produceVBSFile(fileName,TobeProcessed));
        Thread.sleep(3000);

        String produced_vbsFileName =fileName.split("\\.")[0] + VBS.vbsExcelName + "." + fileName.split("\\.")[1];
        Excel producedExcel =Excel.loadExcel(proj_path + produced_vbsFileName);
        producedExcel.assignSheet(0);
        HashMap<String, ValidGoal> Validated = excelFormulaValidator.getValidatedValues(
                TobeProcessed,producedExcel);
        ProduceWordFile.writeToWord(Validated);
        System.out.println("finish");


    }
}
