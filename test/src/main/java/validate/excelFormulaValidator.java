package validate;

import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import dataStructure.ValidGoal;
import msexcel.Excel;
import msexcel.ExcelCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import static common.StringProcessor.*;
import static validate.excelFormulaProcessor.findFormulaForValidate;

public class excelFormulaValidator {

    //possibility:RSQ,D15:D19,
    public static HashSet<Cell> getCellByRef(Excel excel, String keyword) {
        String cellRef = "";
        HashSet<Cell> result = new HashSet<>();
        //D15:D19 ---> D15 D16 D17 D18 D19
        if (Pattern.matches("[A-Z][0-9]+[:][A-Z][0-9]+", keyword)) {
            List<String> cellsAddr = null;
            try {
                cellsAddr = excel.getCellsAddrByRange(keyword);
            } catch (IOException e) {
                e.printStackTrace();
            }
            for (String ref : cellsAddr) {
                result.add(excel.getCell(ref));
            }

        } else {
            if (Pattern.matches("^[$][A-Z][$][0-9]+", keyword)) {
                cellRef = keyword.replaceAll("\\$", "");
            }
            if (Pattern.matches("^[A-Z][0-9]+", keyword)) {
                cellRef = keyword;
            }
            if (!cellRef.isBlank() && excel.getCell(cellRef) != null) {
                result.add(excel.getCell(cellRef));
            }
        }
        return result;
    }

    //公式裡面第一個變數應為 cell的位置 (r1c1)
    //該cell應為純值，否則繼續尋找
    public static HashSet<ExcelCell> getInputCells(Excel excel, Cell outputCell) {
        Cell InputCell = null;
        HashSet<ExcelCell> result = new HashSet<>();
        if (outputCell.getCellType().equals(CellType.FORMULA)) {
            //=IF(ISBLANK(C8),"",RSQ(D15:D19,B15:B19)) ---> RSQ(D15:D19,B15:B19)
            String formulaForValidate = findFormulaForValidate(
                    excel.getCellValue_OriginalFormula(outputCell).toString());
            //RSQ(D15:D19,B15:B19) ---> RSQ D15:D19 B15:B19
            String keywords[] = replaceCustomSymbol(formulaForValidate, " ").split(" ");

            FindInputCellIdx:
            for (String keyword : keywords) {
                HashSet<Cell> inputcells = getCellByRef(excel, keyword);
                for (Cell inputcell : inputcells) {
                    result.add(new ExcelCell(inputcell));
                    result.addAll(getInputCells(excel, inputcell));
                }
            }
        } else {
            return result;
        }
        return result;
    }

    public static HashMap<String, ValidGoal> getValidatedValues(HashMap<String, ValidGoal> prevInfo, Excel expectedExcel) {
        HashMap<String, ValidGoal> result = new HashMap();
        for (Map.Entry<String, ValidGoal> goal : prevInfo.entrySet()) {
            String outputR1C1 = goal.getKey();
            ValidGoal prev = goal.getValue();
            HashSet<ExcelCell> newInputCells = new HashSet<>();

            Cell outputCell = expectedExcel.getCell(outputR1C1);
            Cell inputCell =null;
            if(prev.getInput().getR1c1()!= null)
             inputCell = expectedExcel.getCell(prev.getInput().getR1c1());
            for (ExcelCell c : prev.getAllInputs()) {
                newInputCells.add(new ExcelCell(expectedExcel.getCell(Excel.getR1C1Idx(c.getCell()))));
            }

            ValidGoal newv = new ValidGoal(inputCell, outputCell, prev.getTargetOutput(), newInputCells);
            result.put(outputR1C1, newv);
        }
        return result;
    }
}
