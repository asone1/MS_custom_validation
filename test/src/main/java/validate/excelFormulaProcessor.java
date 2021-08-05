package validate;

import msexcel.Excel;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static common.StringProcessor.*;
public class excelFormulaProcessor {
    static final String SPECIFICATION = "specification";
    static final String ERROR = "error";

public static double getTargetOutput(String specification){

    if(specification.contains("%")){
        return Double.valueOf(specification.replaceAll("[^(0-9|.)]", ""))/100;
    }

    return Double.valueOf(specification.replaceAll("[^(0-9|.)]", ""));
}

    /*
        開始找字彙 specification
        尋找邏輯: 找下列的該欄，再下列的首欄開始找，再找下下列以此類推
     */
    public static String findSpecification(Excel excel, int currentRowIdx, int currentCellIdx) {
        for (int tgtRowIdx = currentRowIdx + 1; tgtRowIdx < excel.getLastRowNum(); ++tgtRowIdx) {
            excel.assignRow(tgtRowIdx);

            //找下列的該欄
            if (ifContain(SPECIFICATION, excel.assignCell(currentCellIdx).getCellValue().toString())) {
                return findSpecificationValue(excel, currentCellIdx);
            } else {
                //找再下列的首欄
                for (int tgtCelldx = 0; tgtCelldx < excel.getLastCellNum(); ++tgtCelldx) {
                    if (ifContain(SPECIFICATION, excel.assignCell(tgtCelldx).getCellValue().toString())) {
                        return findSpecificationValue(excel, tgtCelldx);
                    }
                }
            }
        }
        //結束找詞彙

        return "XXXX";
    }

    public static String findSpecificationValue(Excel excel, int currentCellIdx) {
        if (ifEqual(SPECIFICATION, excel.getCellValue().toString())) {
            return excel.findNextCellStrValue();
        } else
            return excel.getCellValue().toString();
    }

    public static boolean checkBlue(Excel excel, Cell cell) {
        CellStyle c = cell.getCellStyle();
        ////For xls (HSSFWorkbook)  or index =12
        if (excel.getWorkbook() instanceof HSSFWorkbook) {
            System.out.println("HSSF:" + ((HSSFWorkbook) excel.getWorkbook()).getInternalWorkbook().getFontRecordAt(c.getFontIndexAsInt()).getColorPaletteIndex());

        }
        //For xlsx (XSSFWorkbook) rgb="FF0000CC"/> AND index =0
        //or index =12
        if (excel.getWorkbook() instanceof XSSFWorkbook) {
            XSSFColor color = ((XSSFWorkbook) excel.getWorkbook()).getFontAt(c.getFontIndexAsInt()).getXSSFColor();
            //arr contains 4 byte --> first one is for index (please ignore)
            byte[] rgb = color.getARGB();
            byte[] blue_1 = {-1, 0, 0, -1};
            byte[] blue_2 = {-1, 0, 0, -52};
            if (Arrays.equals(rgb, blue_1) || Arrays.equals(rgb, blue_2)) return true;
            //調整顏色用
//            for (byte b:rgb){
//                System.out.print(b+",");
//            }System.out.println("");

        }
        return false;
    }

    public static List<Cell> findAllBlueFormulaCall(Excel excel) {
        List<Cell> FormulaCells = new ArrayList<Cell>();

        for (int rowIdx = 0; rowIdx < excel.getLastRowNum(); ++rowIdx) {
            excel.assignRow(rowIdx);

            for (int cellIndx = 0; cellIndx < excel.getLastCellNum(); ++cellIndx) {
                excel.assignCell(cellIndx);

                if (excel.getCurCell().getCellType().equals(CellType.FORMULA)) {
                    if (checkBlue(excel, excel.getCurCell())) {
                        FormulaCells.add(excel.getCurCell());
                    }
                }
            }
        }
        return FormulaCells;
    }

    public static  List<String>  findNonEmptyValueInMultipleParameter(String parameterSignature){
        String paramters [] = parameterSignature.split(",");
        List<String> results = new ArrayList<String>();
        for(String parameter: paramters){
            if (!parameter.replaceAll("\"","").trim().isBlank()){
                results.add(parameter);
            }
        }
        return  results;
    }

    public static  List<String>  removeNonFormulaString (List<String> StringListToCheck){
        List<String> result = new ArrayList<>();
        for(String StringToCheck: StringListToCheck){
            if(!ifContainMethod(StringToCheck)) {
                result.add(StringToCheck);
            }
        }
        return result;
    }

    public static String findFormulaForValidate(String originalFormula){
        if(originalFormula.contains(",")){
            int secondCommaIdx= originalFormula.replaceFirst(","," ").indexOf(",");
            return originalFormula.substring(
                    secondCommaIdx+1
                    , originalFormula.length());
        }
        else return originalFormula;

    }


}
