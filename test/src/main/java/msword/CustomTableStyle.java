package msword;

import dataStructure.ValidGoal;
import msexcel.Excel;
import msexcel.ExcelCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xwpf.usermodel.*;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;

import static msword.ManageText.*;
import static msword.ManageTable.*;
import static msword.StrOnCustomWord.*;
import static msword.CustomWordStyle.*;

public class CustomTableStyle {

    //using reflection to invoke static method:
    //addTitle_Style1(table.getRow(2).getCell(0).addParagraph(), INPUT_CELL);
    // <cellValue, cellWidth>
    public static void addToTable(CustomWordStyle.types type, XWPFTable table, int rowIdx, HashMap<String, Integer>... values) {
        int numOfcells = 0;
        for (int realNumOfCell = 0; realNumOfCell < table.getRow(rowIdx).getTableCells().size(); realNumOfCell++) {
            if (table.getRow(rowIdx).getCell(realNumOfCell) != null) ++numOfcells;
        }
        if (numOfcells != values.length) {
            System.out.println("String arg size cannot match current table cells");
            return;
        }

        for (int cellIdx = 0; cellIdx < numOfcells; ++cellIdx) {
            HashMap<String, Integer> cellSetting = values[cellIdx];
            if (cellSetting != null) {
                XWPFTableCell cell = table.getRow(rowIdx).getCell(cellIdx);
                XWPFParagraph paragraph = getParagraph(cell);
                for (Map.Entry<String, Integer> en : cellSetting.entrySet()) {
                    if (en.getValue() != null) setCellW(cell, en.getValue());
                    try {
                        Method method = CustomWordStyle.class.getDeclaredMethod("add" + type.getValue() + "_Style", XWPFParagraph.class, String.class);
                        method.invoke(null, paragraph, en.getKey());
                    } catch (NoSuchMethodException e) {
                        e.printStackTrace();
                    } catch (InvocationTargetException e) {
                        e.printStackTrace();
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                    break;
                }

            }


        }


    }

    public static void addToTable(CustomWordStyle.types type, XWPFTable table, int rowIdx, String... values) {
        int numOfcells = 0;
        for (int realNumOfCell = 0; realNumOfCell < table.getRow(rowIdx).getTableCells().size(); realNumOfCell++) {
            if (table.getRow(rowIdx).getCell(realNumOfCell) != null) ++numOfcells;
        }
        if (numOfcells != values.length) {
            System.out.println("String arg size cannot match current table cells");
            return;
        }
        for (int cellIdx = 0; cellIdx < numOfcells; ++cellIdx) {
            try {
                XWPFTableCell cell = table.getRow(rowIdx).getCell(cellIdx);
                XWPFParagraph paragraph = getParagraph(cell);
                Method method = CustomWordStyle.class.getDeclaredMethod("add" + type.getValue() + "_Style", XWPFParagraph.class, String.class);
                method.invoke(null, paragraph, values[cellIdx]);
            } catch (NoSuchMethodException e) {
                e.printStackTrace();
            } catch (InvocationTargetException e) {
                e.printStackTrace();
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }


    }

    public static HashMap<String, Integer> getHashMap(String cellValue, Integer cellSize) {
        HashMap<String, Integer> result = new HashMap<String, Integer>();
        result.put(cellValue, cellSize);
        return result;
    }

    public static XWPFTable getCustomTable(XWPFDocument doc, int row, int col) {
        XWPFTable table = doc.createTable(row, col);
        table.setCellMargins(100, 50, 90, 5);
        return table;
    }

    //Excel.getCellValue_OriginalFormula(input).toString()
    public static void appendToTable1(XWPFTable table, ValidGoal goal) {
        for (ExcelCell input_c : goal.getAllInputs()) {
            Cell input = input_c.getCell();
            XWPFTableRow row = table.createRow();
            row.createCell();
            row.createCell();
            row.createCell();
            addToTable(types.Content, table, table.getNumberOfRows() - 1,
                    Excel.getR1C1Idx(input)
                    , input.getCellType().toString()
                    , goal.getOutput().getR1c1()
                    , goal.getOutput().getCell().getCellType().toString()
            );
        }
    }

    public static void getTable_Style1(XWPFDocument doc, String worksheetName, String itemName, HashMap<String, ValidGoal> goals) {
        XWPFTable table = getCustomTable(doc, 3, 4);
        table.getRow(1).setHeight(700);
        table.getRow(2).setHeight(600);
        setTableW(table);
        mergeCellHorizontally(table, 0, 0, 3);
        mergeCellHorizontally(table, 1, 0, 3);
        addToTable(types.Title, table, 0, WORKSHEET + Colon + worksheetName);
        addToTable(types.Title, table, 1, ITEM_NAME + Colon + itemName);
        addToTable(types.Content, table, 2, INPUT_CELL, FORMAT_DESC, OUTPUT_CELL, FORMAT_DESC);
        for (Map.Entry<String, ValidGoal> goal : goals.entrySet()) {
            appendToTable1(table, goal.getValue());
        }
        endTable(table);
    }

    public static void appendToTable2(XWPFTable table, HashMap<String, ValidGoal> goals) {
        for (Map.Entry<String, ValidGoal> goal : goals.entrySet()) {
            XWPFTableRow row = table.createRow();
//            row.createCell();
//            row.createCell();
            addToTable(types.Content, table, table.getNumberOfRows() - 1,
                    goal.getValue().getOutput().getCell().getCellFormula(),
                    goal.getValue().getOutput().getR1c1(),
                    goal.getValue().getOutput().getCell().getCellType().toString()
            );
        }
    }

    public static void getTable_Style2(XWPFDocument doc, HashMap<String, ValidGoal> goals) {
        XWPFRun title_run = doc.createParagraph().createRun();
        getTableTitleStyle(title_run);
        title_run.setText(FORMULA + " used");
        XWPFTable table = getCustomTable(doc, 1, 3);
        setTableW(table);//FORMULA, CELL_REF, DESCRIPTION
        addToTable(types.Content, table, 0,
                getHashMap(FORMULA, 5500),
                getHashMap(CELL_REF, 2000),
                getHashMap(DESCRIPTION, 5000));
        appendToTable2(table, goals);
        endTable(table);
    }

    public static void appendToTable3(XWPFTable table, HashMap<String, ValidGoal> goals) {
        for (Map.Entry<String, ValidGoal> goal : goals.entrySet()) {
            XWPFTableRow row = table.createRow();
            row.createCell();
            addToTable(types.Content, table, table.getNumberOfRows() - 1,
                    goal.getValue().getOutput().getR1c1(),
                    "operator:" + goal.getValue().getTargetOutput()
            );
        }
    }

    public static void getTable_Style3(XWPFDocument doc, HashMap<String, ValidGoal> goals) {
        XWPFRun title_run = doc.createParagraph().createRun();
        getTableTitleStyle(title_run);
        title_run.setText("Range validation");
        XWPFTable table = getCustomTable(doc, 2, 2);
        setTableW(table);
        mergeCellHorizontally(table, 0, 0, 1);
        addToTable(types.Content, table, 0, range_valid_txt);
        addToTable(types.Content, table, 1,
                getHashMap(CELL, 1000),
                getHashMap(ACCEPT_R, 6000));
        appendToTable3(table, goals);
        endTable(table);
    }

    public static void appendToTable4(XWPFTable table, ValidGoal goal) {
        for (ExcelCell input_c : goal.getAllInputs()) {
            Cell input = input_c.getCell();
            XWPFTableRow row = table.createRow();
            row.createCell();
            row.createCell();
            row.createCell();
            row.createCell();
            addToTable(types.Content, table, table.getNumberOfRows() - 1,
                    Excel.getR1C1Idx(input)
                    ,input_c.getValue()
                    , goal.getOutput().getR1c1()
                    , goal.getOutput().getValue()
                    ,"",String.valueOf(goal.getTargetOutput()),""
            );
        }
    }

    public static void getTable_Style4(XWPFDocument doc, HashMap<String, ValidGoal> goals) {
        XWPFTable table = getCustomTable(doc, 5, 7);
        table.getRow(0).setHeight(700);
        table.getRow(3).setHeight(600);
        table.getRow(2).setHeight(700);
        setTableW(table);
        mergeCellHorizontally(table, 0, 1, 4);
        mergeCellHorizontally(table, 0, 2, 3);
        mergeCellHorizontally(table, 1, 0, 6);
        mergeCellHorizontally(table, 2, 0, 6);
        mergeCellHorizontally(table, 3, 0, 6);

        addToTable(types.Content, table, 0, "Case", "", "Test Case No.:");
        addToTable(types.Content, table, 1, instruction_txt);
        addToTable(types.Content, table, 2, accep_criteria_txt);
        addToTable(types.Content, table, 3, "in spec");
        addToTable(types.Content, table, 4, INPUT_CELL, "Input value",
                OUTPUT_CELL, "Output " + RESULT, "Calculated " + RESULT, "Expected " + RESULT, "Test " + RESULT);
        for (Map.Entry<String, ValidGoal> goal : goals.entrySet()) {
            appendToTable4(table, goal.getValue());
        }
        endTable(table);
    }

}
