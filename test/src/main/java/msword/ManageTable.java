package msword;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.math.BigInteger;

import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

public class ManageTable {
    public static void endTable(XWPFTable table) {
        org.apache.xmlbeans.XmlCursor cursor = table.getCTTbl().newCursor();
        cursor.toEndToken();
    }

    public static void fillTable(XWPFTable table) {
        for (int rowIdx = 0; rowIdx < table.getNumberOfRows(); ++rowIdx) {
            XWPFTableRow row = table.getRow(rowIdx);
            for (XWPFTableCell cell : row.getTableCells()) {
                cell.setText("A");
            }
        }
    }

    public static void mergeCellVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            CTVMerge vmerge = CTVMerge.Factory.newInstance();
            if (rowIndex == fromRow) {
                // The first merged cell is set with RESTART merge value
                vmerge.setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                vmerge.setVal(STMerge.CONTINUE);
                // and the content should be removed
                for (int i = cell.getParagraphs().size(); i > 0; i--) {
                    cell.removeParagraph(0);
                }
                cell.addParagraph();
            }
            // Try getting the TcPr. Not simply setting an new one every time.
            CTTcPr tcPr = cell.getCTTc().getTcPr();
            if (tcPr == null) tcPr = cell.getCTTc().addNewTcPr();
            tcPr.setVMerge(vmerge);
        }
    }

    //merging horizontally by setting grid span instead of using CTHMerge
    public static void mergeCellHorizontally(XWPFTable table, int row, int fromCol, int toCol) {
        XWPFTableCell cell = table.getRow(row).getCell(fromCol);
        // Try getting the TcPr. Not simply setting an new one every time.
        CTTcPr tcPr = cell.getCTTc().getTcPr();
        if (tcPr == null) tcPr = cell.getCTTc().addNewTcPr();
        // The first merged cell has grid span property set
        if (tcPr.isSetGridSpan()) {
            tcPr.getGridSpan().setVal(BigInteger.valueOf(toCol - fromCol + 1));
        } else {
            tcPr.addNewGridSpan().setVal(BigInteger.valueOf(toCol - fromCol + 1));
        }
        // Cells which join (merge) the first one, must be removed
        for (int colIndex = toCol; colIndex > fromCol; colIndex--) {
            table.getRow(row).getCtRow().removeTc(colIndex);
            table.getRow(row).removeCell(colIndex);
        }
    }

    public static void setCellW(XWPFTableCell cell, int width) {
        CTTcPr ctTcPr = cell.getCTTc().addNewTcPr();
        CTTblWidth cellWidth = ctTcPr.addNewTcW();
        cellWidth.setW(BigInteger.valueOf(width));
    }

    //default: 5000
    public static void setTableW(XWPFTable table) {
        setTableW(table, 5600);
    }

    public static void setTableW(XWPFTable table, int width) {
        CTTblPr pr = table.getCTTbl().getTblPr();
        CTTblWidth tblW = pr.getTblW();
        tblW.setW(BigInteger.valueOf(width));
        tblW.setType(STTblWidth.PCT);
        pr.setTblW(tblW);
    }

    public static void setValuesOnTable(XWPFTable table, String... values) {
        XWPFTableRow row = table.createRow();
        for (String value : values) {
            row.addNewTableCell().setText(value);
        }
    }
}
