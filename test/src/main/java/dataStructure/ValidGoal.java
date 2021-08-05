package dataStructure;

import msexcel.ExcelCell;
import org.apache.poi.ss.usermodel.Cell;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;

public class ValidGoal {
    ExcelCell input;
    ExcelCell output;
    double targetOutput;
    HashSet<ExcelCell> allInputs;

    public ValidGoal() {
        input = new ExcelCell();
        output = new ExcelCell();
        allInputs = new HashSet<ExcelCell>();
    }

    public ValidGoal(ExcelCell input, ExcelCell output, HashSet<ExcelCell> allInputs) {
        this.input = input;
        this.output = output;
        this.allInputs = allInputs;
    }

    public ValidGoal(Cell input, Cell output, double targetOutput,HashSet<ExcelCell> allInputs) {
        this.input = new ExcelCell(input);
        this.output = new ExcelCell(output);
        this.targetOutput = targetOutput;
        this.allInputs = allInputs;
    }

    public HashSet<ExcelCell> getAllInputs() {
        return allInputs;
    }

    public void setAllInputs(HashSet<ExcelCell> allInputs) {
        this.allInputs = allInputs;
    }

    public ExcelCell getOutput() {
        return output;
    }

    public void setOutput(ExcelCell output) {
        this.output = output;
    }

    public double getTargetOutput() {
        return targetOutput;
    }

    public void setTargetOutput(double targetOutput) {
        this.targetOutput = targetOutput;
    }

    public ExcelCell getInput() {
        return input;
    }

    public void setInput(ExcelCell input) {
        this.input = input;
    }

    @Override
    public String toString() {
        return "input:" + input + System.getProperty("line.separator")
                + "inputs:" + allInputs + System.getProperty("line.separator")
                + "output:" + output + System.getProperty("line.separator")
                + "target:" + targetOutput + System.getProperty("line.separator");
    }
}
