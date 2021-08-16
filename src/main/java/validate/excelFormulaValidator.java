package validate;

import common.FileHandler;
import dataStructure.ValidGoal;
import msexcel.Excel;
import msexcel.ExcelCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.io.File;
import java.io.IOException;
import java.util.*;
import java.util.regex.Pattern;

import static common.StringProcessor.replaceCustomSymbol;
import static validate.excelFormulaProcessor.findFormulaForValidate;

public class excelFormulaValidator {

    public static void main(String[] args) {
        scanForSpecificationOption();
    }

    public static final String Y = "y";

    public static String scanForSpecificationCellAddr() {
        Scanner sc = new Scanner(System.in);
        String CellC1R1 = "";
        System.out.println("請輸入R1C1:");
        while (sc.hasNext()) {
            CellC1R1 = sc.next();
            if (CellC1R1.matches(Excel.CellAddrR1C1Rex)) {
                break;
            }
            System.out.println("格式錯誤，請輸入R1C1");
        }
        System.out.println("輸入儲存格:" + CellC1R1);
        System.out.println("======================");
        return CellC1R1;
    }

    public static String scanForSpecificationOption() {
        Scanner sc = new Scanner(System.in);
        String userInput = "";
        System.out.println("是否手動設定規格?是請輸入:y;否請按enter");

        userInput = sc.nextLine();
        if (userInput.isBlank()) {
            System.out.println("使用預設方式找規格");
            return "";
        } else {
            while (!userInput.matches("[yY]")) {
                System.out.println("是請輸入:y;否請按enter");
                sc.next();
            }
            System.out.println("請輸入規格的儲存格位置(R1C1)");
            return scanForSpecificationCellAddr();
        }
    }

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
            if (Pattern.matches(Excel.CellAddrR1C1Rex, keyword)) {
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

    public static List<Excel> getNewExcels(String dir, String outputR1C1) {
        List<File> files = FileHandler.getFilesByKeywords(dir, outputR1C1);
        List<Excel> excels = new ArrayList<>();
        for (File file : files) {
            if (file.getName().toLowerCase().endsWith("xls") || file.getName().toLowerCase().endsWith("xlsx"))
                excels.add(Excel.loadExcel(file));
        }
        return excels;
    }

    public static HashMap<String, ValidGoal> getValidatedValues(String sheetName, HashMap<String, ValidGoal> prevInfo, String newPath) {
        HashMap<String, ValidGoal> result = new HashMap();
        for (Map.Entry<String, ValidGoal> goal : prevInfo.entrySet()) {
            String outputR1C1 = goal.getKey();
            ValidGoal prev = goal.getValue();
            HashSet<ExcelCell> newInputCells = new HashSet<>();
            if(outputR1C1.equals("C28")){
                System.out.println("!!! ");
            }
            List<Excel> newExcels = getNewExcels(newPath, outputR1C1);
            int tgt_idx =0;

            for (Excel newExcel : newExcels) {
                newExcel.assignSheet(sheetName);
                String outputR1C1withIdx = outputR1C1 +"_"+ tgt_idx++;
                Cell outputCell = newExcel.getCell(outputR1C1);
                Cell inputCell = null;
                if (prev.getInput().getR1c1() != null)
                    inputCell = newExcel.getCell(prev.getInput().getR1c1());
                for (ExcelCell c : prev.getAllInputs()) {
                    newInputCells.add(new ExcelCell(newExcel.getCell(Excel.getR1C1Idx(c.getCell()))).copyNote(c));
                }
                ValidGoal newv = new ValidGoal(inputCell, outputCell, prev.getMyComparision(), newInputCells);
                newv.getOutput().copyNote(prev.getOutput());
                result.put(outputR1C1withIdx, newv);
            }
        }
        return result;
    }
}
