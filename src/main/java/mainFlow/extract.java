package mainFlow;

import dataStructure.ValidGoal;
import dataStructure.mySheet;
import msexcel.Excel;
import validate.excelFormulaProcessor;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import static test.ExcelForRu.proj_path;

public class extract {
    public static Object[] extractData;

    public static void extractData(String FileName) {

        List<String> SpecificationList = new ArrayList<String>();
        HashMap<String, ValidGoal> TobeProcessed = new HashMap<String, ValidGoal>();
        Excel excel = Excel.loadExcel(proj_path + FileName);


        int sheetSize = excel.getNumberOfSheets();
        for (int a = 0; a < sheetSize; a++) {

            excel.assignSheet(a);

            if (mySheet.checkSheet(excel.getSheet())) {

                mySheet mysheet = new mySheet();

                mysheet = excelFormulaProcessor.collectCellByFontColor(excel, mysheet);


//                List<Cell> allBlueFormulaCall = mysheet.G3;
//
//                //需求一
//                for (Cell formulaCell : allBlueFormulaCall) {
//
//                    String specification = findSpecification(excel, formulaCell.getRowIndex(), formulaCell.getColumnIndex());
//
//                    HashSet<ExcelCell> inputCells = excelFormulaValidator.getInputCells(excel, formulaCell);
//                    ExcelCell inputCell = null;
//                    ValidGoal goal = new ValidGoal();
//                    goal.setOutput(new ExcelCell(formulaCell));
////            goal.setOutput(new ExcelCell(getR1C1Idx(formulaCell),
////                    findFormulaForValidate(excel.getCellValue_OriginalFormula(formulaCell).toString()), formulaCell));
//
//                    for (ExcelCell c : inputCells) {
//                        if (c.getCell().getCellType() != CellType.FORMULA) {
//                            inputCell = c;
//                            break;
//                        }
//                    }
//                    if (inputCell != null) {
//                        goal.setInput(inputCell);
//                    }
//                    goal.setMyComparision(specification);
//                    goal.setAllInputs(inputCells);
//                    SpecificationList.add(specification);
//                    TobeProcessed.put(getR1C1Idx(formulaCell), goal);
//                }
//
////        excel.saveToFile();
//                extractData = new Object[]{SpecificationList, TobeProcessed};
//            }

            }else break;
        }
    }
}
