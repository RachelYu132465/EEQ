package mainFlow;

import dataStructure.ValidGoal;
import msword.CustomTableStyle;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.IOException;
import java.util.HashMap;

import static msword.ManageText.addNewLines;

public class ProduceWordFile {
    public static XWPFDocument writeToWord_general(String sheetName,XWPFDocument doc ,HashMap<String, ValidGoal> goals) throws IOException {

        CustomTableStyle.getTable_Style1(doc,sheetName,"",goals);
        addNewLines(doc);

        CustomTableStyle.getTable_Style2(doc,goals);
        addNewLines(doc);

        CustomTableStyle.getTable_Style3(doc,goals);
        addNewLines(doc);
        return doc;
    }

    public static XWPFDocument writeToWord_testCase(XWPFDocument doc ,HashMap<String, ValidGoal> goals, HashMap<String, ValidGoal> newGoals) throws IOException {

        CustomTableStyle.getTable_Style1(doc,"","",goals);
        addNewLines(doc);

        CustomTableStyle.getTable_Style2(doc,goals);
        addNewLines(doc);

        CustomTableStyle.getTable_Style3(doc,goals);
        addNewLines(doc);

        CustomTableStyle.getTable_Style4(doc,goals,newGoals);
        return doc;
    }
}
