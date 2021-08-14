package mainFlow;

import dataStructure.ValidGoal;
import msword.CustomTableStyle;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.IOException;
import java.util.HashMap;

import static msword.ManageText.addNewLines;

public class ProduceWordFile {
    public static XWPFDocument writeToWord(HashMap<String, ValidGoal> goals, HashMap<String, ValidGoal> newGoals) throws IOException {
        XWPFDocument doc = new XWPFDocument();
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
