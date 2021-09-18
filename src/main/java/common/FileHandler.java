package common;

import msexcel.Excel;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.swing.filechooser.FileSystemView;
import java.awt.*;
import java.io.*;
import java.nio.file.Files;
import java.util.*;
import java.util.List;

import static common.StringProcessor.*;
import static msexcel.Excel.getSheetsBykeywordsIgnoreCase;
import static test.ExcelForRu.proj_path;

public class FileHandler {
    public static final String proj_path = FileSystemView.getFileSystemView().getHomeDirectory().getPath() + "/";

    public static boolean moveFileToSubFolder(String path, String subfolderName, File f) {

        File theDir = new File(path + "/" + subfolderName);
        if (!theDir.exists()){
            theDir.mkdirs();
        }
        if (f.exists() && f.isFile()) {
            f.renameTo(new File(path + "/" + subfolderName+ "/"+f.getName()));
//            f.renameTo(new File("C:\\Users\\tina.yu\\Desktop\\biotin_and_Eng\\test1.xlsx"));

            System.out.println("!!!"+path + "/" + subfolderName+ "/"+f.getName());
            return true;
        }
        return false;
    }

    public static void main(String[] args) {


//File ff = new File("C:\\Users\\tina.yu\\Desktop\\RU\\RT30397.xlsx");
//                Excel excel = Excel.loadExcel(ff);
//load檔案後無法移動檔案
//        excel.outputFile("RT30397");
//        File fW = new File("C:\\Users\\tina.yu\\Desktop\\RU\\RT30397.xlsx");
//                moveFileToSubFolder(proj_path,"biotin_and_Eng",fW);
////                System.out.println(excel.getFileName());
//
////                List<String> sheet = getSheetsBykeywordsIgnoreCase(excel,  "micro");
////                System.out.println(sheet.toString());
//                System.out.println("==========");
//            }



//        moveFileToSubFolder (proj_path,"OK",new File ("C:\\Users\\tina.yu\\Desktop\\RU\\H000019703_Biotin & Micro.xlsx"));
//        check every subfolder under input path
//        try {
        List<File> f = getFilesByKeywords("C:\\Users\\YUY139\\Desktop\\09_13_21_CODE", "");
        Excel excelmicro = Excel.loadExcel("C:\\Users\\YUY139\\Desktop\\micro.xlsx");
        excelmicro.assignSheet(0);
        int excelmicroRow =0;
        int count =0;
        for (int r = 0;r< f.size();r++){
            File ff = f.get(r);
            String extension = FilenameUtils.getExtension(ff.getAbsolutePath());
            if (extension.equalsIgnoreCase("xlsx") ||
                    extension.equalsIgnoreCase("xls")) {
                Excel excel = Excel.loadExcel(ff);

    if (getSheetsBykeywordsIgnoreCase(excel, "biotin", "micro").size()>0){

    System.out.println("===");
        System.out.println(count++);
    int excelmicrocol =0;
    excelmicro.assignRow(excelmicroRow);
    List<Object> data= new ArrayList<>();
    data.add(excel.getFileName());

    excelmicro.getCell(excelmicroRow).setCellValue(excel.getFileName());
    excelmicrocol=excelmicrocol++;

    System.out.println(excel.getFileName());


    List<String> sheet = getSheetsBykeywordsIgnoreCase(excel, "biotin", "micro");
    System.out.println(sheet.toString());

     data.add(sheet.toString());
     excel.writeToOneRow (0,r,0,data);
    excelmicroRow= excelmicroRow++;

}
excelmicro.save();
excelmicro.outputFile("C:\\Users\\YUY139\\Desktop\\09_13_21_CODE\\micro_result.xlsx");

            }}
                /*find
                        biotin*/
        for (int r = 0;r< f.size();r++){
            File ff = f.get(r);
            System.out.println("r" + r);
            Excel excel = Excel.loadExcel(ff);
            int sheetSize= excel.getSheetSize();
            for (int i=0; i< sheetSize;i++){
                excel.assignSheet(i);
                int rowSize =excel.getLastRowNum();
                for (int j=0; j< rowSize;j++){
                    excel.assignRow(j);
                    int colSize =excel.getLastCellNum();
                    for (int k=0; k< colSize;k++){
                        excel.assignCell(k);
                        String cellString =excel.getAbsoluteStringCellValue();
                        if (isMatch(cellString, micro_biotin_rex_pre, micro_biotin_rex_suf)) {
                            System.out.println(excel.getFileName());
                            System.out.println("test_item_name:" + cellString);
                            System.out.println("==========");
                        }
                        ;

                    }
                }


                }

//            for (int r = 0;r< f.size();r++){
//                File ff = f.get(r);
//                Excel excel = Excel.loadExcel(ff);
//                String cellAddr = excel.findFirstWordInWb("test item");
//                if (cellAddr==""){
//                    System.out.println("not contain test item"+excel.getFileName());
//                }
//                else {
//                    Cell test_item = excel.getCell(cellAddr);
//                    String test_item_name = excel.findfirstWordAtRight(test_item.getRowIndex(), test_item.getColumnIndex());
//
//
//                    if (isMatch(test_item_name, micro_biotin_rex_pre, micro_biotin_rex_suf)) {
//                        System.out.println(excel.getFileName());
//                        System.out.println("test_item_name:" + test_item_name);
//                        System.out.println("==========");
//                    }
//                    ;
//
//                }

//                FileOutputStream fileOut = new FileOutputStream(ff);
//
//                fileOut.flush();
//                fileOut.close();
//
//                moveFileToSubFolder(proj_path,"biotin_and_Eng",ff);

//                List<String> sheet = getSheetsBykeywordsIgnoreCase(excel, "item", "micro");
//                System.out.println(sheet.toString());

            }}


//        } catch (Exception e) {
//            e.printStackTrace();




    public static List<File> getFilesByKeywords(String dir, String keyword) {
        File f = new File(dir);
        List<File> matchingFiles = new ArrayList();
        if (f.exists() && f.isDirectory()) {
            File arr[] = f.listFiles();
            for (File filesInNestedDir : arr) {
                matchingFiles.addAll(getFilesByKeywords(filesInNestedDir.getAbsolutePath(), keyword));
            }
        } else {
            if (f.getName().contains(keyword)) {
                matchingFiles.add(f);
                return matchingFiles;
            }
        }
        return matchingFiles;
    }

    public static void save(XWPFDocument doc, String pathName) {
        try {
            FileOutputStream out = new FileOutputStream(pathName);
            doc.write(out);
            //Close document
            out.close();
            System.out.println("created:" + "_" + pathName + " written successfully");
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static String getFileNameWoExt(String FileName) {
        return FileName.split("\\.")[0];
    }

    public static void create_save(String fileAbsolutePath, String content) {
        try {
            File theDir = new File(fileAbsolutePath.substring(0, fileAbsolutePath.lastIndexOf("/")));
            if (!theDir.exists()) {
                theDir.mkdirs();
            }
            File file = new File(fileAbsolutePath);
            if (file.createNewFile()) {
                System.out.println("File created: " + file.getName());
            } else {
                System.out.println("File already exists.");
            }
            FileWriter myWriter = new FileWriter(file);
            myWriter.write(content);
            myWriter.close();
            System.out.println("Successfully wrote to the file.");

        } catch (
                IOException e) {
            System.out.println("An error occurred.");
            e.printStackTrace();
        }

    }

    public static String readFromFile(String path) {

        File file = new File(path);
        return readFromFile(file);
    }

    public static String readFromFile(File file) {
        StringBuilder result = new StringBuilder();
        try {
            Scanner myReader = new Scanner(file);
            while (myReader.hasNextLine()) {
//                result.append(myReader.nextLine().replace(" ",""));
                result.append(myReader.nextLine());
            }
            myReader.close();
        } catch (FileNotFoundException e) {
            System.out.println("An error occurred.");
            e.printStackTrace();
        }
        return result.toString();
    }
}
