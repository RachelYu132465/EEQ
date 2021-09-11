package common;

import msexcel.Excel;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;

import static msexcel.Excel.getSheetsBykeywordsIgnoreCase;
import static test.ExcelForRu.proj_path;

public class FileHandler {
    public static void main(String[] args) {
        //check every subfolder under input path

        List<File> f = getFilesByKeywords("C:\\Users\\tina.yu\\Desktop\\RU", "");
        for (File ff : f) {
            String extension = FilenameUtils.getExtension(ff.getAbsolutePath());
            if (extension.equalsIgnoreCase("xlsx") ||
                    extension.equalsIgnoreCase("xls")) {
                Excel excel = Excel.loadExcel(ff);
                System.out.println(excel.getFileName());

                List<String> sheet =getSheetsBykeywordsIgnoreCase (excel,"item","micro");
                System.out.println(sheet.toString());
                System.out.println("==========");
            }

        }
    }


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

    public static void create_save(String pathName, String content) {
        try {
            File theDir = new File(pathName.substring(0, pathName.lastIndexOf("/")));
            if (!theDir.exists()) {
                theDir.mkdirs();
            }
            File file = new File(pathName);
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
}
