package common;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import static test.ExcelForRu.proj_path;

public class FileHandler {
    public static List<File> getFilesByKeywords(String dir,String keyword){
        File f = new File(dir);
        List<File> matchingFiles = new ArrayList();
        if(f.exists() && f.isDirectory()){
            File arr[] = f.listFiles();
            for(File filesInNestedDir:arr){
                matchingFiles.addAll(getFilesByKeywords(filesInNestedDir.getAbsolutePath(),keyword));
            }
        }else{
            if( f.getName().contains(keyword)){
                matchingFiles.add(f);
                return matchingFiles;
            }
        }
        return matchingFiles;
    }
    public static void save(XWPFDocument doc, String pathName){
        try {
            FileOutputStream out = new FileOutputStream(pathName);
            doc.write(out);
            //Close document
            out.close();
            System.out.println("createdWord" + "_" + pathName + ".docx" + " written successfully");
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static String getFileNameWoExt(String FileName){
        return FileName.split("\\.")[0];
    }

    public static void create_save(String pathName, String content){
        try {
            File theDir = new File(pathName.substring(0,pathName.lastIndexOf("/")));
            if (!theDir.exists()){
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
