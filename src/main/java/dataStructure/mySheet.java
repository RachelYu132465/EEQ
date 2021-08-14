package dataStructure;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;


public class mySheet {
	List<String> testitemList =new ArrayList<String>();
	String sheetName;

	public List<Cell> G1 = new ArrayList<>();
	public List<Cell> G2 = new ArrayList<>();
	public List<Cell> G3 = new ArrayList<>();

	public static boolean checkSheet (Sheet s){
		if (s.getSheetName().toLowerCase().contains("item")){
			return true;
		}
		return false;
	}
	@Override
	public String toString() {
		return "mySheet{" +
				"testitemList=" + testitemList +
				", sheetName='" + sheetName + '\'' +
				", G1=" + G1 +
				", G2=" + G2 +
				", G3=" + G3 +
				'}';
	}

	public mySheet() {

	}

	public mySheet( String sheetName) {
//		this.testItem = itemName;
//		this.spec = spec;
		this.sheetName = sheetName;
	}
}
