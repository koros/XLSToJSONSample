package com.java.sample;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

public class Main {

	public static void main(String[] args) {
		try {
			String file = "FWNewHH.xlsx";
			FileInputStream inp = new FileInputStream(file);//TODO: fix encoding issue here
			Workbook workbook = WorkbookFactory.create(inp);

			// Get the first Sheet.
			Sheet sheet = workbook.getSheetAt(0);

			// Start constructing JSON.
			JSONObject json = new JSONObject();
			JSONObject step = new JSONObject();
			JSONArray titlesRow = new JSONArray();

			// Iterate through the rows.
			JSONArray rows = new JSONArray();
			for (Iterator<Row> rowsIT = sheet.rowIterator(); rowsIT.hasNext();) {
				Row row = rowsIT.next();
				
				// Iterate through the cells.
				JSONObject cells = new JSONObject();
				for (Iterator<Cell> cellsIT = row.cellIterator(); cellsIT.hasNext();) {
					Cell cell = cellsIT.next();
					String cellValue = "";
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_BLANK:
						//cellValue = ""; // already blank anyway
						break;
						
					case Cell.CELL_TYPE_NUMERIC:
						cellValue = String.valueOf(cell.getNumericCellValue());
						break;
						
					case Cell.CELL_TYPE_STRING:
						cellValue = String.valueOf(cell.getStringCellValue());
						break;
						
					case Cell.CELL_TYPE_FORMULA:
						cellValue = String.valueOf(cell.getCellFormula());
						break;
						
					case Cell.CELL_TYPE_BOOLEAN:
						cellValue = String.valueOf(cell.getBooleanCellValue());
						break;
						
					case Cell.CELL_TYPE_ERROR:
						cellValue = String.valueOf(cell.getErrorCellValue());
						break;

					default:
						break;
					}
					if (cell.getRowIndex() == 0) {
						titlesRow.put(cellValue);
					} else {
						String cellKey = titlesRow.getString(cell.getColumnIndex());
						//replace label 
						if (cellKey.equalsIgnoreCase("label::English")) {
							cellKey = "hint";
						}
						//replace type "type": "text"
						if (cellKey.equalsIgnoreCase("type") && cellValue.equalsIgnoreCase("text")) {
							cellValue = "edit_text";
						}
						
						//replace key
						if (cellKey.equalsIgnoreCase("name")) {
							cellKey = "key";
						}
						
						//replace select_one yesno
						if (cellKey.equalsIgnoreCase("type") && cellValue.equalsIgnoreCase("select_one yesno")) {
							cellValue = "radio";
							
							String keyStr = "options";
							Map<String, Object> myMap = new HashMap<String, Object>();
					        myMap.put("yes", "Yes");
					        myMap.put("no", "No");
					        
							JSONArray optionsArray = optionsBuilderFromArray(myMap);
							cells.put(keyStr, optionsArray);
							cells.put("value", "yesno");
						}
						
						if (cellKey != null && !cellKey.isEmpty()) {
							cells.put(cellKey, cellValue);
						}
					}
				}
				
				if (cells.length() != 0) {
					rows.put(cells);
					if (cells.has("hint") && !cells.has("label")) {
						cells.put("label", cells.get("hint"));
					}
				}
			}

			// Create the JSON.
			step.put("fields", rows);
			json.put("count", 1);
			json.put("step1", step);

			//write the string to output file
			PrintWriter writer = new PrintWriter("AndroidJsonFormWizardInput.json", "UTF-8");
			writer.println(json);
			writer.close();

			// Get the JSON text.
			System.out.println(json.toString());
						
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public static JSONArray optionsBuilderFromArray(Map<String, Object> values){
		JSONArray jarray = new JSONArray();
		Iterator<String> it = values.keySet().iterator();
		while (it.hasNext()) {
			String key = (String) it.next();
			Object val = values.get(key);
			JSONObject node = new JSONObject();
			try {
				node.put("key", key);
				node.put("text", val);
			} catch (JSONException e) {
				e.printStackTrace();
			}
			jarray.put(node);
		}
		return jarray;
	}

}
