package mavenweb;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Map;

import javax.servlet.http.Part;
import javax.servlet.ServletException;
import javax.servlet.ServletRequest;
import javax.servlet.annotation.MultipartConfig;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.oreilly.servlet.MultipartRequest;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

//servlet에서는 이거 설정해줘야 multipart 형식으로 들어올 수 있음
@MultipartConfig(fileSizeThreshold = 1024 * 1024, maxFileSize = 1024 * 1024 * 5, maxRequestSize = 1024 * 1024 * 5 * 5)
@WebServlet("/upload")
public class javaservlet extends HttpServlet {

	public static JSONArray convertListToJson(ArrayList<Map<Object, Object>> excelDataList) {
		JSONArray jsonArray = new JSONArray();

		for (Map<Object, Object> map : excelDataList) {
			jsonArray.add(convertMapToJson(map));
		}
		return jsonArray;

	}
	public static JSONObject convertMapToJson(Map<Object, Object> map) {
		JSONObject json = new JSONObject();
		for (Map.Entry<Object, Object> entry : map.entrySet()) {
			Object key = entry.getKey();
			Object value = entry.getValue();
			// json.addProperty(key, value);
			json.put(key, value);
		}
		return json;

	}
	
	@Override
	public void doGet(HttpServletRequest request, HttpServletResponse response) throws IOException {
	}

	@Override
	public void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		System.out.println("post 접근");
		try {
			Part filePart = request.getPart("file");
			// FileInputStream file = new FileInputStream("C:/javademo/Partner.xlsx");
			InputStream file = filePart.getInputStream();
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			
			ArrayList<Map<Object, Object>> excelData = new ArrayList<Map<Object,Object>>();
			Map<Object,Object> propertyObject = new HashMap<Object,Object>();
			ArrayList<Object> header = new ArrayList<Object>();
			XSSFSheet sheet = workbook.getSheetAt(0);
			int rows = sheet.getPhysicalNumberOfRows();
			
			for (int rowNo = 0; rowNo < rows; rowNo++) {
				XSSFRow row = sheet.getRow(rowNo);				
				if (row != null) {
					int cells = row.getPhysicalNumberOfCells();
					for (int cellIndex = 0; cellIndex <= cells; cellIndex++) {
						XSSFCell cell = row.getCell(cellIndex);						
						String value = "";
						if (cell == null) {
							continue;
						} else {
							switch (cell.getCellType()) {
							case XSSFCell.CELL_TYPE_FORMULA:
								value = cell.getCellFormula();
								break;
							case XSSFCell.CELL_TYPE_NUMERIC:
								value = cell.getNumericCellValue() + "";
								break;
							case XSSFCell.CELL_TYPE_STRING:
								value = cell.getStringCellValue() + "";
								break;
							case XSSFCell.CELL_TYPE_BLANK:
								value = cell.getBooleanCellValue() + "";
								break;
							case XSSFCell.CELL_TYPE_ERROR:
								value = cell.getErrorCellValue() + "";
								break;
							}
						}
						if(rowNo==0) {
							header.add(value);							
						}else{
							Object headerProperty = header.get(cellIndex);
							propertyObject.put(headerProperty,value);
						}						
					}
					if(!propertyObject.isEmpty()) {
						excelData.add(propertyObject);
					}	
					propertyObject = new HashMap<Object,Object>();
				}
			}
			response.setContentType("application/json");
			PrintWriter out = response.getWriter();
			out.print(convertListToJson(excelData));			
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
}
