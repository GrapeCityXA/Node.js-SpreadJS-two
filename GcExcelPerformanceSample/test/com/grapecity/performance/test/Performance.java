package test;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.util.Date;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;

public class Performance {

	public static void main(String[] args) {
		System.out.println(System.getProperty("user.dir") + "/sources/jsonData");
		String jsonStr = readTxtFileIntoStringArrList(System.getProperty("user.dir") + "/sources/jsonData");
		JSONArray jsonArr = JSON.parseArray(jsonStr);
		//JSONObject jsonObj = (JSONObject) jsonArr.get(0);
		//System.out.println(jsonObj.get("Film"));
		run(1000, jsonArr);
	}

	public static void run(int times, JSONArray dataArr) {
		String path = System.getProperty("user.dir") + "/results/";
		System.out.println(path + "result.xlsx");
		long start = new Date().getTime();
		for (int i = 0; i < times; i++) {
			Workbook workbook = new Workbook();
			IWorksheet worksheet = workbook.getWorksheets().get(0); 
			for (int j = 0; j < dataArr.size(); j++) {
				JSONObject jsonObj = (JSONObject) dataArr.get(j);
				worksheet.getRange(j, 0, 1, 8).get(0).setValue(jsonObj.get("Film"));
				worksheet.getRange(j, 0, 1, 8).get(1).setValue(jsonObj.get("Genre"));
				worksheet.getRange(j, 0, 1, 8).get(2).setValue(jsonObj.get("Lead Studio"));
				worksheet.getRange(j, 0, 1, 8).get(3).setValue(jsonObj.get("Audience Score %"));
				worksheet.getRange(j, 0, 1, 8).get(4).setValue(jsonObj.get("Profitability"));
				worksheet.getRange(j, 0, 1, 8).get(5).setValue(jsonObj.get("Rating"));
				worksheet.getRange(j, 0, 1, 8).get(6).setValue(jsonObj.get("Worldwide Gross"));
				worksheet.getRange(j, 0, 1, 8).get(7).setValue(jsonObj.get("Year"));
			}
			workbook.save(path + "result" + i + ".xlsx");
		}
		System.out.println("运行"+times+"次花费时常（ms）: " + (new Date().getTime() - start));

	}

	public static String readTxtFileIntoStringArrList(String filePath) {
		StringBuilder list = new StringBuilder();
		try {
			String encoding = "GBK";
			File file = new File(filePath);
			if (file.isFile() && file.exists()) {
				InputStreamReader read = new InputStreamReader(new FileInputStream(file), encoding);// 考虑到编码格式
				BufferedReader bufferedReader = new BufferedReader(read);
				String lineTxt = null;

				while ((lineTxt = bufferedReader.readLine()) != null) {
					list.append(lineTxt);
				}
				bufferedReader.close();
				read.close();
			} else {
				System.out.println("找不到指定的文件");
			}
		} catch (Exception e) {
			System.out.println("读取文件内容出错");
			e.printStackTrace();
		}
		return list.toString();
	}

}
