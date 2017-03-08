package com.ayg.poiutils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.whl.poi.ExcelRSheetProcessor;
import com.whl.poi.ExcelRead;
import com.whl.poi.ExcelType;
import com.whl.poi.ExcelWSheetProcessor;
import com.whl.poi.ExcelWrite;
import com.whl.poi.demo.UserInfo;

public class App {
	public static void main(String[] args) throws Exception {
		read();
		write();
	}
	
	private static void read() throws Exception{
		ExcelRSheetProcessor<UserInfo> sheetProcessor = new ExcelRSheetProcessor<UserInfo>() {
			@Override
			public void beforeProcess() {
				// TODO Auto-generated method stub
			}
			@Override
			public void process(List<UserInfo> list) {
				
				for (UserInfo info : list) {
					System.out.println("id=" + info.getId() + " name=" + info.getName() + " cardNo=" + info.getCardNo()
							+ " salary=" + info.getSalary() + " personInsTitle=" + info.getPersonInsTitle()+ " firmInsTitle=" + info.getFirmInsTitle()
							+ " personIns=" + info.getPersonIns() + " firmIns=" + info.getFirmIns());
				}
			}
			@Override
			public void afterProcess() {
				// TODO Auto-generated method stub
			}
			@Override
			public void exceptionHandler(Exception e) {
				e.printStackTrace();
			}
		};
		sheetProcessor.setRowStartIndex(1);
		sheetProcessor.setPageSize(7);
		final String filePath = "F:\\tmp\\salary.xls";
		File file = new File(filePath);
		FileInputStream input = new FileInputStream(file);
		ExcelRead.read(input, sheetProcessor);
	}
	
	
	private static void write() throws Exception{
		final List<UserInfo> list = new ArrayList<UserInfo>();
		for (int i = 0; i < 5; i++) {
			UserInfo user1 = new UserInfo();
			user1.setId(i);
			user1.setName("tommy" + i);
			String buffer = "";
			if (i < 10) {
				buffer = "000";
			}else if (i < 100) {
				buffer = "00";
			}else if (i < 1000) {
				buffer = "0";
			}
			user1.setCardNo("6224123456" + buffer + i);
			user1.setSalary(new BigDecimal("10000").add(new BigDecimal(i)));
			user1.setPersonIns(new BigDecimal("400").add(new BigDecimal(i)));
			user1.setFirmIns(new BigDecimal("600").add(new BigDecimal(i)));
			user1.setPersonInsTitle("person");
			user1.setFirmInsTitle("firm");
			list.add(user1);
		}

		ExcelWSheetProcessor<UserInfo> sheetProcessor = new ExcelWSheetProcessor<UserInfo>() {

			@Override
			public List<UserInfo> getDataList() {
				return list;
			}

			@Override
			public void beforeProcess() {
				Map<Integer, String> headerMap = new HashMap<Integer, String>();
				Row headRow = getSheet().createRow(getRowStartIndex() - 1);
				getSheet().setColumnWidth(2, 20*256);
				headRow.setHeightInPoints(20);
				headerMap.put(0, "id");
				headerMap.put(1, "name");
				headerMap.put(2, "cardNo");
				headerMap.put(3, "salary");
				headerMap.put(4, "type");
				headerMap.put(5, "insarance");
				for (int col = 0; col < headerMap.size(); col++) {
					String header = headerMap.get(col);
					Cell cell = headRow.createCell(col);
					cell.setCellValue(header);
				}
			}

			@Override
			public void afterProcess() {
			}

			@Override
			public void dealException(Exception e) {
				e.printStackTrace();
			}
		};
		sheetProcessor.setRowStartIndex(1);
		sheetProcessor.setPageSize(7);
		final String outputFilePath = "F:\\tmp\\salary.xls";
		File outputFile = new File(outputFilePath);
		outputFile.createNewFile();
		FileOutputStream output = new FileOutputStream(outputFile);

		ExcelWrite.write(ExcelType.XLS, output, sheetProcessor);
	}
}
