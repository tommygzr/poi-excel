package com.whl.poi;

import java.beans.PropertyDescriptor;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.BeanUtils;

import sun.reflect.generics.reflectiveObjects.ParameterizedTypeImpl;

public class ExcelRead {

	public static void read(InputStream inputStream, ExcelRSheetProcessor<?>... sheetProcessors) throws Exception {
		Workbook workbook = WorkbookFactory.create(inputStream);
		for (ExcelRSheetProcessor<?> sheetProcessor : sheetProcessors) {
			Integer sheetIndex = sheetProcessor.getSheetIndex();
			Sheet sheet = null;
			try {
				sheet = workbook.getSheetAt(sheetIndex);
				
				sheetProcessor.setSheet(sheet);
				sheetProcessor.beforeProcess();
				int pageSize = sheetProcessor.getPageSize();
				int readRowIndex = sheetProcessor.getRowStartIndex();
				Class<?> clazz = (Class<?>) ((ParameterizedTypeImpl) sheetProcessor.getClass().getGenericSuperclass())
						.getActualTypeArguments()[0];
				int total = sheet.getLastRowNum() - readRowIndex + 1;
				RowData rowData = clazz.getAnnotation(RowData.class);
				if( rowData != null ){
					total = (int) Math.ceil(total/rowData.rowAcross()); 
				}
				int size = total;
				List list = null;
				if (pageSize > 0) {
					int part = total / pageSize;//分批数
					size = pageSize;
					list = new ArrayList(size);
					for (int i = 0; i < part; i++) {
						readRowIndex = read(sheetProcessor, list, clazz, size, readRowIndex);
						sheetProcessor.process(list);
						list.clear();
					}
					size = total % pageSize;
				}
				if (size > 0) {
					list = new ArrayList(size);
					readRowIndex = read(sheetProcessor, list, clazz, size, readRowIndex);
					sheetProcessor.process(list);
					list.clear();
				}
			} catch (Exception e) {
				sheetProcessor.exceptionHandler(e);
			}
		}
	}

	private static int read(ExcelRSheetProcessor<?> sheetProcessor, List listPage, Class<?> clazz, int size,
			int readRowIndex) throws Exception{
		RowData rowData = clazz.getAnnotation(RowData.class);
		Sheet sheet = sheetProcessor.getSheet();
		for (int i = 0; i < size; i++) {
			List<Row> rowList = new ArrayList<Row>();
			for (int j = 0; rowData != null && j < rowData.rowAcross() || j < 1; j++) {
				Row row = sheet.getRow(readRowIndex + j);
				if (row == null) {
					continue;
				}
				rowList.add(row);
			}
			Field[] declaredFields = clazz.getDeclaredFields();
			Object obj = clazz.newInstance();
			if(i == 1){
				System.out.println();
			}
			for (Field field : declaredFields) {
				Column column = null ;
				try {
					column = field.getAnnotation(Column.class);
					Cell cell = rowList.get(column.rowIndex()).getCell(column.colIndex());
					Object fieldVal = getCell(cell, field.getType());
					PropertyDescriptor pd = BeanUtils.getPropertyDescriptor(clazz, field.getName());
					pd.getWriteMethod().invoke(obj, fieldVal);
				} catch (Exception e) {
					System.out.println(field.getName()+i);
					e.printStackTrace();
					if (column.dealType().equals(Column.dealType.CONTINUE_CELL)) {
						continue;
					}
					if (column.dealType().equals(Column.dealType.CONTINUE_ROW)) {
						break;
					}
					throw new RuntimeException("deal cell exception", e);
				}
			}
			listPage.add(obj);
			if (rowData != null) {
				readRowIndex = readRowIndex + rowData.rowAcross();
			} else {
				readRowIndex++;
			}
		}

		return readRowIndex;
	}

	private static Object getCell(Cell cell, Class<?> clazz) {
		if (clazz.equals(short.class) || clazz.equals(Short.class)) {
			return (short) cell.getNumericCellValue();
		} else if (clazz.equals(int.class) || clazz.equals(Integer.class)) {
			return (int) cell.getNumericCellValue();
		} else if (clazz.equals(long.class) || clazz.equals(Long.class)) {
			return (long) cell.getNumericCellValue();
		} else if (clazz.equals(float.class) || clazz.equals(Float.class)) {
			return (float) cell.getNumericCellValue();
		} else if (clazz.equals(double.class) || clazz.equals(Double.class)) {
			return cell.getNumericCellValue();
		} else if (clazz.equals(BigDecimal.class)) {
			return new BigDecimal(cell.getNumericCellValue());
		} else {
			return cell.getStringCellValue();
		}
	}

}
