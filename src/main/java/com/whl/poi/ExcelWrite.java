package com.whl.poi;

import java.beans.PropertyDescriptor;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.BeanUtils;
import org.springframework.util.StringUtils;

import sun.reflect.generics.reflectiveObjects.ParameterizedTypeImpl;

public class ExcelWrite {

	public static void write(ExcelType fileType, OutputStream outputStream, ExcelWSheetProcessor<?>... sheetProcessors)
			throws Exception {
		Workbook workbook = null;
		if (fileType == ExcelType.XLS) {
			workbook = new HSSFWorkbook();
		} else {
			workbook = new XSSFWorkbook();
		}
		write(workbook, outputStream, sheetProcessors);
	}

	private static void write(Workbook workbook, OutputStream outputStream, ExcelWSheetProcessor<?>... sheetProcessors)
			throws Exception {
		for (ExcelWSheetProcessor<?> sheetProcessor : sheetProcessors) {
			try {
				String sheetName = sheetProcessor.getSheetName();
				Integer sheetIndex = sheetProcessor.getSheetIndex();
				Sheet sheet = null;
				if (sheetName != null) {
					sheet = workbook.getSheet(sheetName);
					if (sheet != null) {
						if (sheetIndex != null && !sheetIndex.equals(workbook.getSheetIndex(sheet))) {
							throw new IllegalArgumentException(
									"sheetName[" + sheetName + "] and sheetIndex[" + sheetIndex + "] not match.");
						}
					} else {
						sheet = workbook.createSheet(sheetName);
						if (sheetIndex != null) {
							workbook.setSheetOrder(sheetName, sheetIndex);
						}
					}
				} else if (sheetIndex != null) {
					sheet = workbook.createSheet();
					workbook.setSheetOrder(sheet.getSheetName(), sheetIndex);
				} else {
					sheet = workbook.createSheet();
					workbook.setSheetOrder(sheet.getSheetName(), 0);
				}

				if (sheetIndex == null) {
					sheetIndex = workbook.getSheetIndex(sheet);
				}
				if (sheetName == null) {
					sheetName = sheet.getSheetName();
				}
				// proc sheet
				sheetProcessor.setSheet(sheet);

				sheetProcessor.beforeProcess();

				int writeRowIndex = sheetProcessor.getRowStartIndex();
				List<?> dataList = sheetProcessor.getDataList();
				Class<?> clazz = (Class<?>) ((ParameterizedTypeImpl) sheetProcessor.getClass().getGenericSuperclass())
						.getActualTypeArguments()[0];
				Map<String, CellStyle> cellStyles = getCellStyle(workbook, clazz);
				int pageSize = sheetProcessor.getPageSize();
				if (dataList != null && dataList.size() > 0) {
					if (pageSize > 0) {
						int size = dataList.size();
						if (pageSize < size) {
							int part = size / pageSize;//分批数
							for (int i = 0; i < part; i++) {
								List<?> listPage = dataList.subList(0, pageSize);
								writeRowIndex = write(sheetProcessor, listPage, clazz, cellStyles, writeRowIndex);
								dataList.subList(0, pageSize).clear();
							}
						}
					}
					if (!dataList.isEmpty()) {
						writeRowIndex = write(sheetProcessor, dataList, clazz, cellStyles, writeRowIndex);
					}
					dataList.clear();
				}
				try {
					workbook.write(outputStream);
				} catch (IOException e) {
					throw new RuntimeException(e);
				}
			} catch (Exception e) {
				sheetProcessor.dealException(e);
			}
		}
	}

	private static int write(ExcelWSheetProcessor<?> sheetProcessor, List<?> listPage, Class<?> clazz,
			Map<String, CellStyle> cellStyles, int writeRowIndex) throws Exception {
		Sheet sheet = sheetProcessor.getSheet();
		RowData rowData = clazz.getAnnotation(RowData.class);
		Field[] declaredFields = clazz.getDeclaredFields();
		for (Object obj : listPage) {
			List<Row> rowList = new ArrayList<Row>();
			for (int i = 0; rowData != null && i < rowData.rowAcross() || i < 1; i++) {
				Row row = sheet.getRow(writeRowIndex + i);
				if (row == null) {
					row = sheet.createRow(writeRowIndex + i);
				}
				rowList.add(row);
			}
			for (Field field : declaredFields) {
				PropertyDescriptor pd = BeanUtils.getPropertyDescriptor(clazz, field.getName());
				Object val = pd.getReadMethod().invoke(obj);
				Column column = field.getAnnotation(Column.class);
				Cell cell = rowList.get(column.rowIndex()).createCell(column.colIndex());
				if (rowData != null && column.isAcross()) {
					CellRangeAddress cra = new CellRangeAddress(writeRowIndex, writeRowIndex + rowData.rowAcross() - 1,
							column.colIndex(), column.colIndex());
					sheet.addMergedRegion(cra);
				}
				try {
					setCell(cell, field.getType(), val);
					CellStyle cellStyle = cellStyles.get(field.getName());
					if (cellStyle != null) {
						cell.setCellStyle(cellStyle);
					}
				} catch (Exception e) {
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
			if (rowData != null) {
				writeRowIndex = writeRowIndex + rowData.rowAcross();
			} else {
				writeRowIndex++;
			}
		}
		return writeRowIndex;
	}

	private static void setCell(Cell cell, Class<?> clazz, Object val) {
		if (clazz.equals(short.class) || clazz.equals(Short.class)) {
			cell.setCellValue((Short) val);
		} else if (clazz.equals(int.class) || clazz.equals(Integer.class)) {
			cell.setCellValue((Integer) val);
		} else if (clazz.equals(long.class) || clazz.equals(Long.class)) {
			cell.setCellValue((Long) val);
		} else if (clazz.equals(float.class) || clazz.equals(Float.class)) {
			cell.setCellValue((Float) val);
		} else if (clazz.equals(double.class) || clazz.equals(Double.class)) {
			cell.setCellValue((Double) val);
		} else if (clazz.equals(BigDecimal.class)) {
			cell.setCellValue(((BigDecimal) val).doubleValue());
			
		} else {
			cell.setCellValue(val.toString());
		}
	}

	private static Map<String, CellStyle> getCellStyle(Workbook workbook, Class<?> clazz) {
		Map<String, CellStyle> cellStyles = new HashMap<String, CellStyle>();
		Field[] declaredFields = clazz.getDeclaredFields();
		for (Field field : declaredFields) {
			Column column = field.getAnnotation(Column.class);
			if (!StringUtils.isEmpty(column.format())) {
				CellStyle cellStyle =workbook.createCellStyle();
				cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(column.format()));
				cellStyles.put(field.getName(), cellStyle);
			}
		}
		return cellStyles;
	}

}
