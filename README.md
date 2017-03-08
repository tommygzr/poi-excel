# JAVA poi-excel
<ol>
	<li>
		将poi封装成对象与excel的行对应关系(支持xls,xlsx)
	</li>
	<li>
		基于注解的poi解析，@RowData(有合并单元格时，使用)，@Column 对应的字段与单元格配置
	</li>
	<li>
		支持输出excel设置单元格格式
	</li>
</ol>

# Examples

<pre>
	<code>
		@RowData(rowAcross = 2, desc = "用户数据")
		public class UserInfo {
			//
			@Column(colIndex = 0, isAcross = true, format = "0")
			private int id;
			//
			@Column(colIndex = 1, isAcross = true)
			private String name;
			//
			@Column(colIndex = 2, isAcross = true)
			private String cardNo;
			//
			@Column(colIndex = 3, isAcross = true, format = "#,##0.00")
			private BigDecimal salary;
			//
			@Column(colIndex = 4, rowIndex = 0)
			private String personInsTitle;
			//
			@Column(colIndex = 4, rowIndex = 1)
			private String firmInsTitle;
			//
			@Column(colIndex = 5, rowIndex = 0, format = "#,##0.00")
			private BigDecimal personIns;
			//
			@Column(colIndex = 5, rowIndex = 1, format = "#,##0.00")
			private BigDecimal firmIns;
			//get,set方法省略......
		}
		//
		<p>code for read</p>
		//
		public static void read() throws Exception{
			ExcelRSheetProcessor<UserInfo> sheetProcessor = new ExcelRSheetProcessor<UserInfo>() {
				@Override
				public void beforeProcess() {
					// TODO Auto-generated method stub
				}
				@Override
				public void process(List<UserInfo> list) {
					for (UserInfo info : list) {
						System.out.println("id=" + info.getId() + " name=" + info.getName() 
							+ " cardNo=" + info.getCardNo()
							+ " salary=" + info.getSalary() + " personInsTitle=" + info.getPersonInsTitle()
							+ " firmInsTitle=" + info.getFirmInsTitle()
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
			sheetProcessor.setPageSize(2);
			final String filePath = "F:\\tmp\\salary.xls";
			File file = new File(filePath);
			FileInputStream input = new FileInputStream(file);
			ExcelRead.read(input, sheetProcessor);
		}
	</code>
</pre>
