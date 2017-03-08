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

<p>code for read</p>
<pre>
	<code>
	public static void read() throws Exception{
		ExcelRSheetProcessor<UserInfo> sheetProcessor = new ExcelRSheetProcessor<UserInfo>() {
			@Override
			public void beforeProcess() {
				// TODO Auto-generated method stub
			}
			@Override
			public void process(List<UserInfo> list) {
				for (UserInfo info : list) {
					System.out.println("id=" + info.getId() + " name=" + info.getName() + " cardNo=" + info.getCardNo()
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
		sheetProcessor.setPageSize(7);
		final String filePath = "F:\\tmp\\salary.xls";
		File file = new File(filePath);
		FileInputStream input = new FileInputStream(file);
		ExcelRead.read(input, sheetProcessor);
	}
	</code>
</pre>
