package demo2;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.xmlbeans.XmlException;

/**
 * @author xww
 * 这是一个工具类定义了操作excel表的相关方法
 * */
public class DoWork {

	public static void main(String[] args) throws IOException, XmlException,
			OpenXML4JException {

		List<String> list = WordReadUtil.flagWordVersion(
				WordReadUtil.getPath("2.doc"), "3", "公共组件");
		System.out.println(list);
		String[] excelHeader = { "工作任务", "作者", };
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet("Campaign");
		// sheet.setColumnWidth(0,100*2*256);
		sheet.autoSizeColumn(1, true);
		HSSFRow row = sheet.createRow((int) 0);
		HSSFCellStyle style = wb.createCellStyle();
		HSSFFont font = wb.createFont();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		// 设置表头
		for (int i = 0; i < excelHeader.length; i++) {
			HSSFCell cell = row.createCell(i);
			cell.setCellValue(excelHeader[i]);
			cell.setCellStyle(style);
		}
		// 填充数据
		for (int i = 0; i < list.size(); i++) {
			sheet.setColumnWidth(i, list.get(i).length() * 2 * 256);//宽度宽一点
			row = sheet.createRow(i + 1);
			row.createCell(0).setCellValue(list.get(i));
			row.createCell(1).setCellValue("");
		}
		OutputStream ouputStream = new FileOutputStream(new File(
				"E:\\testExcel\\task.xls"));
		wb.write(ouputStream);
	}

}
