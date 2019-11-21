import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;

public class TestPoiExport {
    public static void main(String[] args) throws Exception {
        //创建HSSFWorkbook对象(excel的文档对象)
        HSSFWorkbook wb = new HSSFWorkbook();
        //建立新的sheet对象（excel的表单）
        HSSFSheet sheet = wb.createSheet("成绩表");
        //在sheet里创建第一行，参数为行索引(excel的行)，可以是0～65535之间的任何一个
        HSSFRow row1 = sheet.createRow(0);
        //创建单元格（excel的单元格，参数为列索引，可以是0～255之间的任何一个
        HSSFCell cell = row1.createCell(0);
        //设置单元格内容
        cell.setCellValue("学员考试成绩一览表");
        //合并单元格CellRangeAddress构造参数依次表示起始行，截至行，起始列， 截至列
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
        //在sheet里创建第二行
        HSSFRow row2 = sheet.createRow(1);
        //创建单元格并设置单元格内容
        row2.createCell(0).setCellValue("姓名");
        row2.createCell(1).setCellValue("班级");
        row2.createCell(2).setCellValue("笔试成绩");
        row2.createCell(3).setCellValue("机试成绩");
        //在sheet里创建第三行
        HSSFRow row3 = sheet.createRow(2);
        row3.createCell(0).setCellValue("李明");
        row3.createCell(1).setCellValue("As178");
        row3.createCell(2).setCellValue(87);
        row3.createCell(3).setCellValue(78);
        //第四行
        HSSFRow row4 = sheet.createRow(3);
        row3.createCell(0).setCellValue("张飞");
        row3.createCell(1).setCellValue("As255");
        row3.createCell(2).setCellValue(78);
        row3.createCell(3).setCellValue(90);
        //第五行
        HSSFRow row5 = sheet.createRow(4);
        row3.createCell(0).setCellValue("王菲");
        row3.createCell(1).setCellValue("As336");
        row3.createCell(2).setCellValue(82);
        row3.createCell(3).setCellValue(69);

        //向指定文件写入内容
        FileOutputStream fos= new FileOutputStream("d:/学员考试成绩一览表.xlsx");
        wb.write(fos);
        fos.close();

    }
}
