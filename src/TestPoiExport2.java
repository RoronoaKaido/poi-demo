import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;

public class TestPoiExport2 {
        @SuppressWarnings("resource")
        public static void main(String[] args) throws Exception {
            //创建工作簿---->XSSF代表10版的Excel(HSSF是03版的Excel)
            XSSFWorkbook wb = new XSSFWorkbook();
            //工作表

            XSSFCellStyle red = wb.createCellStyle();
            XSSFCellStyle re1 = wb.createCellStyle();
            XSSFFont font = wb.createFont();
            font.setColor(Font.COLOR_RED);
            red.setFont(font);
            XSSFSheet sheet = wb.createSheet("会员表");
            //标头行，代表第一行
            XSSFRow header=sheet.createRow(0);
            //创建单元格，0代表第一行第一列
            XSSFCell cell0=header.createCell(0);
            //设置单元格内容
            cell0.setCellValue("会员信息表");
            //合并单元格CellRangeAddress构造参数依次表示起始行，截至行，起始列， 截至列
            sheet.addMergedRegion(new CellRangeAddress(0,0,0,6));
            //在sheet里创建第二行
            XSSFRow header2 = sheet.createRow(1);
            header2.createCell(0).setCellValue("会员级别");
            header2.createCell(1).setCellValue("会员编号");
            header2.getCell(1).setCellStyle(red);
            header2.createCell(2).setCellValue("会员姓名");
            header2.createCell(3).setCellValue("推荐人编号");
            header2.createCell(4).setCellValue("负责人编号");
            header2.createCell(5).setCellValue("地址编号");
            header2.createCell(6).setCellValue("注册时间");
            //设置列的宽度
            //getPhysicalNumberOfCells()代表这行有多少包含数据的列
            for(int i=0;i<header.getPhysicalNumberOfCells();i++){
                //POI设置列宽度时比较特殊，它的基本单位是1/255个字符大小，
                //因此我们要想让列能够盛的下20个字符的话，就需要用255*20
                sheet.setColumnWidth(i, 255*20);
            }
            //设置行高，行高的单位就是像素，因此30就是30像素的意思
            header.setHeightInPoints(30);
            deleteColumn(sheet,0);

            //上面设置好了内容，我们当然是要输出到某个文件的，输出就需要有输出流
            FileOutputStream fos= new FileOutputStream("d:/2010.xlsx");
            //向指定文件写入内容
            wb.write(fos);
            fos.close();
        }

    /**
     * 删除列
     * @param sheet
     * @param columnToDelete
     */
    public static void deleteColumn(XSSFSheet sheet, int columnToDelete) {
        for (int rId = 0; rId <= sheet.getLastRowNum(); rId++) {
            Row row = sheet.getRow(rId);
            for (int cID = columnToDelete; cID <= row.getLastCellNum(); cID++) {
                Cell cOld = row.getCell(cID);
                if (cOld != null) {
                    row.removeCell(cOld);
                }
                Cell cNext = row.getCell(cID + 1);
                if (cNext != null) {
                    Cell cNew = row.createCell(cID, cNext.getCellTypeEnum());
                    cloneCell(cNew, cNext);
                    //Set the column width only on the first row.
                    //Other wise the second row will overwrite the original column width set previously.
                    if (rId == 0) {
                        sheet.setColumnWidth(cID, sheet.getColumnWidth(cID + 1));

                    }
                }
            }
        }
    }

    /**
     * 右边列左移
     * @param cNew
     * @param cOld
     */
    public static void cloneCell(Cell cNew, Cell cOld) {
        cNew.setCellComment(cOld.getCellComment());
        cNew.setCellStyle(cOld.getCellStyle());

        if (CellType.BOOLEAN == cNew.getCellTypeEnum()) {
            cNew.setCellValue(cOld.getBooleanCellValue());
//            cNew.getCellType()
        } else if (CellType.NUMERIC == cNew.getCellTypeEnum()) {
            cNew.setCellValue(cOld.getNumericCellValue());
        } else if (CellType.STRING == cNew.getCellTypeEnum()) {
            cNew.setCellValue(cOld.getStringCellValue());
        } else if (CellType.ERROR == cNew.getCellTypeEnum()) {
            cNew.setCellValue(cOld.getErrorCellValue());
        } else if (CellType.FORMULA == cNew.getCellTypeEnum()) {
            cNew.setCellValue(cOld.getCellFormula());
        }
    }
}

