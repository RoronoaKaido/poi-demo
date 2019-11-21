import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class PoiExportTest {

    //导出树形结构数据的Excel表格
    public static void main(String[] args)  {

        List<Tag> tagList = getTagList();
        System.out.println(tagList.toString());

        //创建工作簿---->XSSF代表10版的Excel(HSSF是03版的Excel)
        XSSFWorkbook wb = new XSSFWorkbook();
        //工作表
        XSSFSheet sheet = wb.createSheet("知识点编码表-树形结构");
        sheet.setDefaultColumnWidth(255*30);

        //设置标题栏的样式
        XSSFCellStyle titleStyle = getTitleStyle(wb);
        //设置非标题栏的样式
        XSSFCellStyle style = getStyle(wb);

        //标头行，代表第一行
        createTitleRow(sheet, titleStyle);


        //在sheet里创建内容行
        createContentRow(sheet, style, tagList);

        //输出流
        try {
            FileOutputStream fos= new FileOutputStream("d:/树形结构数据表.xlsx");
            //向指定文件写入内容
            wb.write(fos);
            wb.close();
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<Tag> getTagList() {
        List<Tag> tagList = new ArrayList<>();
        Tag tag1 = new Tag(100L,"算法工程师");
        Tag tag2 = new Tag(200L,"编程语言");
        Tag tag3 = new Tag(201L,"组合数学");
        Tag tag4 = new Tag(202L,"数据分析");
        Tag tag5 = new Tag(300L,"基础");
        Tag tag6 = new Tag(301L,"Python");
        Tag tag7 = new Tag(302L,"SPSS");
        Tag tag8 = new Tag(303L,"智力题");
        Tag tag9 = new Tag(304L,"数理统计");
        Tag tag10 = new Tag(305L,"业务知识");
        Tag tag11 = new Tag(306L,"解决方案");
        Tag tag12 = new Tag(400L,"图片");
        Tag tag13 = new Tag(401L,"查找");
        Tag tag14 = new Tag(402L,"学习Python");
        Tag tag15 = new Tag(403L,"智力题1");
        Tag tag16 = new Tag(404L,"数理统计1");
        Tag tag17 = new Tag(405L,"业务知识1");
        Tag tag18 = new Tag(406L,"解决方案");

        //四级集合
        List<Tag> fourth1 = new ArrayList<>();
        List<Tag> fourth2 = new ArrayList<>();
        List<Tag> fourth3 = new ArrayList<>();
        List<Tag> fourth4 = new ArrayList<>();
        List<Tag> fourth5 = new ArrayList<>();
        List<Tag> fourth6 = new ArrayList<>();

        List<Tag> third1 = new ArrayList<>();
        List<Tag> third2= new ArrayList<>();
        List<Tag> third3 = new ArrayList<>();

        List<Tag> second1 = new ArrayList<>();

        //基础子级
        fourth1.add(tag12);
        fourth1.add(tag13);
        tag5.setChilrenList(fourth1);
        //PYTHON子级
        fourth2.add(tag14);
        tag6.setChilrenList(fourth2);
        //智力题子级
        fourth3.add(tag15);
        tag8.setChilrenList(fourth3);
        //数理统计子级
        fourth4.add(tag16);
        tag9.setChilrenList(fourth4);
        //业务知识子级
        fourth5.add(tag17);
        tag10.setChilrenList(fourth5);
        //解决方案子级
        fourth6.add(tag18);
        tag11.setChilrenList(fourth6);

        //编程语言子级
        third1.add(tag5);
        third1.add(tag6);
        third1.add(tag7);
        tag2.setChilrenList(third1);
        //组合数学子级
        third2.add(tag8);
        third2.add(tag9);
        tag3.setChilrenList(third2);
        //数据分析子级
        third3.add(tag10);
        third3.add(tag11);
        tag4.setChilrenList(third3);

        second1.add(tag2);
        second1.add(tag3);
        second1.add(tag4);
        tag1.setChilrenList(second1);

        tagList.add(tag1);
        return tagList;
    }

    private static void createContentRow(XSSFSheet sheet, XSSFCellStyle style, List<Tag> tagList) {
        int index = 1;
        int firstLevelStartMergeAdress = 1;
        int secondLevelStartMergeAdress = 1;
        int thirdLevelStartMergeAdress = 1;
        int fourthLevelStartMergeAdress = 1;
        for (Tag firstTag : tagList) {
            List<Tag> secondList = firstTag.getChilrenList();
            if (CollectionUtils.isEmpty(secondList)) {
                //输入一级
                index = createRow(sheet, style, index, firstTag, null, null, null);
                secondLevelStartMergeAdress++;
                thirdLevelStartMergeAdress++;
                fourthLevelStartMergeAdress++;
            } else {
                for (Tag secondTag : secondList) {
                    List<Tag> thirdList = secondTag.getChilrenList();
                    if (CollectionUtils.isEmpty(thirdList)) {
                        //输入一级、二级
                        index = createRow(sheet, style, index, firstTag, secondTag, null, null);
                        thirdLevelStartMergeAdress++;
                        fourthLevelStartMergeAdress++;
                    } else {
                        for (Tag thirdTag : thirdList) {
                            List<Tag> fourthList = thirdTag.getChilrenList();
                            if (CollectionUtils.isEmpty(fourthList)) {
                                //输入一级、二级、三级
                                index = createRow(sheet, style, index, firstTag, secondTag, thirdTag, null);
                                fourthLevelStartMergeAdress++;
                            } else {
                                for (Tag fourthTag : fourthList) {
                                    //输入一级、二级、三级、四级
                                    index = createRow(sheet, style, index, firstTag, secondTag, thirdTag, fourthTag);
                                    fourthLevelStartMergeAdress++;
                                }
                                if (thirdLevelStartMergeAdress < fourthLevelStartMergeAdress - 1) {
                                    //合并三级单元格
                                    sheet.addMergedRegion(new CellRangeAddress(thirdLevelStartMergeAdress,
                                            fourthLevelStartMergeAdress - 1
                                            , 4, 4));
                                    sheet.addMergedRegion(new CellRangeAddress(thirdLevelStartMergeAdress,
                                            fourthLevelStartMergeAdress - 1
                                            , 5, 5));
                                }
                            }
                            thirdLevelStartMergeAdress = fourthLevelStartMergeAdress;
                        }
                        if (secondLevelStartMergeAdress < thirdLevelStartMergeAdress - 1) {
                            //合并二级单元格
                            sheet.addMergedRegion(new CellRangeAddress(secondLevelStartMergeAdress,
                                    thirdLevelStartMergeAdress - 1
                                    , 2, 2));
                            sheet.addMergedRegion(new CellRangeAddress(secondLevelStartMergeAdress,
                                    thirdLevelStartMergeAdress - 1
                                    , 3, 3));
                        }
                    }
                    secondLevelStartMergeAdress = thirdLevelStartMergeAdress;
                }
                if (firstLevelStartMergeAdress < secondLevelStartMergeAdress - 1) {
                    //合并一级单元格
                    sheet.addMergedRegion(new CellRangeAddress(firstLevelStartMergeAdress,
                            secondLevelStartMergeAdress - 1
                            , 0, 0));
                    sheet.addMergedRegion(new CellRangeAddress(firstLevelStartMergeAdress,
                            secondLevelStartMergeAdress - 1
                            , 1, 1));
                }
            }
            firstLevelStartMergeAdress = secondLevelStartMergeAdress;
        }
    }

    private static void createTitleRow(XSSFSheet sheet, XSSFCellStyle titleStyle) {
        XSSFRow header=sheet.createRow(0);
        int index = 0;
        Cell cell0;
        cell0 = header.createCell(index++);
        cell0.setCellStyle(titleStyle);
        cell0.setCellValue("一级知识点名称");
        cell0 = header.createCell(index++);
        cell0.setCellStyle(titleStyle);
        cell0.setCellValue("一级知识点编码");
        cell0 = header.createCell(index++);
        cell0.setCellStyle(titleStyle);
        cell0.setCellValue("二级知识点名称");
        cell0 = header.createCell(index++);
        cell0.setCellStyle(titleStyle);
        cell0.setCellValue("二级知识点编码");
        cell0 = header.createCell(index++);
        cell0.setCellStyle(titleStyle);
        cell0.setCellValue("三级知识点名称");
        cell0 = header.createCell(index++);
        cell0.setCellStyle(titleStyle);
        cell0.setCellValue("三级知识点编码");
        cell0 = header.createCell(index++);
        cell0.setCellStyle(titleStyle);
        cell0.setCellValue("四级知识点名称");
        cell0 = header.createCell(index++);
        cell0.setCellStyle(titleStyle);
        cell0.setCellValue("四级知识点编码");
    }

    private static XSSFCellStyle getStyle(XSSFWorkbook wb) {
        XSSFCellStyle style = wb.createCellStyle();
        style.setWrapText(true);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置非标题栏的字体
        XSSFFont font = wb.createFont();
        font.setColor(HSSFColor.BLACK.index);
        font.setFontHeightInPoints((short) 11);
        style.setFont(font);
        return style;
    }

    private static XSSFCellStyle getTitleStyle(XSSFWorkbook wb) {
        XSSFCellStyle titleStyle = wb.createCellStyle();
        titleStyle.setFillBackgroundColor(HSSFColor.BLUE_GREY.index);
        titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        //设置标题栏的字体
        XSSFFont titleFont = wb.createFont();
        titleFont.setColor(HSSFColor.WHITE.index);
        titleFont.setFontHeightInPoints((short) 11);
        titleFont.setBold(true);
        titleStyle.setFont(titleFont);
        return titleStyle;
    }

    private static int createRow(XSSFSheet sheet, XSSFCellStyle style,
                                 int index, Tag firstTag, Tag secondTag, Tag thirdTag, Tag fourthTag) {
        XSSFRow row = sheet.createRow(index);
//        Cell cell;
//        for (int i = 0; i < 8; i++) {
//            cell = row.createCell(i);
//            cell.setCellStyle(style);
//        }
        Cell cell0 = row.createCell(0);
        cell0.setCellStyle(style);
        Cell cell1 = row.createCell(1);
        cell1.setCellStyle(style);
        Cell cell2 = row.createCell(2);
        cell2.setCellStyle(style);
        Cell cell3 = row.createCell(3);
        cell3.setCellStyle(style);
        Cell cell4 = row.createCell(4);
        cell4.setCellStyle(style);
        Cell cell5 = row.createCell(5);
        cell5.setCellStyle(style);
        Cell cell6 = row.createCell(6);
        cell6.setCellStyle(style);
        Cell cell7 = row.createCell(7);
        cell7.setCellStyle(style);

        if (firstTag != null){
            /*row.getCell(0)*/cell0.setCellValue(firstTag.getName());
            /*row.getCell(1)*/cell1.setCellValue(firstTag.getId());
        }
         if (secondTag != null){
            /*row.getCell(2)*/cell2.setCellValue(secondTag.getName());
            /*row.getCell(3)*/cell3.setCellValue(secondTag.getId());
        }
         if (thirdTag != null){
            /*row.getCell(4)*/cell4.setCellValue(thirdTag.getName());
            /*row.getCell(5)*/cell5.setCellValue(thirdTag.getId());
        }
         if (fourthTag != null) {
             /*row.getCell(6)*/cell6.setCellValue(fourthTag.getName());
             /*row.getCell(7)*/cell7.setCellValue(fourthTag.getId());
         }
         index++;
         return index;
    }
}
