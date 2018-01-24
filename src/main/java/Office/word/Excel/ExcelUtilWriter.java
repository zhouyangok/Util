package Office.word.Excel;

/**
 * Created with IntelliJ IDEA.
 *
 * @Author medo_zy
 * @Desciption:生成固定格式的excel报表
 * @Date 2017-12-18 14:07
 */

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.List;
public class ExcelUtilWriter {
//默认单元格内容为数字时格式
private static DecimalFormat df = new DecimalFormat("0");
// 默认单元格格式化日期字符串
private static SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
// 格式化数字
private static DecimalFormat nf = new DecimalFormat("0.00");


   /* *@return 将返回结果存储在ArrayList内，存储结构与二位数组类似
    *lists.get(0).

    get(0)表示过去Excel中0行0列单元格*/


public static String writeExcel(List<FrontSpaceSeries> result, String path) {
        if(result == null){
        return null;
        }
        String seriesName = result.get(0).getSeriesName();
        String typeName = result.get(0).getTypeName();
        String xx = result.get(0).getXx();
        String xy = result.get(0).getXy();
        String xz = result.get(0).getXz();
        String yx = result.get(0).getYx();
        String yy = result.get(0).getYy();
        String yz = result.get(0).getYz();
        String zx = result.get(0).getZx();
        String zy = result.get(0).getZy();
        String zz = result.get(0).getZz();
        //新建excel报表
        HSSFWorkbook wb = new HSSFWorkbook();
        //添加一个sheet，名字叫做excel报表
        HSSFSheet sheet = wb.createSheet("excel报表");
        sheet.setColumnWidth(0, 3366);
        sheet.setColumnWidth(1, 3366);
        sheet.setColumnWidth(2, 3366);
        //新建单元格样式
        HSSFCellStyle cellStyle = wb.createCellStyle();
        HSSFCellStyle cellStyle2 = wb.createCellStyle();
        //设置边框
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框

        cellStyle2.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
        cellStyle2.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
        cellStyle2.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
        cellStyle2.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
        //设置字体
        HSSFFont font2 = wb.createFont();
        font2.setFontName("仿宋_GB2312");
        font2.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//粗体显示
        font2.setFontHeightInPoints((short) 12);  //字体大小

        cellStyle2.setFont(font2);//选择需要用到的字体格式

        // 居中
        cellStyle2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //设置自适应列表宽度
        sheet.autoSizeColumn(1, true);
        //第一行
        HSSFRow row0 = sheet.createRow(0);
        row0.setHeight((short)500);
        HSSFCell cell0_0 = row0.createCell(0);
        HSSFCell cell0_1 = row0.createCell(1);
        HSSFCell cell0_2 = row0.createCell(2);
        HSSFCell cell0_3 = row0.createCell(3);
        HSSFCell cell0_4 = row0.createCell(4);
        HSSFCell cell0_5 = row0.createCell(5);
        HSSFCell cell0_6 = row0.createCell(6);
        HSSFCell cell0_7 = row0.createCell(7);
        HSSFCell cell0_8 = row0.createCell(8);
        HSSFCell cell0_9 = row0.createCell(9);
        HSSFCell cell0_10 = row0.createCell(10);
        HSSFCell cell0_11 = row0.createCell(11);
        HSSFCell cell0_12 = row0.createCell(12);
        HSSFCell cell0_13 = row0.createCell(13);
        HSSFCell cell0_14 = row0.createCell(14);
        cell0_0.setCellStyle(cellStyle2);
        cell0_1.setCellStyle(cellStyle);
        cell0_2.setCellStyle(cellStyle);
        cell0_3.setCellStyle(cellStyle);
        cell0_4.setCellStyle(cellStyle);
        cell0_5.setCellStyle(cellStyle);
        cell0_6.setCellStyle(cellStyle);
        cell0_7.setCellStyle(cellStyle);
        cell0_8.setCellStyle(cellStyle);
        cell0_9.setCellStyle(cellStyle);
        cell0_10.setCellStyle(cellStyle);
        cell0_11.setCellStyle(cellStyle);
        cell0_12.setCellStyle(cellStyle);
        cell0_13.setCellStyle(cellStyle);
        cell0_14.setCellStyle(cellStyle);
        cell0_0.setCellValue("测量数据对比分析报告");
        //第二行
        HSSFRow row1 = sheet.createRow(1);
        row1.setHeight((short) 400);
        HSSFCell cell1_0 = row1.createCell(0);
        HSSFCell cell1_1 = row1.createCell(1);
        HSSFCell cell1_2 = row1.createCell(2);
        HSSFCell cell1_3 = row1.createCell(3);
        HSSFCell cell1_4 = row1.createCell(4);
        HSSFCell cell1_5 = row1.createCell(5);
        HSSFCell cell1_6 = row1.createCell(6);
        HSSFCell cell1_7 = row1.createCell(7);
        HSSFCell cell1_8 = row1.createCell(8);
        HSSFCell cell1_9 = row1.createCell(9);
        HSSFCell cell1_10 = row1.createCell(10);
        HSSFCell cell1_11 = row1.createCell(11);
        HSSFCell cell1_12 = row1.createCell(12);
        HSSFCell cell1_13 = row1.createCell(13);
        HSSFCell cell1_14 = row1.createCell(14);
        cell1_0.setCellStyle(cellStyle);
        cell1_1.setCellStyle(cellStyle);
        cell1_2.setCellStyle(cellStyle);
        cell1_3.setCellStyle(cellStyle);
        cell1_4.setCellStyle(cellStyle);
        cell1_5.setCellStyle(cellStyle);
        cell1_6.setCellStyle(cellStyle);
        cell1_7.setCellStyle(cellStyle);
        cell1_8.setCellStyle(cellStyle);
        cell1_9.setCellStyle(cellStyle);
        cell1_10.setCellStyle(cellStyle);
        cell1_11.setCellStyle(cellStyle);
        cell1_12.setCellStyle(cellStyle);
        cell1_13.setCellStyle(cellStyle);
        cell1_14.setCellStyle(cellStyle);
        cell1_9.setCellValue("测量值单位");
        cell1_11.setCellValue("变化量单位");
        //第三行
        HSSFRow row2 = sheet.createRow(2);
        row2.setHeight((short) 450);
        HSSFCell cell2_0 = row2.createCell(0);
        HSSFCell cell2_1 = row2.createCell(1);
        HSSFCell cell2_2 = row2.createCell(2);
        HSSFCell cell2_3 = row2.createCell(3);
        HSSFCell cell2_4 = row2.createCell(4);
        HSSFCell cell2_5 = row2.createCell(5);
        HSSFCell cell2_6 = row2.createCell(6);
        HSSFCell cell2_7 = row2.createCell(7);
        HSSFCell cell2_8 = row2.createCell(8);
        HSSFCell cell2_9 = row2.createCell(9);
        HSSFCell cell2_10 = row2.createCell(10);
        HSSFCell cell2_11 = row2.createCell(11);
        HSSFCell cell2_12 = row2.createCell(12);
        HSSFCell cell2_13 = row2.createCell(13);
        HSSFCell cell2_14 = row2.createCell(14);
        cell2_0.setCellStyle(cellStyle);
        cell2_1.setCellStyle(cellStyle);
        cell2_2.setCellStyle(cellStyle);
        cell2_3.setCellStyle(cellStyle);
        cell2_4.setCellStyle(cellStyle);
        cell2_5.setCellStyle(cellStyle);
        cell2_6.setCellStyle(cellStyle);
        cell2_7.setCellStyle(cellStyle);
        cell2_8.setCellStyle(cellStyle);
        cell2_9.setCellStyle(cellStyle);
        cell2_10.setCellStyle(cellStyle);
        cell2_11.setCellStyle(cellStyle);
        cell2_12.setCellStyle(cellStyle);
        cell2_13.setCellStyle(cellStyle);
        cell2_14.setCellStyle(cellStyle);
        cell2_2.setCellValue("型号名称");
        cell2_3.setCellValue(typeName);
        cell2_3.setCellStyle(cellStyle);
        //第四行
        HSSFRow row3 = sheet.createRow(3);
        row3.setHeight((short)450);
        HSSFCell cell3_0 = row3.createCell(0);
        HSSFCell cell3_1 = row3.createCell(1);
        HSSFCell cell3_2 = row3.createCell(2);
        HSSFCell cell3_3 = row3.createCell(3);
        HSSFCell cell3_4 = row3.createCell(4);
        HSSFCell cell3_5 = row3.createCell(5);
        HSSFCell cell3_6 = row3.createCell(6);
        HSSFCell cell3_7 = row3.createCell(7);
        HSSFCell cell3_8 = row3.createCell(8);
        HSSFCell cell3_9 = row3.createCell(9);
        HSSFCell cell3_10 = row3.createCell(10);
        HSSFCell cell3_11 = row3.createCell(11);
        HSSFCell cell3_12 = row3.createCell(12);
        HSSFCell cell3_13 = row3.createCell(13);
        HSSFCell cell3_14 = row3.createCell(14);
        cell3_2.setCellValue("测量项目");
        cell3_3.setCellValue(seriesName);
        cell3_3.setCellStyle(cellStyle);
        cell3_0.setCellStyle(cellStyle);
        cell3_1.setCellStyle(cellStyle);
        cell3_2.setCellStyle(cellStyle);
        cell3_4.setCellStyle(cellStyle);
        cell3_5.setCellStyle(cellStyle);
        cell3_6.setCellStyle(cellStyle);
        cell3_7.setCellStyle(cellStyle);
        cell3_8.setCellStyle(cellStyle);
        cell3_9.setCellStyle(cellStyle);
        cell3_10.setCellStyle(cellStyle);
        cell3_11.setCellStyle(cellStyle);
        cell3_12.setCellStyle(cellStyle);
        cell3_13.setCellStyle(cellStyle);
        cell3_14.setCellStyle(cellStyle);
        //第五行
        HSSFRow row4 = sheet.createRow(4);
        row4.setHeight((short)450);
        HSSFCell cell4_0 = row4.createCell(0);
        HSSFCell cell4_1 = row4.createCell(1);
        HSSFCell cell4_2 = row4.createCell(2);
        HSSFCell cell4_3 = row4.createCell(3);
        HSSFCell cell4_4 = row4.createCell(4);
        HSSFCell cell4_5 = row4.createCell(5);
        HSSFCell cell4_6 = row4.createCell(6);
        HSSFCell cell4_7 = row4.createCell(7);
        HSSFCell cell4_8 = row4.createCell(8);
        HSSFCell cell4_9 = row4.createCell(9);
        HSSFCell cell4_10 = row4.createCell(10);
        HSSFCell cell4_11 = row4.createCell(11);
        HSSFCell cell4_12 = row4.createCell(12);
        HSSFCell cell4_13 = row4.createCell(13);
        HSSFCell cell4_14 = row4.createCell(14);
        cell4_0.setCellValue("数据结果");
        cell4_0.setCellStyle(cellStyle2);
        cell4_1.setCellStyle(cellStyle);
        cell4_2.setCellStyle(cellStyle);
        cell4_3.setCellStyle(cellStyle);
        cell4_4.setCellStyle(cellStyle);
        cell4_5.setCellStyle(cellStyle);
        cell4_6.setCellStyle(cellStyle);
        cell4_7.setCellStyle(cellStyle);
        cell4_8.setCellStyle(cellStyle);
        cell4_9.setCellStyle(cellStyle);
        cell4_10.setCellStyle(cellStyle);
        cell4_11.setCellStyle(cellStyle);
        cell4_12.setCellStyle(cellStyle);
        cell4_13.setCellStyle(cellStyle);
        cell4_14.setCellStyle(cellStyle);
        //第六行
        HSSFRow row5 = sheet.createRow(5);
        row5.setHeight((short)450);
        HSSFCell cell5_0 = row5.createCell(0);
        HSSFCell cell5_1 = row5.createCell(1);
        HSSFCell cell5_2 = row5.createCell(2);
        HSSFCell cell5_3 = row5.createCell(3);
        HSSFCell cell5_4 = row5.createCell(4);
        HSSFCell cell5_5 = row5.createCell(5);
        HSSFCell cell5_6 = row5.createCell(6);
        HSSFCell cell5_7 = row5.createCell(7);
        HSSFCell cell5_8 = row5.createCell(8);
        HSSFCell cell5_9 = row5.createCell(9);
        HSSFCell cell5_10 = row5.createCell(10);
        HSSFCell cell5_11 = row5.createCell(11);
        HSSFCell cell5_12 = row5.createCell(12);
        HSSFCell cell5_13 = row5.createCell(13);
        HSSFCell cell5_14 = row5.createCell(14);
        cell5_0.setCellStyle(cellStyle);
        cell5_1.setCellStyle(cellStyle);
        cell5_2.setCellStyle(cellStyle);
        cell5_3.setCellStyle(cellStyle);
        cell5_4.setCellStyle(cellStyle);
        cell5_5.setCellStyle(cellStyle);
        cell5_6.setCellStyle(cellStyle);
        cell5_7.setCellStyle(cellStyle);
        cell5_8.setCellStyle(cellStyle);
        cell5_9.setCellStyle(cellStyle);
        cell5_10.setCellStyle(cellStyle);
        cell5_11.setCellStyle(cellStyle);
        cell5_12.setCellStyle(cellStyle);
        cell5_13.setCellStyle(cellStyle);
        cell5_14.setCellStyle(cellStyle);
        cell5_3.setCellValue("基准镜");
        //第七行
        HSSFRow row6 = sheet.createRow(6);
        row6.setHeight((short)450);
        HSSFCell cell6_0 = row6.createCell(0);
        HSSFCell cell6_1 = row6.createCell(1);
        HSSFCell cell6_2 = row6.createCell(2);
        HSSFCell cell6_3 = row6.createCell(3);
        HSSFCell cell6_4 = row6.createCell(4);
        HSSFCell cell6_5 = row6.createCell(5);
        HSSFCell cell6_6 = row6.createCell(6);
        HSSFCell cell6_7 = row6.createCell(7);
        HSSFCell cell6_8 = row6.createCell(8);
        HSSFCell cell6_9 = row6.createCell(9);
        HSSFCell cell6_10 = row6.createCell(10);
        HSSFCell cell6_11 = row6.createCell(11);
        HSSFCell cell6_12 = row6.createCell(12);
        HSSFCell cell6_13 = row6.createCell(13);
        HSSFCell cell6_14 = row6.createCell(14);
        cell6_0.setCellStyle(cellStyle);
        cell6_1.setCellStyle(cellStyle);
        cell6_2.setCellStyle(cellStyle);
        cell6_3.setCellStyle(cellStyle);
        cell6_4.setCellStyle(cellStyle);
        cell6_5.setCellStyle(cellStyle);
        cell6_6.setCellStyle(cellStyle);
        cell6_7.setCellStyle(cellStyle);
        cell6_8.setCellStyle(cellStyle);
        cell6_9.setCellStyle(cellStyle);
        cell6_10.setCellStyle(cellStyle);
        cell6_11.setCellStyle(cellStyle);
        cell6_12.setCellStyle(cellStyle);
        cell6_13.setCellStyle(cellStyle);
        cell6_14.setCellStyle(cellStyle);
        cell6_2.setCellValue("测量基准");
        cell6_3.setCellValue("X(单位°)");
        cell6_6.setCellValue("Y(单位°)");
        cell6_9.setCellValue("Z(单位°)");
        cell6_12.setCellValue("坐标(单位mm)");
        //第八行
        HSSFRow row7 = sheet.createRow(7);
        row7.setHeight((short)450);
        HSSFCell cell7_0 = row7.createCell(0);
        HSSFCell cell7_1 = row7.createCell(1);
        HSSFCell cell7_2 = row7.createCell(2);
        HSSFCell cell7_3 = row7.createCell(3);
        HSSFCell cell7_4 = row7.createCell(4);
        HSSFCell cell7_5 = row7.createCell(5);
        HSSFCell cell7_6 = row7.createCell(6);
        HSSFCell cell7_7 = row7.createCell(7);
        HSSFCell cell7_8 = row7.createCell(8);
        HSSFCell cell7_9 = row7.createCell(9);
        HSSFCell cell7_10 = row7.createCell(10);
        HSSFCell cell7_11 = row7.createCell(11);
        HSSFCell cell7_12 = row7.createCell(12);
        HSSFCell cell7_13 = row7.createCell(13);
        HSSFCell cell7_14 = row7.createCell(14);
        cell7_0.setCellStyle(cellStyle);
        cell7_1.setCellStyle(cellStyle);
        cell7_2.setCellStyle(cellStyle);
        cell7_3.setCellStyle(cellStyle);
        cell7_4.setCellStyle(cellStyle);
        cell7_5.setCellStyle(cellStyle);
        cell7_6.setCellStyle(cellStyle);
        cell7_7.setCellStyle(cellStyle);
        cell7_8.setCellStyle(cellStyle);
        cell7_9.setCellStyle(cellStyle);
        cell7_10.setCellStyle(cellStyle);
        cell7_11.setCellStyle(cellStyle);
        cell7_12.setCellStyle(cellStyle);
        cell7_13.setCellStyle(cellStyle);
        cell7_14.setCellStyle(cellStyle);
        cell7_1.setCellValue("测量项目");
        cell7_3.setCellValue("α");
        cell7_4.setCellValue("β");
        cell7_5.setCellValue("γ");
        cell7_6.setCellValue("α");
        cell7_7.setCellValue("β");
        cell7_8.setCellValue("γ");
        cell7_9.setCellValue("α");
        cell7_10.setCellValue("β");
        cell7_11.setCellValue("γ");
        cell7_12.setCellValue("x");
        cell7_13.setCellValue("y");
        cell7_14.setCellValue("z");
        //第九行
        HSSFRow row8 = sheet.createRow(8);
        row8.setHeight((short)450);

       /* for(int i= 0;i<15;i++){
            HSSFCell cell8_i = row8.createCell(i);
            cell8_i.setCellStyle(cellStyle);
        }*/
        HSSFCell cell8_0 = row8.createCell(0);
        HSSFCell cell8_1 = row8.createCell(1);
        HSSFCell cell8_2 = row8.createCell(2);
        HSSFCell cell8_3 = row8.createCell(3);
        HSSFCell cell8_4 = row8.createCell(4);
        HSSFCell cell8_5 = row8.createCell(5);
        HSSFCell cell8_6 = row8.createCell(6);
        HSSFCell cell8_7 = row8.createCell(7);
        HSSFCell cell8_8 = row8.createCell(8);
        HSSFCell cell8_9 = row8.createCell(9);
        HSSFCell cell8_10 = row8.createCell(10);
        HSSFCell cell8_11 = row8.createCell(11);
        HSSFCell cell8_12 = row8.createCell(12);
        HSSFCell cell8_13 = row8.createCell(13);
        HSSFCell cell8_14 = row8.createCell(14);
        cell8_0.setCellStyle(cellStyle);
        cell8_1.setCellStyle(cellStyle);
        cell8_2.setCellStyle(cellStyle);
        cell8_3.setCellStyle(cellStyle);
        cell8_4.setCellStyle(cellStyle);
        cell8_5.setCellStyle(cellStyle);
        cell8_6.setCellStyle(cellStyle);
        cell8_7.setCellStyle(cellStyle);
        cell8_8.setCellStyle(cellStyle);
        cell8_9.setCellStyle(cellStyle);
        cell8_10.setCellStyle(cellStyle);
        cell8_11.setCellStyle(cellStyle);
        cell8_12.setCellStyle(cellStyle);
        cell8_13.setCellStyle(cellStyle);
        cell8_14.setCellStyle(cellStyle);
        cell8_0.setCellValue("项目名称");
        cell8_1.setCellValue("设计值");
        cell8_3.setCellValue(xx);
        cell8_4.setCellValue(xy);
        cell8_5.setCellValue(xz);
        cell8_6.setCellValue(yx);
        cell8_7.setCellValue(yy);
        cell8_8.setCellValue(yz);
        cell8_9.setCellValue(zx);
        cell8_10.setCellValue(zy);
        cell8_11.setCellValue(zz);
        cell8_12.setCellValue(zx);
        cell8_13.setCellValue(yz);
        cell8_14.setCellValue(xx);



        //第十行
        HSSFRow row9 = sheet.createRow(9);
        row9.setHeight((short)450);
        HSSFCell cell9_0 = row9.createCell(0);
        HSSFCell cell9_1 = row9.createCell(1);
        HSSFCell cell9_2 = row9.createCell(2);
        HSSFCell cell9_3 = row9.createCell(3);
        HSSFCell cell9_4 = row9.createCell(4);
        HSSFCell cell9_5 = row9.createCell(5);
        HSSFCell cell9_6 = row9.createCell(6);
        HSSFCell cell9_7 = row9.createCell(7);
        HSSFCell cell9_8 = row9.createCell(8);
        HSSFCell cell9_9 = row9.createCell(9);
        HSSFCell cell9_10 = row9.createCell(10);
        HSSFCell cell9_11 = row9.createCell(11);
        HSSFCell cell9_12 = row9.createCell(12);
        HSSFCell cell9_13 = row9.createCell(13);
        HSSFCell cell9_14 = row9.createCell(14);
        cell9_0.setCellStyle(cellStyle);
        cell9_1.setCellStyle(cellStyle);
        cell9_2.setCellStyle(cellStyle);
        cell9_3.setCellStyle(cellStyle);
        cell9_4.setCellStyle(cellStyle);
        cell9_5.setCellStyle(cellStyle);
        cell9_6.setCellStyle(cellStyle);
        cell9_7.setCellStyle(cellStyle);
        cell9_8.setCellStyle(cellStyle);
        cell9_9.setCellStyle(cellStyle);
        cell9_10.setCellStyle(cellStyle);
        cell9_11.setCellStyle(cellStyle);
        cell9_12.setCellStyle(cellStyle);
        cell9_13.setCellStyle(cellStyle);
        cell9_14.setCellStyle(cellStyle);
        cell9_1.setCellValue("振前测量");
        cell9_2.setCellValue("测量值");
        cell9_3.setCellValue(xx);
        cell9_4.setCellValue(xy);
        cell9_5.setCellValue(xz);
        cell9_6.setCellValue(yx);
        cell9_7.setCellValue(yy);
        cell9_8.setCellValue(yz);
        cell9_9.setCellValue(zx);
        cell9_10.setCellValue(zy);
        cell9_11.setCellValue(zz);
        cell9_12.setCellValue(zx);
        cell9_13.setCellValue(yz);
        cell9_14.setCellValue(xx);
        //第十一行
        HSSFRow row10 = sheet.createRow(10);
        row10.setHeight((short)450);
        HSSFCell cell10_0 = row10.createCell(0);
        HSSFCell cell10_1 = row10.createCell(1);
        HSSFCell cell10_2 = row10.createCell(2);
        HSSFCell cell10_3 = row10.createCell(3);
        HSSFCell cell10_4 = row10.createCell(4);
        HSSFCell cell10_5 = row10.createCell(5);
        HSSFCell cell10_6 = row10.createCell(6);
        HSSFCell cell10_7 = row10.createCell(7);
        HSSFCell cell10_8 = row10.createCell(8);
        HSSFCell cell10_9 = row10.createCell(9);
        HSSFCell cell10_10 = row10.createCell(10);
        HSSFCell cell10_11 = row10.createCell(11);
        HSSFCell cell10_12 = row10.createCell(12);
        HSSFCell cell10_13 = row10.createCell(13);
        HSSFCell cell10_14 = row10.createCell(14);
        cell10_0.setCellStyle(cellStyle);
        cell10_1.setCellStyle(cellStyle);
        cell10_2.setCellStyle(cellStyle);
        cell10_3.setCellStyle(cellStyle);
        cell10_4.setCellStyle(cellStyle);
        cell10_5.setCellStyle(cellStyle);
        cell10_6.setCellStyle(cellStyle);
        cell10_7.setCellStyle(cellStyle);
        cell10_8.setCellStyle(cellStyle);
        cell10_9.setCellStyle(cellStyle);
        cell10_10.setCellStyle(cellStyle);
        cell10_11.setCellStyle(cellStyle);
        cell10_12.setCellStyle(cellStyle);
        cell10_13.setCellStyle(cellStyle);
        cell10_14.setCellStyle(cellStyle);
        cell10_2.setCellValue("变化值");
        cell10_3.setCellValue(xx);
        cell10_4.setCellValue(xy);
        cell10_5.setCellValue(xz);
        cell10_6.setCellValue(yx);
        cell10_7.setCellValue(yy);
        cell10_8.setCellValue(yz);
        cell10_9.setCellValue(zx);
        cell10_10.setCellValue(zy);
        cell10_11.setCellValue(zz);
        cell10_12.setCellValue(zx);
        cell10_13.setCellValue(yz);
        cell10_14.setCellValue(xx);

        //单元格范围 参数（int firstRow, int lastRow, int firstCol, int lastCol)
        CellRangeAddress cellRangeAddress0_0 =new CellRangeAddress(0, 0, 0, 14);
        CellRangeAddress cellRangeAddress1_9 =new CellRangeAddress(1, 1, 9, 10);
        CellRangeAddress cellRangeAddress1_11 =new CellRangeAddress(1, 1 ,11, 12);
        CellRangeAddress cellRangeAddress2_3 =new CellRangeAddress(2, 2, 3, 14);
        CellRangeAddress cellRangeAddress3_3 =new CellRangeAddress(3, 3, 3, 14);
        CellRangeAddress cellRangeAddress4_0 =new CellRangeAddress(4, 4, 0, 14);
        CellRangeAddress cellRangeAddress5_0 =new CellRangeAddress(5, 5, 3, 14);
        CellRangeAddress cellRangeAddress6_3 =new CellRangeAddress(6, 6, 3, 5);
        CellRangeAddress cellRangeAddress6_6 =new CellRangeAddress(6, 6, 6, 8);
        CellRangeAddress cellRangeAddress6_9 =new CellRangeAddress(6, 6, 9, 11);
        CellRangeAddress cellRangeAddress6_12 =new CellRangeAddress(6, 6, 12, 14);
        CellRangeAddress cellRangeAddress9_1 =new CellRangeAddress(9, 10, 1, 1);
        //在sheet里增加合并单元格
        sheet.addMergedRegion(cellRangeAddress0_0);
        sheet.addMergedRegion(cellRangeAddress1_9);
        sheet.addMergedRegion(cellRangeAddress1_11);
        sheet.addMergedRegion(cellRangeAddress2_3);
        sheet.addMergedRegion(cellRangeAddress3_3);
        sheet.addMergedRegion(cellRangeAddress4_0);
        sheet.addMergedRegion(cellRangeAddress5_0);
        sheet.addMergedRegion(cellRangeAddress6_3);
        sheet.addMergedRegion(cellRangeAddress6_6);
        sheet.addMergedRegion(cellRangeAddress6_9);
        sheet.addMergedRegion(cellRangeAddress6_12);
        sheet.addMergedRegion(cellRangeAddress9_1);

        ByteArrayOutputStream os = new ByteArrayOutputStream();
        try {
        wb.write(os);
        } catch (IOException e) {
        e.printStackTrace();
        }
        byte[] content = os.toByteArray();
        File file = new File(path);//Excel文件生成后存储的位置。
        OutputStream fos = null;
        try {
        fos = new FileOutputStream(file);
        fos.write(content);
        os.close();
        fos.close();
        } catch (Exception e) {
        e.printStackTrace();
        }
        return "success";
        }

public static DecimalFormat getDf() {
        return df;
        }

public static void setDf(DecimalFormat df) {
        ExcelUtilWriter.df = df;
        }

public static SimpleDateFormat getSdf() {
        return sdf;
        }

public static void setSdf(SimpleDateFormat sdf) {
        ExcelUtilWriter.sdf = sdf;
        }

public static DecimalFormat getNf() {
        return nf;
        }

public static void setNf(DecimalFormat nf) {
        ExcelUtilWriter.nf = nf;
        }
        }
