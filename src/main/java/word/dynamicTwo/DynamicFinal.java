package word.dynamicTwo;

import Excel.FrontSpaceSeries;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

/**
 * Created with IntelliJ IDEA.
 *
 * @Author medo_zy
 * @Desciption:
 * @Date 2017-12-18 8:53
 */
public class DynamicFinal {
    
      public static void main(String[] args) throws Exception {
        List<FrontSpaceSeries> list = new ArrayList<>();
        FrontSpaceSeries frontSpaceSeries = new FrontSpaceSeries();
        frontSpaceSeries.setSeriesName("神舟系列");
        frontSpaceSeries.setTypeName("神舟三号");
        frontSpaceSeries.setMeasureStatus("振前测量");
        frontSpaceSeries.setBenchmarkCmName("基准立方镜");
        frontSpaceSeries.setMeasureCmName("太阳立方镜");
        frontSpaceSeries.setMeasureModel("太阳立方镜");
        frontSpaceSeries.setMeasureProcessName("过程1");
        frontSpaceSeries.setMeasureOrder("1");
        frontSpaceSeries.setXx("1");
        frontSpaceSeries.setXy("2");
        frontSpaceSeries.setXz("3");
        frontSpaceSeries.setYx("3");
        frontSpaceSeries.setYy("3");
        frontSpaceSeries.setYz("3");
        frontSpaceSeries.setZx("3");
        frontSpaceSeries.setZy("3");
        frontSpaceSeries.setZz("2.432");
        frontSpaceSeries.setX0("3.3423");
        frontSpaceSeries.setY0("3.4564");
        frontSpaceSeries.setZ0("7.23423");
        list.add(frontSpaceSeries);
        list.add(frontSpaceSeries);
        list.add(frontSpaceSeries);
        list.add(frontSpaceSeries);
        list.add(frontSpaceSeries);
          DynamicFinal.generate(list);
    }
    
    public static void generate(List<FrontSpaceSeries> list) throws IOException, XmlException {
        XWPFDocument document = new XWPFDocument();

        String filePath = "D:\\generateWord\\";
        String fileName = System.currentTimeMillis()+".docx";

        //Write the Document in file system  
        FileOutputStream out = new FileOutputStream(new File(filePath+fileName));


        //添加标题  
        XWPFParagraph titleParagraph = document.createParagraph();
        //设置段落居中  
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun titleParagraphRun = titleParagraph.createRun();
        titleParagraphRun.setText("数据分析报表");
        titleParagraphRun.setColor("000000");
        titleParagraphRun.setFontSize(20);

        //两个表格之间加个换行  
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun paragraphRun = paragraph.createRun();
        paragraphRun.setText("\r");


        //工作经历表格  
        XWPFTable ComTable = document.createTable();
        
        //列宽自动分割  
        CTTblWidth comTableWidth = ComTable.getCTTbl().addNewTblPr().addNewTblW();
        comTableWidth.setType(STTblWidth.DXA);
        comTableWidth.setW(BigInteger.valueOf(9072));
       
        
        

        //表格第一行
        XWPFTableRow comTableRowHead = ComTable.getRow(0);
        comTableRowHead.getCell(0).setText("测量数据对比分析报告");
        comTableRowHead.addNewTableCell();
        comTableRowHead.addNewTableCell();
        comTableRowHead.addNewTableCell();
        comTableRowHead.addNewTableCell();
        comTableRowHead.addNewTableCell();
        comTableRowHead.addNewTableCell();
        comTableRowHead.addNewTableCell();
        comTableRowHead.addNewTableCell();
        comTableRowHead.addNewTableCell();
        comTableRowHead.addNewTableCell();
        comTableRowHead.addNewTableCell();
        comTableRowHead.addNewTableCell();
        comTableRowHead.addNewTableCell();
        comTableRowHead.addNewTableCell();
        comTableRowHead.addNewTableCell();
       
       // 表格第二行
         XWPFTableRow comTableRowTwo = ComTable.createRow();
       
        comTableRowTwo.getCell(9).setText("测量值单位");
        comTableRowTwo.getCell(11).setText("测量值单位");
      
        // 表格第三行
         XWPFTableRow comTableRowTree = ComTable.createRow();
        comTableRowTree.getCell(2).setText("型号");
        comTableRowTree.getCell(11).setText("测量型号");
         // 表格第四行
         XWPFTableRow comTableRowFour = ComTable.createRow();
        comTableRowFour.getCell(2).setText("项目");
        comTableRowFour.getCell(11).setText("神舟系列");
         // 表格第五行
         XWPFTableRow comTableRowFive = ComTable.createRow();
        comTableRowFive.getCell(2).setText("数据结果");
          // 表格第六行
         XWPFTableRow comTableRowSix = ComTable.createRow();
        comTableRowFive.getCell(2).setText("基准镜");
          // 表格第七行
         XWPFTableRow comTableRowSeven = ComTable.createRow();
        comTableRowSeven.getCell(2).setText("测量基准");
        comTableRowSeven.getCell(3).setText("X");
        comTableRowSeven.getCell(6).setText("Y");
        comTableRowSeven.getCell(9).setText("Z");
        comTableRowSeven.getCell(12).setText("坐标");
         // 表格第八行
         XWPFTableRow comTableRowEight = ComTable.createRow();
        comTableRowEight.getCell(1).setText("测量项目");
        comTableRowEight.getCell(3).setText("α");
        comTableRowEight.getCell(4).setText("β");
        comTableRowEight.getCell(5).setText("γ");
        comTableRowEight.getCell(6).setText("α");
        comTableRowEight.getCell(7).setText("β");
        comTableRowEight.getCell(8).setText("γ");
        comTableRowEight.getCell(9).setText("α");
        comTableRowEight.getCell(10).setText("β");
        comTableRowEight.getCell(11).setText("γ");
        comTableRowEight.getCell(12).setText("X");
        comTableRowEight.getCell(13).setText("Y");
        comTableRowEight.getCell(14).setText("Z");
        //
        
      
        for (int i = 0; i < list.size(); i++) {
            XWPFTableRow comTableRowX = ComTable.createRow();
            comTableRowX.getCell(0).setText("项目名称");
            comTableRowX.getCell(1).setText("设计值");
            comTableRowX.getCell(2).setText(list.get(i).getXx());
            comTableRowX.getCell(3).setText(list.get(i).getXy());
            comTableRowX.getCell(4).setText(list.get(i).getXz());
            comTableRowX.getCell(5).setText(list.get(i).getYx());
            comTableRowX.getCell(6).setText(list.get(i).getYy());
            comTableRowX.getCell(7).setText(list.get(i).getYz());
            comTableRowX.getCell(8).setText(list.get(i).getZx());
            comTableRowX.getCell(9).setText(list.get(i).getZy());
            comTableRowX.getCell(10).setText(list.get(i).getZz());
            comTableRowX.getCell(11).setText(list.get(i).getX0());
            comTableRowX.getCell(12).setText(list.get(i).getY0());
            comTableRowX.getCell(13).setText(list.get(i).getZ0());
            XWPFTableRow comTableRowY = ComTable.createRow();
            comTableRowY.getCell(1).setText("振前测量");
            comTableRowY.getCell(2).setText(list.get(i).getXx());
            comTableRowY.getCell(3).setText(list.get(i).getXy());
            comTableRowY.getCell(4).setText(list.get(i).getXz());
            comTableRowY.getCell(5).setText(list.get(i).getYx());
            comTableRowY.getCell(6).setText(list.get(i).getYy());
            comTableRowY.getCell(7).setText(list.get(i).getYz());
            comTableRowY.getCell(8).setText(list.get(i).getZx());
            comTableRowY.getCell(9).setText(list.get(i).getZy());
            comTableRowY.getCell(10).setText(list.get(i).getZz());
            comTableRowY.getCell(11).setText(list.get(i).getX0());
            comTableRowY.getCell(12).setText(list.get(i).getY0());
            comTableRowY.getCell(13).setText(list.get(i).getZ0());
           
        }

        WordParam.mergeCellsHorizontal(ComTable,0,0,10);
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
        
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);

        //添加页眉  
        CTP ctpHeader = CTP.Factory.newInstance();
        CTR ctrHeader = ctpHeader.addNewR();
        CTText ctHeader = ctrHeader.addNewT();
        String headerText = "航天器精测数据与自动分析管理系统";
        ctHeader.setStringValue(headerText);
        XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, document);
        //设置为右对齐  
        headerParagraph.setAlignment(ParagraphAlignment.RIGHT);
        XWPFParagraph[] parsHeader = new XWPFParagraph[1];
        parsHeader[0] = headerParagraph;
        policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, parsHeader);


        //添加页脚  
        CTP ctpFooter = CTP.Factory.newInstance();
        CTR ctrFooter = ctpFooter.addNewR();
        CTText ctFooter = ctrFooter.addNewT();
        String footerText = "航天器精测数据与自动分析管理系统";
        ctFooter.setStringValue(footerText);
        XWPFParagraph footerParagraph = new XWPFParagraph(ctpFooter, document);
        headerParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFParagraph[] parsFooter = new XWPFParagraph[1];
        parsFooter[0] = footerParagraph;
        policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, parsFooter);


        document.write(out);
        out.close();
        System.out.println("create_table document written success.");
    }
    
}
