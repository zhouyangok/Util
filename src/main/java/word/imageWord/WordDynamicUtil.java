package word.imageWord;

import Excel.FrontSpaceSeries;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import word.template.CustomXWPFDocument;

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
 * @Desciption:动态生成表格，用于创建带图片的表格
 * @Date 2017-12-18 8:53
 */
public class WordDynamicUtil {

    public static void main(String[] args) throws IOException {
        List<FrontSpaceSeries> list = new ArrayList<>();
        FrontSpaceSeries frontSpaceSeries = new FrontSpaceSeries();
        frontSpaceSeries.setSeriesName("神舟系列");
        frontSpaceSeries.setTypeName("神舟三号");
        frontSpaceSeries.setMeasureStatus("振前测量");
        frontSpaceSeries.setBenchmarkCmName("基准立方镜");
        frontSpaceSeries.setMeasureCmName("太阳立方镜");
        frontSpaceSeries.setMeasureStatus("太阳立方镜");
        frontSpaceSeries.setMeasureModel("太阳立方镜");
        frontSpaceSeries.setMeasureProcessName("过程1");
        frontSpaceSeries.setMeasureOrder("1");
        frontSpaceSeries.setXx("1");
        frontSpaceSeries.setXy("2");
        frontSpaceSeries.setXz("3");
        frontSpaceSeries.setYx("1");
        frontSpaceSeries.setYy("2");
        frontSpaceSeries.setYz("3");
        frontSpaceSeries.setZx("1");
        frontSpaceSeries.setZy("2");
        frontSpaceSeries.setZz("3");
        frontSpaceSeries.setX0("1");
        frontSpaceSeries.setY0("2");
        frontSpaceSeries.setZ0("3");
        list.add(frontSpaceSeries);
        list.add(frontSpaceSeries);
        list.add(frontSpaceSeries);
        list.add(frontSpaceSeries);

        generateWord(list);
    }

    public static void generateWord(List<FrontSpaceSeries> list) throws IOException {
        //Blank Document  
        XWPFDocument document = new XWPFDocument();

        //Write the Document in file system  

        String filePath = "D:\\generateWord\\imageWord\\";
        String fileName = System.currentTimeMillis() + ".docx";
        FileOutputStream out = new FileOutputStream(new File(filePath + fileName));

        try {
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
            XWPFTableRow comTableRowOne = ComTable.getRow(0);
            comTableRowOne.getCell(0).setText("系列");
            comTableRowOne.addNewTableCell().setText("型号");
            comTableRowOne.addNewTableCell().setText("测量状态");
            comTableRowOne.addNewTableCell().setText("基准立方镜");
            comTableRowOne.addNewTableCell().setText("太阳立方镜");
            comTableRowOne.addNewTableCell().setText("测量模式");
            comTableRowOne.addNewTableCell().setText("测量过程");
            comTableRowOne.addNewTableCell().setText("测量次序");
            comTableRowOne.addNewTableCell().setText("XX");
            comTableRowOne.addNewTableCell().setText("XY");
            comTableRowOne.addNewTableCell().setText("XZ");
            comTableRowOne.addNewTableCell().setText("YX");
            comTableRowOne.addNewTableCell().setText("YY");
            comTableRowOne.addNewTableCell().setText("YZ");
            comTableRowOne.addNewTableCell().setText("ZX");
            comTableRowOne.addNewTableCell().setText("ZY");
            comTableRowOne.addNewTableCell().setText("ZZ");
            comTableRowOne.addNewTableCell().setText("X0");
            comTableRowOne.addNewTableCell().setText("Y0");
            comTableRowOne.addNewTableCell().setText("Z0");
            for (int i = 0; i < list.size(); i++) {
                XWPFTableRow comTableRowTwo = ComTable.createRow();
                comTableRowTwo.getCell(1).setText(list.get(i).getTypeName());
                comTableRowTwo.getCell(0).setText(list.get(i).getSeriesName());
                comTableRowTwo.getCell(2).setText(list.get(i).getMeasureStatus());
                comTableRowTwo.getCell(3).setText(list.get(i).getBenchmarkCmName());
                comTableRowTwo.getCell(4).setText(list.get(i).getMeasureCmName());
                comTableRowTwo.getCell(5).setText(list.get(i).getMeasureModel());
                comTableRowTwo.getCell(6).setText(list.get(i).getMeasureProcessName());
                comTableRowTwo.getCell(7).setText(list.get(i).getMeasureOrder());
                comTableRowTwo.getCell(8).setText(list.get(i).getXx());
                comTableRowTwo.getCell(9).setText(list.get(i).getXy());
                comTableRowTwo.getCell(10).setText(list.get(i).getXz());
                comTableRowTwo.getCell(11).setText(list.get(i).getYx());
                comTableRowTwo.getCell(12).setText(list.get(i).getYy());
                comTableRowTwo.getCell(13).setText(list.get(i).getYz());
                comTableRowTwo.getCell(14).setText(list.get(i).getZx());
                comTableRowTwo.getCell(15).setText(list.get(i).getZy());
                comTableRowTwo.getCell(16).setText(list.get(i).getZz());
                comTableRowTwo.getCell(17).setText(list.get(i).getX0());
                comTableRowTwo.getCell(18).setText(list.get(i).getY0());
                comTableRowTwo.getCell(19).setText(list.get(i).getZ0());
            }


            CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
            XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);

            

//            =========添加图片到word中==============================================

          /*  IMGGenerate docc = null;
            docc = new IMGGenerate();

            byte[] byteArray = (byte[]) WordUtil.inputStream2ByteArray(new FileInputStream("D:\\generateWord\\echarts3.png"),true);
            ByteArrayInputStream byteInputStream = new ByteArrayInputStream(byteArray);
            XWPFParagraph xwpfParagraph =  docc.createParagraph();
            XWPFRun run = xwpfParagraph.createRun();

            String text = "";
            text += run.getText(0);
            try {
                String ind = "123";
                docc.createPicture(ind, 800, 500, xwpfParagraph);
                System.out.println("图片替换");
            } catch (Exception e) {
                e.printStackTrace();
            }
            
            run.setText(text, 0);
            IMGGenerate doc = new IMGGenerate();

            doc.createPicture("1", 2000, 500, paragraph);*/

            XWPFParagraph firstParagraph = document.createParagraph();
            XWPFRun run = firstParagraph.createRun();
            run.setText("");
            run.setColor("696969");
            run.setFontSize(16);

//=====================================================================
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
        } catch (Exception e) {
            e.printStackTrace();
        }

    }


    /**
     * 根据图片类型，取得对应的图片类型代码
     *
     * @param picType
     * @return int
     */
    private static int getPictureType(String picType) {
        int res = CustomXWPFDocument.PICTURE_TYPE_PICT;
        if (picType != null) {
            if (picType.equalsIgnoreCase("png")) {
                res = CustomXWPFDocument.PICTURE_TYPE_PNG;
            } else if (picType.equalsIgnoreCase("dib")) {
                res = CustomXWPFDocument.PICTURE_TYPE_DIB;
            } else if (picType.equalsIgnoreCase("emf")) {
                res = CustomXWPFDocument.PICTURE_TYPE_EMF;
            } else if (picType.equalsIgnoreCase("jpg") || picType.equalsIgnoreCase("jpeg")) {
                res = CustomXWPFDocument.PICTURE_TYPE_JPEG;
            } else if (picType.equalsIgnoreCase("wmf")) {
                res = CustomXWPFDocument.PICTURE_TYPE_WMF;
            }
        }
        return res;
    }

}
