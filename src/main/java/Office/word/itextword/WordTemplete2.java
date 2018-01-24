package Office.word.itextword;

import com.lowagie.text.*;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.rtf.RtfWriter2;
import com.sun.xml.internal.txw2.Document;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 根据itext提供的java类库，构建word模板，并添加相应的内容，从而导出word报告；平台不相关
 * 需要引入iText-2.1.7.jar;iTextAsian.jar;iText-rtf-2.1.7.jar
 *
 * @author chen
 */
public class WordTemplete2 {
    private Document document;
    private BaseFont bfChinese;
    public BaseFont getBfChinese() {
        return bfChinese;
    }

    public void setBfChinese(BaseFont bfChinese) {
        this.bfChinese = bfChinese;
    }

    public Document getDocument() {
        return document;
    }

    public void setDocument(Document document) {
        this.document = document;
    }

    public WordTemplete2(){
        this.document = new Document(PageSize.A4);
        /*ByteArrayOutputStream baos = new ByteArrayOutputStream();
        *//** 建立一个书写器与document对象关联，通过书写器可以将文档写入到输出流中*//*
        RtfWriter2.getInstance(document, baos);
        document.open();*/

    }
    /**
     * @param filePath 建立一个书写器(Writer)与document对象关联，通过书写器(Writer)可以将文档写入到磁盘中
     * @throws DocumentException
     * @throws IOException
     */
    public void openDocument(String filePath) throws DocumentException,
            IOException {
        //       建立一个书写器(Writer)与document对象关联，通过书写器(Writer)可以将文档写入到磁盘中
        RtfWriter2.getInstance(this.document, new FileOutputStream(filePath));
        this.document.open();

    }
    /*
     * 测试控制清单
     * */
    public  void insertRiskControlTable() throws DocumentException{
        Table aTable = new Table(15,11);
        int width[] = {15, 15, 15, 8,  8,  8,  8,  8,  8,  8,  8 ,  8,  8,  8,  8};
        aTable.setWidths(width);// 设置每列所占比例
        aTable.setWidth(100); // 占页面宽度 90%
        aTable.setAlignment(Element.ALIGN_CENTER);// 居中显示
        aTable.setAlignment(Element.ALIGN_MIDDLE);// 纵向居中显示
        aTable.setAutoFillEmptyCells(true); // 自动填满
        aTable.setBorderWidth(0); // 边框宽度
        aTable.setBorderColor(new Color(241, 241, 241)); // 边框颜色
        aTable.setPadding(3);
        Font fontChinese = new Font(bfChinese, 15, Font.BOLD);

        Cell cells = new Cell(new Phrase("车辆数据对比分析报告",fontChinese));
        cells.setVerticalAlignment(Element.ALIGN_CENTER);
        cells.setHorizontalAlignment(Element.ALIGN_CENTER);
        cells.setBorderColor(new Color(0, 0, 0));
        cells.setColspan(15);
        cells.setBorderWidthTop(40);
        aTable.addCell(cells);

        aTable.addCell(new Cell());
        aTable.addCell(new Cell());
        aTable.addCell(new Cell());
        aTable.addCell(new Cell());
        aTable.addCell(new Cell());
        aTable.addCell(new Cell());
        aTable.addCell(new Cell());
        aTable.addCell(new Cell());
        aTable.addCell(new Cell());

        Font fontChinese2 = new Font(bfChinese, 10, Font.NORMAL);
        Cell cell1 = new Cell(new Phrase("测量值单位", fontChinese2));
        cell1.setVerticalAlignment(Element.ALIGN_CENTER);
        cell1.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell1.setBorderColor(new Color(0, 0, 0));
        cell1.setColspan(2);
        aTable.addCell(cell1);

        Cell cell2 = new Cell(new Phrase("变化量单位", fontChinese2));
        cell2.setVerticalAlignment(Element.ALIGN_CENTER);
        cell2.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell2.setBorderColor(new Color(0, 0, 0));
        cell2.setColspan(2);
        aTable.addCell(cell2);

        aTable.addCell(new Cell());
        aTable.addCell(new Cell());

        aTable.addCell(new Cell());
        aTable.addCell(new Cell());

        Font fontChinese3 = new Font(bfChinese, 13, Font.BOLD);
        Cell cell3 = new Cell(new Phrase("型号", fontChinese3));
        cell3.setVerticalAlignment(Element.ALIGN_CENTER);
        cell3.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell3.setBorderColor(new Color(0, 0, 0));
        aTable.addCell(cell3);

        Cell cell4 = new Cell(new Phrase("尖兵二号", fontChinese));
        cell4.setVerticalAlignment(Element.ALIGN_CENTER);
        cell4.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell4.setBorderColor(new Color(0, 0, 0));
        cell4.setColspan(12);
        aTable.addCell(cell4);

        aTable.addCell(new Cell());
        aTable.addCell(new Cell());

        Cell cell5 = new Cell(new Phrase("项目", fontChinese3));
        cell5.setVerticalAlignment(Element.ALIGN_CENTER);
        cell5.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell5.setBorderColor(new Color(0, 0, 0));
        aTable.addCell(cell5);

        Cell cell6 = new Cell(new Phrase("尖兵系列", fontChinese));
        cell6.setVerticalAlignment(Element.ALIGN_CENTER);
        cell6.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell6.setBorderColor(new Color(0, 0, 0));
        cell6.setColspan(12);
        aTable.addCell(cell6);

        Cell cell7 = new Cell(new Phrase("数据结果", fontChinese));
        cell7.setVerticalAlignment(Element.ALIGN_CENTER);
        cell7.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell7.setBorderColor(new Color(0, 0, 0));
        cell7.setColspan(15);
        aTable.addCell(cell7);

        aTable.addCell(new Cell());
        aTable.addCell(new Cell());
        aTable.addCell(new Cell());

        Cell cell8 = new Cell(new Phrase("基准镜", fontChinese3));
        cell8.setVerticalAlignment(Element.ALIGN_CENTER);
        cell8.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell8.setBorderColor(new Color(0, 0, 0));
        cell8.setColspan(12);
        aTable.addCell(cell8);


        aTable.addCell(new Cell());
        aTable.addCell(new Cell());

        Cell cell9 = new Cell(new Phrase("测量基准", fontChinese3));
        cell9.setVerticalAlignment(Element.ALIGN_CENTER);
        cell9.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell9.setBorderColor(new Color(0, 0, 0));
        aTable.addCell(cell9);

        Cell cell10 = new Cell(new Phrase("X", fontChinese3));
        cell10.setVerticalAlignment(Element.ALIGN_CENTER);
        cell10.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell10.setBorderColor(new Color(0, 0, 0));
        cell10.setColspan(3);
        aTable.addCell(cell10);

        Cell cell11 = new Cell(new Phrase("Y", fontChinese3));
        cell11.setVerticalAlignment(Element.ALIGN_CENTER);
        cell11.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell11.setBorderColor(new Color(0, 0, 0));
        cell11.setColspan(3);
        aTable.addCell(cell11);

        Cell cell12 = new Cell(new Phrase("Z", fontChinese3));
        cell12.setVerticalAlignment(Element.ALIGN_CENTER);
        cell12.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell12.setBorderColor(new Color(0, 0, 0));
        cell12.setColspan(3);
        aTable.addCell(cell12);


        Cell cell13 = new Cell(new Phrase("坐标", fontChinese3));
        cell13.setVerticalAlignment(Element.ALIGN_CENTER);
        cell13.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell13.setBorderColor(new Color(0, 0, 0));
        cell13.setColspan(3);
        aTable.addCell(cell13);

        aTable.addCell(new Cell());
        Cell cell14 = new Cell(new Phrase("测量项目", fontChinese3));

        cell14.setVerticalAlignment(Element.ALIGN_CENTER);
        cell14.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell14.setBorderColor(new Color(0, 0, 0));
        aTable.addCell(cell14);

        aTable.addCell(new Cell());

        for(int i=0;i<3;i++){
            aTable.addCell(getcell("α",fontChinese2));
            aTable.addCell(getcell("β",fontChinese2));
            aTable.addCell(getcell("γ",fontChinese2));
        }
        aTable.addCell(getcell("X",fontChinese2));
        aTable.addCell(getcell("Y",fontChinese2));
        aTable.addCell(getcell("Z",fontChinese2));

        for(int q=0;q<3;q++){
            Cell cell15 = new Cell(new Phrase("项目名称", fontChinese3));
            cell15.setVerticalAlignment(Element.ALIGN_CENTER);
            cell15.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell15.setBorderColor(new Color(0, 0, 0));
            aTable.addCell(cell15);

            Cell cell16 = new Cell(new Phrase("设计值", fontChinese3));
            cell16.setVerticalAlignment(Element.ALIGN_CENTER);
            cell16.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell16.setBorderColor(new Color(0, 0, 0));
            aTable.addCell(cell16);

            aTable.addCell(new Cell());

            for(int i=0;i<12;i++){
                aTable.addCell(getcell("1",fontChinese2));
            }
            aTable.addCell(new Cell());

            Cell cell17 = new Cell(new Phrase("振动前", fontChinese3));
            cell17.setVerticalAlignment(Element.ALIGN_CENTER);
            cell17.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell17.setBorderColor(new Color(0, 0, 0));
            cell17.setRowspan(2);
            aTable.addCell(cell17);

            Cell cell18 = new Cell(new Phrase("测量值", fontChinese3));
            cell18.setVerticalAlignment(Element.ALIGN_CENTER);
            cell18.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell18.setBorderColor(new Color(0, 0, 0));

            aTable.addCell(cell18);

            for(int i=0;i<12;i++){
                aTable.addCell(getcell("1",fontChinese2));
            }
            aTable.addCell(new Cell());


            Cell cell19 = new Cell(new Phrase("变化值", fontChinese3));
            cell19.setVerticalAlignment(Element.ALIGN_CENTER);
            cell19.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell19.setBorderColor(new Color(0, 0, 0));

            aTable.addCell(cell19);
            for(int i=0;i<12;i++){
                aTable.addCell(getcell("1",fontChinese2));
            }
        }




        document.add(aTable);
        document.add(new Paragraph("\n"));
    }


    Cell cell14;
    public Cell getcell(String name, Font fontChinese) throws BadElementException {
        cell14= new Cell(new Phrase(name, fontChinese));
        cell14.setVerticalAlignment(Element.ALIGN_CENTER);
        cell14.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell14.setBorderColor(new Color(0, 0, 0));
        return cell14;
    }


    public void closeDocument() throws DocumentException{
        this.document.close();
    }

    public static void main(String[] args) throws DocumentException, IOException {
        WordTemplete2 wt = new WordTemplete2();
        wt.openDocument("d:\\dome1.doc");
        wt.insertRiskControlTable();
        wt.closeDocument();
    }
}