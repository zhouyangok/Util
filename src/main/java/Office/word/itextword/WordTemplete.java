package Office.word.itextword;

import com.lowagie.text.*;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.rtf.RtfWriter2;
import com.medo.web.iim.entity.FrontDesignVo;
import com.medo.web.iim.entity.FrontSpaceSeries;
import org.apache.commons.lang.StringUtils;
import org.springframework.stereotype.Service;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.util.List;

/**
 * 根据itext提供的java类库，构建word模板，并添加相应的内容，从而导出word报告；平台不相关
 * 需要引入iText-2.1.7.jar;iTextAsian.jar;iText-rtf-2.1.7.jar
 *
 * @author chen
 */

@SuppressWarnings("unchecked")
@Service
public class WordTemplete {
    private Document document;
    private BaseFont bfChinese;
    private  Cell cells;
    private Cell cellTile;
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

    public WordTemplete(){
        document = new Document(PageSize.A4);
        /*ByteArrayOutputStream baos = new ByteArrayOutputStream();
        // 建立一个书写器与document对象关联，通过书写器可以将文档写入到输出流中*//*
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
        document = new Document(PageSize.A4);
        //建立一个书写器(Writer)与document对象关联，通过书写器(Writer)可以将文档写入到磁盘中
        RtfWriter2.getInstance(this.document, new FileOutputStream(filePath));
        this.document.open();

    }

    /**
     * @param titleStr 标题
     * @param fontsize 字体大小
     * @param fontStyle 字体样式
     * @param elementAlign 对齐方式
     * @throws DocumentException
     */
    public void insertTitle(String titleStr,int fontsize,int fontStyle,int elementAlign) throws DocumentException{
        Font titleFont = new Font(this.bfChinese, fontsize, fontStyle);
        Paragraph title = new Paragraph(titleStr);
        // 设置标题格式对齐方式
        title.setAlignment(elementAlign);
        title.setFont(titleFont);

         document.add(title);
    }

    /**
     * @param contextStr 内容
     * @param fontsize 字体大小
     * @param fontStyle 字体样式
     * @param elementAlign 对齐方式
     * @throws DocumentException
     */
    public void insertContext(String contextStr,int fontsize,int fontStyle,int elementAlign) throws DocumentException{
        // 正文字体风格
        Font contextFont = new Font(bfChinese, fontsize, fontStyle);
        Paragraph context = new Paragraph(contextStr);
        //设置行距
        context.setLeading(30f);
        // 正文格式左对齐
        context.setAlignment(elementAlign);
        context.setFont(contextFont);
        // 离上一段落（标题）空的行数
        context.setSpacingBefore(5);
        // 设置第一行空的列数
        context.setFirstLineIndent(20);
        document.add(context);
    }

    /*
     * 测试控制清单
     * */
    public  void insertRiskControlTable(List<FrontSpaceSeries> list) throws DocumentException{
        Table aTable = new Table(19,list.size());
        int width[] = { 6, 9, 9, 9, 9, 9, 9, 9, 6, 6, 6, 6 , 6 , 6 ,  6 ,6, 6 ,  6 ,6};
        String[] colName = { "名称", "型号名", "测量阶段", "基准立方镜", "测量立方镜", "测量模式", "测量过程", "次序", "XX", "XY", "XZ", "YX" , "YY" , "YZ" ,  "ZX" ,"ZY","X0","Y0","Z0"};
        aTable.setWidths(width);// 设置每列所占比例
        aTable.setWidth(100); // 占页面宽度 90%
        aTable.setAlignment(Element.ALIGN_CENTER);// 居中显示
        aTable.setAlignment(Element.ALIGN_MIDDLE);// 纵向居中显示
        aTable.setAutoFillEmptyCells(true); // 自动填满
        aTable.setBorderWidth(0); // 边框宽度
        aTable.setBorderColor(new Color(241, 241, 241)); // 边框颜色

        Font fontChinese = new Font(bfChinese, 10, Font.NORMAL);
        for(String name : colName ){
            aTable.addCell(getCellTile(name,fontChinese));
        }
        for(int i=0;i<list.size();i++){
            FrontSpaceSeries frontSpaceSeries=list.get(i);
            aTable.addCell(getCell(frontSpaceSeries.getSeriesName(),fontChinese));
            aTable.addCell(getCell(frontSpaceSeries.getTypeName(),fontChinese));
            aTable.addCell(getCell(frontSpaceSeries.getMeasureStatus(),fontChinese));
            aTable.addCell(getCell(frontSpaceSeries.getBenchmarkCmName(),fontChinese));
            aTable.addCell(getCell(frontSpaceSeries.getMeasureCmName(),fontChinese));
            aTable.addCell(getCell(frontSpaceSeries.getMeasureModel(),fontChinese));
            aTable.addCell(getCell(frontSpaceSeries.getMeasureProcessName(),fontChinese));
            aTable.addCell(getCell(frontSpaceSeries.getMeasureOrder(),fontChinese));
            aTable.addCell(getCell(frontSpaceSeries.getXx(),fontChinese));
            aTable.addCell(getCell(frontSpaceSeries.getXy(),fontChinese));
            aTable.addCell(getCell(frontSpaceSeries.getXz(),fontChinese));
            aTable.addCell(getCell(frontSpaceSeries.getYx(),fontChinese));
            aTable.addCell(getCell(frontSpaceSeries.getYy(),fontChinese));
            aTable.addCell(getCell(frontSpaceSeries.getYz(),fontChinese));
            aTable.addCell(getCell(frontSpaceSeries.getZx(),fontChinese));
            aTable.addCell(getCell(frontSpaceSeries.getZy(),fontChinese));
            aTable.addCell(getCell(frontSpaceSeries.getX0(),fontChinese));
            aTable.addCell(getCell(frontSpaceSeries.getY0(),fontChinese));
            aTable.addCell(getCell(frontSpaceSeries.getZ0(),fontChinese));
        }
        document.add(aTable);
        document.add(new Paragraph("\n"));
    }


    private Cell getCellTile(String info, Font fontChinese) throws BadElementException {
        cellTile = new Cell(new Phrase(info, fontChinese));
        cellTile.setVerticalAlignment(Element.ALIGN_CENTER);
        cellTile.setHorizontalAlignment(Element.ALIGN_CENTER);
        cellTile.setBorderColor(new Color(0, 0, 0));
        cellTile.setBackgroundColor(new Color(241, 241, 241));
        return  cellTile;
    }

    private Cell getCell(String info, Font fontChinese) throws BadElementException {
        cells = new Cell(new Phrase(info, fontChinese));
        cells.setVerticalAlignment(Element.ALIGN_CENTER);
        cells.setHorizontalAlignment(Element.ALIGN_CENTER);
        cells.setBorderColor(new Color(0, 0, 0));
        return  cells;
    }

    /**
     * @param imgUrl 图片路径
     * @param imageAlign 显示位置
     * @param height 显示高度
     * @param weight 显示宽度
     * @param percent 显示比例
     * @param heightPercent 显示高度比例
     * @param weightPercent 显示宽度比例
     * @param rotation 显示图片旋转角度
     * @throws MalformedURLException
     * @throws IOException
     * @throws DocumentException
     */
    public void insertImg(String imgUrl,int imageAlign,int height,int weight,int percent,int heightPercent,int weightPercent,int rotation) throws MalformedURLException, IOException, DocumentException{
        //添加图片
        Image img = Image.getInstance(imgUrl);
        if(img==null)
            return;
        img.setAbsolutePosition(0, 0);
        img.setAlignment(imageAlign);
        img.scaleAbsolute(height, weight);
        img.scalePercent(percent);
        img.scalePercent(heightPercent, weightPercent);
        img.setRotation(rotation);
        document.add(img);
    }

    public void closeDocument() throws DocumentException{
        document.close();
        document=null;
    }

    public  void insertControlTable(List<FrontDesignVo> fdList) throws DocumentException{
        Table aTable = new Table(15,8);
        int width[] = {15, 12, 12, 9,  9,  9,  9,  9,  9,  9,  9 ,  9,  9,  9,  9};
        aTable.setWidths(width);// 设置每列所占比例
        aTable.setWidth(100); // 占页面宽度 90%
        aTable.setAlignment(Element.ALIGN_CENTER);// 居中显示
        aTable.setAlignment(Element.ALIGN_MIDDLE);// 纵向居中显示
        aTable.setAutoFillEmptyCells(true); // 自动填满
        aTable.setBorderWidth(0); // 边框宽度
        aTable.setBorderColor(new Color(241, 241, 241)); // 边框颜色
        aTable.setPadding(2);
        Font fontChinese = new Font(bfChinese, 18, Font.BOLD);

        Cell cells = new Cell(new Phrase("测量数据对比分析报告",fontChinese));
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
        Font fontChinese3 = new Font(bfChinese, 10, Font.BOLD);
/*
    aTable.addCell(new Cell());
    aTable.addCell(new Cell());

    Font fontChinese3 = new Font(bfChinese, 12, Font.BOLD);
    Cell cell3 = new Cell(new Phrase("型号", fontChinese3));
    cell3.setVerticalAlignment(Element.ALIGN_CENTER);
    cell3.setHorizontalAlignment(Element.ALIGN_CENTER);
    cell3.setBorderColor(new Color(0, 0, 0));
    aTable.addCell(cell3);

    Cell cell4 = new Cell(new Phrase("尖兵二号", fontChinese3));
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

    Cell cell6 = new Cell(new Phrase("尖兵系列", fontChinese3));
    cell6.setVerticalAlignment(Element.ALIGN_CENTER);
    cell6.setHorizontalAlignment(Element.ALIGN_CENTER);
    cell6.setBorderColor(new Color(0, 0, 0));
    cell6.setColspan(12);
    aTable.addCell(cell6);*/

    Cell cell7 = new Cell(new Phrase("数据结果", fontChinese3));
    cell7.setVerticalAlignment(Element.ALIGN_CENTER);
    cell7.setHorizontalAlignment(Element.ALIGN_CENTER);
    cell7.setBorderColor(new Color(0, 0, 0));
    cell7.setColspan(15);
    aTable.addCell(cell7);

    /*aTable.addCell(new Cell());
    aTable.addCell(new Cell());
    aTable.addCell(new Cell());

    Cell cell8 = new Cell(new Phrase("基准镜", fontChinese3));
    cell8.setVerticalAlignment(Element.ALIGN_CENTER);
    cell8.setHorizontalAlignment(Element.ALIGN_CENTER);
    cell8.setBorderColor(new Color(0, 0, 0));
    cell8.setColspan(12);
    aTable.addCell(cell8);*/


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
    cell10.setBackgroundColor(new Color(215,215,215));
    cell10.setBorderColor(new Color(0, 0, 0));
    cell10.setColspan(3);
    aTable.addCell(cell10);

    Cell cell11 = new Cell(new Phrase("Y", fontChinese3));
    cell11.setVerticalAlignment(Element.ALIGN_CENTER);
    cell11.setHorizontalAlignment(Element.ALIGN_CENTER);
    cell11.setBackgroundColor(new Color(215,215,215));
    cell11.setBorderColor(new Color(0, 0, 0));
    cell11.setColspan(3);
    aTable.addCell(cell11);

    Cell cell12 = new Cell(new Phrase("Z", fontChinese3));
    cell12.setVerticalAlignment(Element.ALIGN_CENTER);
    cell12.setHorizontalAlignment(Element.ALIGN_CENTER);
    cell12.setBackgroundColor(new Color(215,215,215));
    cell12.setBorderColor(new Color(0, 0, 0));
    cell12.setColspan(3);
    aTable.addCell(cell12);


    Cell cell13 = new Cell(new Phrase("坐标", fontChinese3));
    cell13.setVerticalAlignment(Element.ALIGN_CENTER);
    cell13.setHorizontalAlignment(Element.ALIGN_CENTER);
    cell13.setBackgroundColor(new Color(215,215,215));
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

    for(FrontDesignVo vo : fdList){
        Cell cell15 = new Cell(new Phrase(vo.getSeriesName()+"("+vo.getTypeName()+")", fontChinese3));
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

        aTable.addCell(getcell(vo.getT12Xx(),fontChinese2));
        aTable.addCell(getcell(vo.getT12Xy(),fontChinese2));
        aTable.addCell(getcell(vo.getT12Xz(),fontChinese2));

        aTable.addCell(getcell(vo.getT12Yx(),fontChinese2));
        aTable.addCell(getcell(vo.getT12Yy(),fontChinese2));
        aTable.addCell(getcell(vo.getT12Yz(),fontChinese2));

        aTable.addCell(getcell(vo.getT12Zx(),fontChinese2));
        aTable.addCell(getcell(vo.getT12Zy(),fontChinese2));
        aTable.addCell(getcell(vo.getT12Zz(),fontChinese2));

        aTable.addCell(getcell(vo.getT12X(),fontChinese2));
        aTable.addCell(getcell(vo.getT12Y(),fontChinese2));
        aTable.addCell(getcell(vo.getT12Z(),fontChinese2));

        Cell cell20 = new Cell(new Phrase(vo.getBenchCmName(), fontChinese3));
        cell20.setVerticalAlignment(Element.ALIGN_CENTER);
        cell20.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell20.setBorderColor(new Color(0, 0, 0));
        cell20.setRowspan(2);
        aTable.addCell(cell20);

        Cell cell17 = new Cell(new Phrase(vo.getMeasureStatus(), fontChinese3));
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

        aTable.addCell(getcell(vo.getXx(),fontChinese2));
        aTable.addCell(getcell(vo.getXy(),fontChinese2));
        aTable.addCell(getcell(vo.getXz(),fontChinese2));

        aTable.addCell(getcell(vo.getYx(),fontChinese2));
        aTable.addCell(getcell(vo.getYy(),fontChinese2));
        aTable.addCell(getcell(vo.getYz(),fontChinese2));

        aTable.addCell(getcell(vo.getZx(),fontChinese2));
        aTable.addCell(getcell(vo.getZy(),fontChinese2));
        aTable.addCell(getcell(vo.getZz(),fontChinese2));

        aTable.addCell(getcell(vo.getX0(),fontChinese2));
        aTable.addCell(getcell(vo.getY0(),fontChinese2));
        aTable.addCell(getcell(vo.getZ0(),fontChinese2));

        Cell cell19 = new Cell(new Phrase("变化值", fontChinese3));
        cell19.setVerticalAlignment(Element.ALIGN_CENTER);
        cell19.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell19.setBorderColor(new Color(0, 0, 0));

        aTable.addCell(cell19);

        aTable.addCell(getcell(changeNum(vo.getXx(),vo.getT12Xx()),isPass(vo.getXx(),vo.getT12Xx(),vo.getL12Xx()) ));
        aTable.addCell(getcell(changeNum(vo.getXy(),vo.getT12Xy()),isPass(vo.getXy(),vo.getT12Xy(),vo.getL12Xy()) ));
        aTable.addCell(getcell(changeNum(vo.getXz(),vo.getT12Xz()),isPass(vo.getXz(),vo.getT12Xz(),vo.getL12Xz()) ));

        aTable.addCell(getcell(changeNum(vo.getYx(),vo.getT12Yx()),isPass(vo.getYx(),vo.getT12Yx(),vo.getL12Yx()) ));
        aTable.addCell(getcell(changeNum(vo.getYy(),vo.getT12Yy()),isPass(vo.getYy(),vo.getT12Yy(),vo.getL12Yy()) ));
        aTable.addCell(getcell(changeNum(vo.getYz(),vo.getT12Yz()),isPass(vo.getYz(),vo.getT12Yz(),vo.getL12Yz()) ));

        aTable.addCell(getcell(changeNum(vo.getZx(),vo.getT12Zx()),isPass(vo.getZx(),vo.getT12Zx(),vo.getL12Zx()) ));
        aTable.addCell(getcell(changeNum(vo.getZy(),vo.getT12Zy()),isPass(vo.getZy(),vo.getT12Zy(),vo.getL12Zy()) ));
        aTable.addCell(getcell(changeNum(vo.getZz(),vo.getT12Xz()),isPass(vo.getZz(),vo.getT12Zz(),vo.getL12Zz()) ));

        aTable.addCell(getcell(changeNum(vo.getX0(),vo.getT12X()),isPass(vo.getX0(),vo.getT12X(),vo.getL12X()) ));
        aTable.addCell(getcell(changeNum(vo.getY0(),vo.getT12Y()),isPass(vo.getY0(),vo.getT12Y(),vo.getL12Y()) ));
        aTable.addCell(getcell(changeNum(vo.getZ0(),vo.getT12Z()),isPass(vo.getZ0(),vo.getT12Z(),vo.getL12Z()) ));

    }


        document.add(aTable);
        document.add(new Paragraph("\n"));
    }

    public String changeNum(String numStr,String numStr2){
        if(!StringUtils.isEmpty(numStr)&&!StringUtils.isEmpty(numStr2)){
            return String.valueOf(Float.valueOf(numStr)-Float.valueOf(numStr2));
        }else{
            return "";
        }
    }
    public boolean isPass(String numStr, String numStr2,String limit){
        if(!StringUtils.isEmpty(numStr)&&!StringUtils.isEmpty(numStr2)&&!StringUtils.isEmpty(limit)){
            return  Float.valueOf(numStr)-Float.valueOf(numStr2)>Float.valueOf(limit);
        }else{
            return false;
        }
    }

    Cell cell14;
    public Cell getcell(String name, Font fontChinese) throws BadElementException {
        cell14= new Cell(new Phrase(name, fontChinese));
        cell14.setVerticalAlignment(Element.ALIGN_CENTER);
        cell14.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell14.setBorderColor(new Color(0, 0, 0));
        return cell14;
    }


    /**
     *  动态改变word表格中不满足条件字体颜色
     * @param name
     * @param b
     * @return
     * @throws BadElementException
     */
    Font font;
    public Cell getcell(String name ,boolean b) throws BadElementException {
         font=new Font(bfChinese, 10, Font.NORMAL);
        if(b){
            font.setColor(Color.RED);
        }
        cell14= new Cell(new Phrase(name, font));
        cell14.setVerticalAlignment(Element.ALIGN_CENTER);
        cell14.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell14.setBorderColor(new Color(0, 0, 0));
        return cell14;
    }


}