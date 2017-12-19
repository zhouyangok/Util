package word.dynamicTwo;


import Excel.FrontSpaceSeries;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;


/**
 * Created with IntelliJ IDEA.
 *
 * @Author medo_zy
 * @Desciption:
 * @Date 2017-12-18 16:29
 */
public class WordUtils {
    public static void main(String[] args) throws Exception {
        WordUtils t = new WordUtils();
        XWPFDocument document = new XWPFDocument();
        t.megerTableCell(document);
        t.saveDocument(document, "D:\\generateWord\\" + System.currentTimeMillis() + ".docx");
    }

    public void megerTableCell(XWPFDocument document) {
        XWPFTable table1 = document.createTable();
        setTableWidth(table1, "8000");
        fillTable(table1);
        mergeCellsVertically(table1, 1, 9, 10);
        mergeCellsVertically(table1, 1, 1, 4);
        mergeCellsVertically(table1, 4, 2, 4);
        mergeCellsHorizontal(table1, 0, 0, 14);
        mergeCellsHorizontal(table1, 2, 3, 14);
        mergeCellsHorizontal(table1, 3, 3, 14);
        mergeCellsHorizontal(table1, 4, 0, 14);
        mergeCellsHorizontal(table1, 5, 3, 14);
        mergeCellsHorizontal(table1, 6, 3, 5);
        mergeCellsHorizontal(table1, 6, 6, 8);
        mergeCellsHorizontal(table1, 6, 9, 11);
        mergeCellsHorizontal(table1, 6, 12, 14);
    }

    public void fillTable(XWPFTable ComTable) {

//自动分配列宽
        CTTblWidth comTableWidth = ComTable.getCTTbl().addNewTblPr().addNewTblW();
        comTableWidth.setType(STTblWidth.DXA);
        comTableWidth.setW(BigInteger.valueOf(9072));

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
        frontSpaceSeries.setXx("3.232");
        frontSpaceSeries.setXy("2.342");
        frontSpaceSeries.setXz("3.564");
        frontSpaceSeries.setYx("3.334");
        frontSpaceSeries.setYy("3.452");
        frontSpaceSeries.setYz("3.452");
        frontSpaceSeries.setZx("3.452");
        frontSpaceSeries.setZy("3.3423");
        frontSpaceSeries.setZz("2.432");
        frontSpaceSeries.setX0("3.3423");
        frontSpaceSeries.setY0("3.4564");
        frontSpaceSeries.setZ0("7.23423");
        list.add(frontSpaceSeries);
        list.add(frontSpaceSeries);
        list.add(frontSpaceSeries);
        list.add(frontSpaceSeries);
        list.add(frontSpaceSeries);
//        DynamicFinal.generate(list);
//        for (int rowIndex = 0; rowIndex < table.getNumberOfRows(); rowIndex++) {
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
            comTableRowX.getCell(0).setColor("FFEC8B");
            comTableRowX.getCell(2).setText("设计值");
            comTableRowX.getCell(3).setText(list.get(i).getXx());
            comTableRowX.getCell(4).setText(list.get(i).getXy());
            comTableRowX.getCell(5).setText(list.get(i).getXz());
            comTableRowX.getCell(6).setText(list.get(i).getYx());
            comTableRowX.getCell(7).setText(list.get(i).getYy());
            comTableRowX.getCell(8).setText(list.get(i).getYz());
            comTableRowX.getCell(9).setText(list.get(i).getZx());
            comTableRowX.getCell(10).setText(list.get(i).getZy());
            comTableRowX.getCell(11).setText(list.get(i).getZz());
            comTableRowX.getCell(12).setText(list.get(i).getX0());
            comTableRowX.getCell(13).setText(list.get(i).getY0());
            comTableRowX.getCell(14).setText(list.get(i).getZ0());
            XWPFTableRow comTableRowY = ComTable.createRow();
            comTableRowY.getCell(1).setText("振前测量");
            comTableRowY.getCell(2).setText("测量值");
            comTableRowY.getCell(3).setText(list.get(i).getXx());
            comTableRowY.getCell(4).setText(list.get(i).getXy());
            comTableRowY.getCell(5).setText(list.get(i).getXz());
            comTableRowY.getCell(6).setText(list.get(i).getYx());
            comTableRowY.getCell(7).setText(list.get(i).getYy());
            comTableRowY.getCell(8).setText(list.get(i).getYz());
            comTableRowY.getCell(9).setText(list.get(i).getZx());
            comTableRowY.getCell(10).setText(list.get(i).getZy());
            comTableRowY.getCell(11).setText(list.get(i).getZz());
            comTableRowY.getCell(12).setText(list.get(i).getX0());
            comTableRowY.getCell(13).setText(list.get(i).getY0());
            comTableRowY.getCell(14).setText(list.get(i).getZ0());
            XWPFTableRow comTableRowZ = ComTable.createRow();
            comTableRowZ.getCell(2).setText("变化值");
            comTableRowZ.getCell(3).setText(list.get(i).getXx());
            comTableRowZ.getCell(4).setText(list.get(i).getXy());
            comTableRowZ.getCell(5).setText(list.get(i).getXz());
            comTableRowZ.getCell(6).setText(list.get(i).getYx());
            comTableRowZ.getCell(7).setText(list.get(i).getYy());
            comTableRowZ.getCell(8).setText(list.get(i).getYz());
            comTableRowZ.getCell(9).setText(list.get(i).getZx());
            comTableRowZ.getCell(10).setText(list.get(i).getZy());
            comTableRowZ.getCell(11).setText(list.get(i).getZz());
            comTableRowZ.getCell(12).setText(list.get(i).getX0());
            comTableRowZ.getCell(13).setText(list.get(i).getY0());
            comTableRowZ.getCell(14).setText(list.get(i).getZ0());

        }
           /* XWPFTableRow row = table.getRow(rowIndex);
            row.setHeight(380);
            for (int colIndex = 0; colIndex < row.getTableCells().size(); colIndex++) {
                XWPFTableCell cell = row.getCell(colIndex);
                if(rowIndex%2==0){
                    setCellText(cell, " cell " + rowIndex + colIndex + " ", "D4DBED", 1000);
                }else{
                    setCellText(cell, " cell " + rowIndex + colIndex + " ", "AEDE72", 1000);
                }
            }*/
    }

    public void setCellText(XWPFTableCell cell, String text, String bgcolor, int width) {
        CTTc cttc = cell.getCTTc();
        CTTcPr cellPr = cttc.addNewTcPr();
        cellPr.addNewTcW().setW(BigInteger.valueOf(width));
        //cell.setColor(bgcolor);  
        CTTcPr ctPr = cttc.addNewTcPr();
        CTShd ctshd = ctPr.addNewShd();
        ctshd.setFill(bgcolor);
        ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
        cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
        cell.setText(text);
    }


    /**
     * @Description: 跨列合并
     */
    public static void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
            if (cellIndex == fromCell) {
                // The first merged cell is set with RESTART merge value  
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE  
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }


    public static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            if (rowIndex == fromRow) {
                // The first merged cell is set with RESTART merge value  
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE  
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    public void setTableWidth(XWPFTable table, String width) {
        CTTbl ttbl = table.getCTTbl();
        CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
        CTTblWidth tblWidth = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
        CTJc cTJc = tblPr.addNewJc();
        cTJc.setVal(STJc.Enum.forString("center"));
        tblWidth.setW(new BigInteger(width));
        tblWidth.setType(STTblWidth.DXA);
    }

    public void saveDocument(XWPFDocument document, String savePath)
            throws Exception {
        FileOutputStream fos = new FileOutputStream(savePath);
        document.write(fos);
        fos.close();
    }
} 

