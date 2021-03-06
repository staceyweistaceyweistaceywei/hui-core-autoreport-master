package com.hui.core.report.app;


import java.io.FileOutputStream;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTCatAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLegend;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTScaling;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.STAxPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STCrosses;
import org.openxmlformats.schemas.drawingml.x2006.chart.STLegendPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STOrientation;
import org.openxmlformats.schemas.drawingml.x2006.chart.STTickLblPos;

public class LineChart {

        public static void main(String[] args) throws Exception {
            Workbook wb = new XSSFWorkbook();
            Sheet sheet = wb.createSheet("Sheet1");

            Row row;
            Cell cell;

            row = sheet.createRow(0);
            row.createCell(0);
            row.createCell(1).setCellValue("??????");
            row.createCell(2).setCellValue("??????");

            for (int r = 1; r < 7; r++) {
                row = sheet.createRow(r);
                cell = row.createCell(0);
                cell.setCellValue("C" + r);
                cell = row.createCell(1);
                cell.setCellValue(10 + new java.util.Random().nextDouble());
                cell = row.createCell(2);
                cell.setCellValue(12 + new java.util.Random().nextDouble()*10d);
            }

            XSSFDrawing drawing = (XSSFDrawing)sheet.createDrawingPatriarch();
            ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 15);

            XSSFChart chart = drawing.createChart(anchor);

            CTChart ctChart = ((XSSFChart)chart).getCTChart();
            CTPlotArea ctPlotArea = ctChart.getPlotArea();

            //the first line chart
            CTLineChart ctBarChart = ctPlotArea.addNewLineChart();
            CTBoolean ctBoolean = ctBarChart.addNewVaryColors();
            ctBoolean.setVal(true);

            //the first line series
            CTLineSer ctBarSer = ctBarChart.addNewSer();
            CTSerTx ctSerTx = ctBarSer.addNewTx();
            CTStrRef ctStrRef = ctSerTx.addNewStrRef();

            String titleRef = new CellReference(sheet.getSheetName(), 0, 1, true, true).formatAsString();
//            ctStrRef.setF("Sheet1!$B$1");
            ctStrRef.setF(titleRef);
            ctBarSer.addNewIdx().setVal(0);
            CTAxDataSource cttAxDataSource = ctBarSer.addNewCat();
            ctStrRef = cttAxDataSource.addNewStrRef();

            ctStrRef.setF("Sheet1!$A$2:$A$7");
            CTNumDataSource ctNumDataSource = ctBarSer.addNewVal();
            CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
            ctNumRef.setF("Sheet1!$B$2:$B$7");

            //   at least the border lines in Libreoffice Calc ;-)
            ctBarSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[] {0,0,0});

            //telling the first LineChart that it has axes and giving them Ids
            ctBarChart.addNewAxId().setVal(123456); //cat axis 1 (lines)
            ctBarChart.addNewAxId().setVal(123457); //val axis 1 (left)

            // ==========   the second line chart ===============
            CTLineChart ctLineChart = ctPlotArea.addNewLineChart();
            ctBoolean = ctLineChart.addNewVaryColors();
            ctBoolean.setVal(true);


            //the second line series
            CTLineSer ctLineSer = ctLineChart.addNewSer();
            ctSerTx = ctLineSer.addNewTx();
            ctStrRef = ctSerTx.addNewStrRef();

            String titleRef1 = new CellReference(sheet.getSheetName(), 0, 2, true, true).formatAsString();
            ctStrRef.setF(titleRef1);
            ctLineSer.addNewIdx().setVal(1);
            cttAxDataSource = ctLineSer.addNewCat();
            ctStrRef = cttAxDataSource.addNewStrRef();


            ctStrRef.setF("Sheet1!$A$2:$A$7");
            ctNumDataSource = ctLineSer.addNewVal();
            ctNumRef = ctNumDataSource.addNewNumRef();

            ctNumRef.setF("Sheet1!$C$2:$C$7");

            //at least the border lines in Libreoffice Calc ;-)
            ctLineSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[] {0,0,0});

            //telling the second LineChart that it has axes and giving them Ids
            ctLineChart.addNewAxId().setVal(123458); //cat axis 2 (lines)
            ctLineChart.addNewAxId().setVal(123459); //val axis 2 (right)


            //cat axis 1 (lines)
            CTCatAx ctCatAx = ctPlotArea.addNewCatAx();
            ctCatAx.addNewAxId().setVal(123456); //id of the cat axis
            CTScaling ctScaling = ctCatAx.addNewScaling();
            ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
            ctCatAx.addNewDelete().setVal(false);
            ctCatAx.addNewAxPos().setVal(STAxPos.B);
            ctCatAx.addNewCrossAx().setVal(123457); //id of the val axis
            ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

            //val axis 1 (left)
            CTValAx ctValAx = ctPlotArea.addNewValAx();
            ctValAx.addNewAxId().setVal(123457); //id of the val axis
            ctScaling = ctValAx.addNewScaling();
            ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
            ctValAx.addNewDelete().setVal(false);
            ctValAx.addNewAxPos().setVal(STAxPos.L);
            ctValAx.addNewCrossAx().setVal(123456); //id of the cat axis
            ctValAx.addNewCrosses().setVal(STCrosses.AUTO_ZERO); //this val axis crosses the cat axis at zero
            ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

            //cat axis 2 (lines)
            ctCatAx = ctPlotArea.addNewCatAx();
            ctCatAx.addNewAxId().setVal(123458); //id of the cat axis
            ctScaling = ctCatAx.addNewScaling();
            ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
            ctCatAx.addNewDelete().setVal(true); //this cat axis is deleted
            ctCatAx.addNewAxPos().setVal(STAxPos.B);
            ctCatAx.addNewCrossAx().setVal(123459); //id of the val axis
            ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

            //val axis 2 (right)
            ctValAx = ctPlotArea.addNewValAx();
            ctValAx.addNewAxId().setVal(123459); //id of the val axis
            ctScaling = ctValAx.addNewScaling();
            ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
            ctValAx.addNewDelete().setVal(false);
            ctValAx.addNewAxPos().setVal(STAxPos.R);
            ctValAx.addNewCrossAx().setVal(123458); //id of the cat axis
            ctValAx.addNewCrosses().setVal(STCrosses.MAX); //this val axis crosses the cat axis at max value
            ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

            //legend
            CTLegend ctLegend = ctChart.addNewLegend();
            ctLegend.addNewLegendPos().setVal(STLegendPos.B);
            ctLegend.addNewOverlay().setVal(false);

            FileOutputStream fileOut = new FileOutputStream("E:\\program\\workspace\\hui-core-autoreport\\hui-core-autoreport-master\\src\\main\\resources\\CombinedLineChart.xlsx");
            wb.write(fileOut);
            fileOut.close();
            wb.close();
        }
}
