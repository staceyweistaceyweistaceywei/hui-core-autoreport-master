package com.hui.core.report.app;

import com.spire.presentation.FileFormat;
import com.spire.presentation.Presentation;
import com.spire.presentation.charts.ChartType;
import com.spire.presentation.charts.IChart;
import com.spire.presentation.drawing.FillFormatType;

import java.awt.geom.Rectangle2D;

public class CombinationChart {
    public static void main(String[] args) throws Exception {

        //创建PowerPoint文档
        Presentation presentation = new Presentation();

        //添加一个柱状图
        Rectangle2D.Double rect = new  Rectangle2D.Double(50, 100, 550, 300);
        IChart chart = presentation.getSlides().get(0).getShapes().appendChart(ChartType.COLUMN_CLUSTERED, rect);

        //设置图表名称
        chart.getChartTitle().getTextProperties().setText("销售额VS单价");
        chart.getChartTitle().getTextProperties().isCentered(true);
        chart.getChartTitle().setHeight(30);
        chart.hasTitle(true);

        //写入图表数据
        chart.getChartData().get(0,0).setText("月份");
        chart.getChartData().get(0,1).setText("单价");
        chart.getChartData().get(0,2).setText("销售额");
        chart.getChartData().get(1,0).setText("一月");
        chart.getChartData().get(1,1).setNumberValue(120);
        chart.getChartData().get(1,2).setNumberValue(5600);
        chart.getChartData().get(2,0).setText("二月");
        chart.getChartData().get(2,1).setNumberValue(100);
        chart.getChartData().get(2,2).setNumberValue(7300);
        chart.getChartData().get(3,0).setText("三月");
        chart.getChartData().get(3,1).setNumberValue(80);
        chart.getChartData().get(3,2).setNumberValue(10200);
        chart.getChartData().get(4,0).setText("四月");
        chart.getChartData().get(4,1).setNumberValue(120);
        chart.getChartData().get(4,2).setNumberValue(5900);
        chart.getChartData().get(5,0).setText("五月");
        chart.getChartData().get(5,1).setNumberValue(90);
        chart.getChartData().get(5,2).setNumberValue(9500);
        chart.getChartData().get(6,0).setText("六月");
        chart.getChartData().get(6,1).setNumberValue(110);
        chart.getChartData().get(6,2).setNumberValue(7200);

        //设置系列标签数据来源
        chart.getSeries().setSeriesLabel(chart.getChartData().get("B1", "C1"));

        //设置分类标签数据来源
        chart.getCategories().setCategoryLabels(chart.getChartData().get("A2", "A7"));

        //设置系列的数据来源
        chart.getSeries().get(0).setValues(chart.getChartData().get("B2", "B7"));
        chart.getSeries().get(1).setValues(chart.getChartData().get("C2", "C7"));

        //将系列2的图表类型设置为折线图
        chart.getSeries().get(1).setType(ChartType.LINE_MARKERS);

        //将系列2绘制在次坐标轴
        chart.getSeries().get(1).setUseSecondAxis(true);

        //不显示次坐标轴的网格线
        chart.getSecondaryValueAxis().getMajorGridTextLines().setFillType(FillFormatType.NONE);

        //设置系列重叠
        chart.setOverLap(-50);

        //设置分类间距
        chart.setGapDepth(200);

        //保存文档
        presentation.saveToFile("E:\\program\\workspace\\hui-core-autoreport\\hui-core-autoreport-master\\src\\main\\resources\\CombinationChart-cn.pptx", FileFormat.PPTX_2010);
    }

}
