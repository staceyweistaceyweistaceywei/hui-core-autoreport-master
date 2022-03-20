package com.hui.core.report.app;

import com.spire.xls.Workbook;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

public class ChartToImage {

    public static void main(String[] args) throws IOException {
        //加载Excel文档
        Workbook workbook = new Workbook();
        workbook.loadFromFile("E:\\program\\workspace\\hui-core-autoreport\\hui-core-autoreport-master\\src\\main\\resources\\CombinedLineChart.xlsx");
        //将Excel文档第一个工作表中的第一个图表保存为图片
        BufferedImage image= workbook.saveChartAsImage(workbook.getWorksheets().get(0), 0);
        ImageIO.write(image,"png", new File("ChartToImage.png"));
    }

}
