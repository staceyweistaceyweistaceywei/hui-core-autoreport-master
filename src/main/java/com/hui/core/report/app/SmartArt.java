package com.hui.core.report.app;

import com.spire.presentation.FileFormat;
import com.spire.presentation.ISlide;
import com.spire.presentation.Presentation;
import com.spire.presentation.diagrams.*;

public class SmartArt {

    public static void main(String[] args) throws Exception{
        //创建PPT文档，获取一张幻灯片（创建的空白PPT文档，默认包含一张幻灯片）
        Presentation ppt = new Presentation();
        ISlide slide = ppt.getSlides().get(0);

        //创建SmartArt图形1
        ISmartArt smartArt1 = slide.getShapes().appendSmartArt(50,50,200,200, SmartArtLayoutType.BASIC_CYCLE);//在幻灯片指定位置添加指定大小和布局类型的SmartArt图形
        smartArt1.setColorStyle(SmartArtColorType.COLORFUL_ACCENT_COLORS_4_TO_5);//设置SmartArt图形颜色类型
        smartArt1.setStyle(SmartArtStyleType.INTENCE_EFFECT);//设置SmartArt图形样式
        ISmartArtNode smartArtNode1 = smartArt1.getNodes().get(0);//获取节点
        smartArtNode1.getTextFrame().setText("设计");//添加内容
        smartArt1.getNodes().get(1).getTextFrame().setText("模仿");
        smartArt1.getNodes().get(2).getTextFrame().setText("学习");
        smartArt1.getNodes().get(3).getTextFrame().setText("实践");
        smartArt1.getNodes().get(4).getTextFrame().setText("创新");


        //创建SmartArt图形2，自定义节点内容
        ISmartArt smartArt2 = slide.getShapes().appendSmartArt(400,200,200,200,SmartArtLayoutType.BASIC_RADIAL);
        smartArt2.setColorStyle(SmartArtColorType.DARK_2_OUTLINE);
        smartArt2.setStyle(SmartArtStyleType.MODERATE_EFFECT);
        //删除默认的节点（SmartArt中的图形）
        for (Object a : smartArt2.getNodes())
        {
            smartArt2.getNodes().removeNode((ISmartArtNode) a);
        }
        //添加一个母节点
        ISmartArtNode node2 = smartArt2.getNodes().addNode();
        //在母节点下添加三个子节点
        ISmartArtNode node2_1 = node2.getChildNodes().addNode();
        ISmartArtNode node2_2 = node2.getChildNodes().addNode();
        ISmartArtNode node2_3 = node2.getChildNodes().addNode();
        //在节点上设置文字及文字大小
        node2.getTextFrame().setText("设备");
        node2.getTextFrame().getTextRange().setFontHeight(14f);
        node2_1.getTextFrame().setText("机械");
        node2_1.getTextFrame().getTextRange().setFontHeight(12f);
        node2_2.getTextFrame().setText("电气");
        node2_2.getTextFrame().getTextRange().setFontHeight(12f);
        node2_3.getTextFrame().setText("自动化");
        node2_3.getTextFrame().getTextRange().setFontHeight(12f);

        //保存文档
        ppt.saveToFile("AddSmartArt.pptx", FileFormat.PPTX_2013);
        ppt.dispose();
    }

}
