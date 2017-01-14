package net.topikachu.excel;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;
import org.springframework.boot.CommandLineRunner;
import org.springframework.stereotype.Component;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Random;

/**
 * Created by 禕 on 2017/1/14.
 */
@Component
public class ExcelRunner implements CommandLineRunner {
    public void genBarChart() {

        int colCount = 5;
        int rowCount = 10;
        try (InputStream templateStream = ExcelRunner.class.getClassLoader().getResourceAsStream("barChartTemplate.xlsx")) {
            XSSFWorkbook wb = new XSSFWorkbook(templateStream);
            XSSFSheet sheet = wb.getSheetAt(0);


            XSSFRow headers = sheet.createRow(0);

            for (int i = 1; i <= colCount; i++) {
                headers.createCell(i).setCellValue("header" + i);
            }
            CellRangeAddress headerRange = new CellRangeAddress(0, 0, 1, colCount);

            Random rand = new Random(System.currentTimeMillis());

            XSSFDrawing drawing = sheet.getDrawingPatriarch();

            CTBarChart barChart = drawing.getCharts().get(0).getCTChart().getPlotArea().getBarChartArray()[0];
            CTBarSer serTemplate = barChart.getSerArray()[0];
            int serLength = barChart.getSerArray().length;
            for (int i=serLength-1;i>=1;i--) {
                barChart.removeSer(i);
            }
            String sheetName = sheet.getSheetName();
            for (int i = 1; i <= rowCount ; i++) {
                XSSFRow row = sheet.createRow(i);
                XSSFCell rowName = row.createCell(0);
                rowName.setCellValue("row" + i);
                for (int j = 1; j <= colCount; j++) {
                    row.createCell(j).setCellValue(rand.nextInt(10) + 1);
                }
                CTBarSer ser;
                if (i != 1) {
                    barChart.addNewSer();
                }
                ser = (CTBarSer) serTemplate.copy();
                ser.getIdx().setVal(i - 1);
                ser.getOrder().setVal(i - 1);
                ser.getTx().getStrRef().setF(new CellRangeAddress(i, i, 0, 0).formatAsString(sheetName, true));
                ser.getCat().getStrRef().setF(headerRange.formatAsString(sheetName, true));
                ser.getVal().getNumRef().setF(new CellRangeAddress(i, i, 1, colCount).formatAsString(sheetName, true));
                if (ser.getSpPr().getSolidFill().isSetSchemeClr()) {
                    ser.getSpPr().getSolidFill().unsetSchemeClr();
                }
                barChart.setSerArray(i - 1, ser);
            }


            FileOutputStream fos = new FileOutputStream("barChartOutput.xlsx");
            wb.write(fos);

            fos.flush();
            fos.close();
            wb.close();

            System.out.println();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }


    }

    public void genPieChart() {

        int colCount = 5;
        try (InputStream templateStream = ExcelRunner.class.getClassLoader().getResourceAsStream("pieChartTemplate.xlsx")) {
            XSSFWorkbook wb = new XSSFWorkbook(templateStream);
            XSSFSheet sheet = wb.getSheetAt(0);


            XSSFRow headers = sheet.createRow(0);

            for (int i = 1; i <= colCount; i++) {
                headers.createCell(i).setCellValue("header" + i);
            }
            CellRangeAddress headerRange = new CellRangeAddress(0, 0, 1, colCount);

            Random rand = new Random(System.currentTimeMillis());

            XSSFDrawing drawing = sheet.getDrawingPatriarch();

            CTPieChart pieChart = drawing.getCharts().get(0).getCTChart().getPlotArea().getPieChartArray()[0];
            CTPieSer serTemplate = pieChart.getSerArray()[0];
            int serLength = pieChart.getSerArray().length;
            for (int i=serLength-1;i>=1;i--) {
                pieChart.removeSer(i);
            }
            String sheetName = sheet.getSheetName();
            for (int i = 1; i <= 1 ; i++) {
                XSSFRow row = sheet.createRow(i);
                XSSFCell rowName = row.createCell(0);
                rowName.setCellValue("row" + i);
                for (int j = 1; j <= colCount; j++) {
                    row.createCell(j).setCellValue(rand.nextInt(10) + 1);
                }
                CTPieSer ser;
                if (i != 1) {
                    pieChart.addNewSer();
                }
                ser = (CTPieSer) serTemplate.copy();

                ser.getTx().getStrRef().setF(new CellRangeAddress(i, i, 0, 0).formatAsString(sheetName, true));
                ser.getCat().getStrRef().setF(headerRange.formatAsString(sheetName, true));
                ser.getVal().getNumRef().setF(new CellRangeAddress(i, i, 1, colCount).formatAsString(sheetName, true));

                pieChart.setSerArray(i - 1, ser);
            }


            FileOutputStream fos = new FileOutputStream("pieChartOutput.xlsx");
            wb.write(fos);

            fos.flush();
            fos.close();
            wb.close();

            System.out.println();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }


    }

    @Override
    public void run(String... strings) throws Exception {
        genBarChart();
        genPieChart();
    }
}
