package net.topikachu.excel;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
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
    public void gen() {

        int colCount = 5;
        int rowCount = 10;
        try (InputStream templateStream = ExcelRunner.class.getClassLoader().getResourceAsStream("template.xlsx")) {
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
            String sheetName = sheet.getSheetName();
            for (int i = 1; i <= rowCount + 1; i++) {
                XSSFRow row = sheet.createRow(i);
                XSSFCell rowName = row.createCell(0);
                rowName.setCellValue("row" + i);
                for (int j = 1; j <= colCount; j++) {
                    row.createCell(j).setCellValue(rand.nextInt(rowCount) + 1);
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


            FileOutputStream fos = new FileOutputStream("output.xlsx");
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
        gen();
    }
}
