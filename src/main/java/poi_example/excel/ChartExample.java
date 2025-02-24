package poi_example.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Excelファイルにグラフを作成するサンプル。
 */
public class ChartExample {
    public static void main(String[] args) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Chart Data");

            // データの作成
            Object[][] data = {
                {"月", "売上"},
                {1, 1000},
                {2, 2000},
                {3, 1500},
                {4, 3000}
            };

            // データをシートに書き込み
            int rowCount = 0;
            for (Object[] rowData : data) {
                Row row = sheet.createRow(rowCount++);
                int columnCount = 0;
                for (Object field : rowData) {
                    Cell cell = row.createCell(columnCount++);
                    if (field instanceof String) {
                        cell.setCellValue((String) field);
                    } else if (field instanceof Integer) {
                        cell.setCellValue((Integer) field);
                    }
                }
            }

            // グラフの作成
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 5, 10, 20);

            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("月別売上");
            chart.setTitleOverlay(false);

            // データソースの設定
            XDDFDataSource<Double> xs = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                new CellRangeAddress(1, 4, 0, 0));
            XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                new CellRangeAddress(1, 4, 1, 1));

            // グラフ種類の設定（縦棒グラフ）
            XDDFChartData chartData = chart.createData(ChartTypes.BAR,
                chart.createCategoryAxis(AxisPosition.BOTTOM),
                chart.createValueAxis(AxisPosition.LEFT));

            XDDFChartData.Series series = chartData.addSeries(xs, ys);
            series.setTitle("売上", null);
            chart.plot(chartData);

            // ファイルに保存
            new File("output").mkdirs();
            try (FileOutputStream fileOut = new FileOutputStream("output/chart_example.xlsx")) {
                workbook.write(fileOut);
            }
        }
        System.out.println("グラフを含むExcelファイルを作成しました。");
    }
}