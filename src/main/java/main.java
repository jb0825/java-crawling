import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import javax.lang.model.element.ExecutableElement;
import java.io.*;
import java.net.URL;
import java.sql.Array;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class main {
    public static void main(String[] args) throws IOException {
        crawling();
    }

    public static List<List<String>> readExcel(String fileName) throws IOException {
        FileInputStream in = new FileInputStream(fileName);
        XSSFWorkbook workbook = new XSSFWorkbook(in);
        XSSFSheet sheet = workbook.getSheetAt(0);

        int rowNum = sheet.getPhysicalNumberOfRows();
        System.out.println("ROWS : " + rowNum);

        List<List<String>> datas = new ArrayList<>();

        for (int i = 0; i < rowNum; i++) {
            XSSFRow row = sheet.getRow(i);
            if (row != null) {
                int cellNum = row.getPhysicalNumberOfCells();

                List<String> data = new ArrayList<>();

                for (int j = 0; j <= cellNum; j++) {
                    XSSFCell cell = row.getCell(j);

                    String value = "";
                    if (cell == null) {
                        continue;
                    } else {
                        switch (cell.getCellType()) {
                            case XSSFCell.CELL_TYPE_FORMULA:
                                value = cell.getCellFormula();
                                break;
                            case XSSFCell.CELL_TYPE_NUMERIC:
                                value = Math.round(cell.getNumericCellValue()) + "";
                                break;
                            case XSSFCell.CELL_TYPE_STRING:
                                value = cell.getStringCellValue() + "";
                                break;
                            case XSSFCell.CELL_TYPE_BLANK | XSSFCell.CELL_TYPE_ERROR:
                                value = "";
                                break;
                        }
                    }
                    data.add(value);
                }
                datas.add(data);
            }
        }

        //for (int i = 0; i < datas.size(); i++) System.out.println(datas.get(i));
        in.close();
        return datas;
    }

    static void writeExcel(List<List<String>> datas) throws IOException {
        File file = new File("test.xlsx");
        FileOutputStream out = new FileOutputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();

        for (int i = 0; i < datas.size(); i++) {
            List<String> data = datas.get(i);
            XSSFRow rows = sheet.createRow(i);

            for (int j = 0; j < data.size(); j++) {
                XSSFCell cell = rows.createCell(j);
                cell.setCellValue(data.get(j));
            }
        }

        workbook.write(out);
        out.close();
    }

    public static void crawling() throws IOException {
        String defaultUrl = "https://bizno.net/";

        List<List<String>> datas = readExcel("C:\\Users\\주빈\\Desktop\\file.xlsx");
        List<List<String>> writeData = new ArrayList<>();

        for (List<String> data : datas) {
            String companyName = data.get(1).replaceAll("\\(주\\)", "").replaceAll("주식회사", "").replaceAll("㈜", "").trim();
            System.out.println(companyName);
            String url = defaultUrl + "?query=" + companyName;
            Document doc = Jsoup.connect(url).get();
            Elements elements = doc.getElementsByClass("details");

            if (elements.size() > 0) {
                Element element = elements.get(0);

                if (elements.size() > 1) {
                    for (Element e : elements) {
                        String address = element.getElementsByTag("p").last().text();
                        if (address.contains("서울")) {
                            element = e;
                            break;
                        }
                    }
                }

                Elements titles = element.getElementsByClass("titles");

                for (Element title : titles) {
                    Element aTag = element.getElementsByTag("a").first();
                    if (aTag.text().contains(companyName)) {
                        doc = Jsoup.connect(defaultUrl + aTag.attr("href")).get();
                        String text = doc.getElementsByClass("table_guide01").text();
                        String businessNo = text.split("사업자등록번호 ")[1].split("법인등록번호")[0];
                        String homepage = text.split("홈페이지 ")[1].split("IR")[0];
                        String item = text.split("주요제품 ")[1].split("연매출액")[0];

                        data.set(5, businessNo);
                        data.set(6, homepage);
                        data.set(7, item);
                        System.out.println(data);
                        writeData.add(data);
                        break;
                    }
                }
            }
        }

        writeExcel(datas);
    }

}
