import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class Main {

    public static ExecutorService executorService = Executors.newFixedThreadPool(CrawlingThread.THREAD_SIZE);

    public static void main(String... args) {
        start();
    }

    public static void start() {
        executorService.submit(() -> {
            for (int i = 0; i < CrawlingThread.THREAD_SIZE; i++) {
                CrawlingThread crawlingThread = new CrawlingThread();
                new Thread(crawlingThread).start();
            }
        });
    }

    public static void end() {
        Map<Integer, String> data = CrawlingThread.writeData;
        List<Integer> keySet = new ArrayList<>(data.keySet());
        List<String> listData = new ArrayList<>();

        keySet.sort(Integer::compareTo);
        for (Integer key : keySet) listData.add(data.get(key));

        new Excel().writeExcel(listData);
    }

    public static void noThreadCrawling() {
        try {
            String defaultUrl = "https://bizno.net/";

            List<List<String>> datas = new Excel().readExcel2("C:\\Users\\2021.05.03\\Desktop\\excel\\a.xlsx");
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

            new Excel().writeExcel2(datas);
        } catch (IOException e) {

        }
    }
}

class Excel {

    public Map<Integer, String> readExcel(String fileName) {
        System.out.println("[READ EXCEL FILE]");
        try {
            FileInputStream in = new FileInputStream(fileName);
            XSSFWorkbook workbook = new XSSFWorkbook(in);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Map<Integer, String> data = new HashMap<>();

            for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                XSSFRow row = sheet.getRow(i);
                if (row != null) {
                    XSSFCell cell = row.getCell(3);
                    String value = "";

                    switch (cell.getCellType()) {
                        case XSSFCell.CELL_TYPE_NUMERIC:
                            value = Math.round(cell.getNumericCellValue()) + "";
                            break;
                        case XSSFCell.CELL_TYPE_STRING:
                            value = cell.getStringCellValue();
                            break;
                    }
                    data.put(i, value);
                }
            }

            return data;
        } catch (IOException e) {
            System.out.println("[EXCEL READ ERROR]");
            e.printStackTrace();
            return null;
        }
    }

    public void writeExcel(List<String> data) {
        System.out.println("[WRITE EXCEL FILE]");

        data.forEach(System.out::println);

        try {
            File file = new File("test.xlsx");
            FileOutputStream out = new FileOutputStream(file);
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet();

            for (int i = 0; i < data.size(); i++) {
                XSSFRow rows = sheet.createRow(i);
                XSSFCell cell = rows.createCell(0);
                cell.setCellValue(data.get(i));
            }

            workbook.write(out);
        } catch (IOException e) {
            System.out.println("[EXCEL WRITE ERROR]");
            e.printStackTrace();
        }
    }

    // Cell 이 여러개인 엑셀파일 읽기
    public List<List<String>> readExcel2(String fileName) throws IOException {
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

    // Cell 이 여러개인 엑셀파일 쓰기
    static void writeExcel2(List<List<String>> datas) throws IOException {
        File file = new File("test2.xlsx");
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
}

class CrawlingThread implements Runnable {

    public static final int ROW_SIZE = 35742;
    public static final int THREAD_SIZE = 161;
    public static final int SIZE = ROW_SIZE / THREAD_SIZE;
    public static final String DEFAULT_URL = "https://bizno.net/";

    public static Map<Integer, String> data;
    public static Map<Integer, String> writeData = new HashMap<>();
    public static int threadCount = 0;
    public static int idx = 0;

    static {
        data = new Excel().readExcel("C:\\Users\\2021.05.03\\Desktop\\excel\\b.xlsx");
    }

    int mapIdx = 0;

    public CrawlingThread() { idx += SIZE; mapIdx = idx; }

    @Override
    public void run() {
        System.out.println("*** THREAD " + mapIdx / SIZE + " ***");
        String naverUrl = "https://search.naver.com/search.naver?where=nexearch&sm=top_hty&fbm=1&ie=utf8&query=";

        for (int i = mapIdx - SIZE; i < mapIdx; i++) {
            String homepage = "";
            try {
                if (i > ROW_SIZE) return;

                String companyName = data.get(i);
                companyName = companyName.replaceAll("\\(주\\)", "").replaceAll("주식회사", "").replaceAll("㈜", "").replaceAll(" ", "").trim();
                homepage = naverUrl + companyName;
                System.out.println("[" + i + " : " + companyName + "]");
                String url = DEFAULT_URL + "?query=" + companyName;

                Document doc = Jsoup.connect(url).get();
                Elements elements = doc.getElementsByClass("details");

                if (elements.size() > 0) {
                    Element element = elements.get(0);

                    if (elements.size() > 1)
                        for (Element e : elements)
                            if (element.getElementsByTag("p").last().text().contains("서울")) {
                                element = e;
                                break;
                            }

                    Element aTag = element.getElementsByTag("a").first();
                    String aCompanyName = aTag.text();
                    String aHref = aTag.attr("href");

                    if (aCompanyName.contains(companyName) || companyName.contains(aCompanyName) || aCompanyName.equals(companyName)) {
                        doc = Jsoup.connect(DEFAULT_URL + aHref).get();
                        homepage = doc.getElementsByClass("table_guide01").text().split("홈페이지 ")[1].split("IR")[0];

                        if (homepage.equals("- ") || homepage.contains("오픈마켓") ||
                            homepage.contains("스마트스토어") || homepage.contains("스토어팜") ||
                            homepage.contains("네이버") || homepage.contains("옥션")
                        )
                            homepage = naverUrl + companyName;
                    }
                }
            }
            catch(IOException e) { e.printStackTrace(); }
            finally { writeData.put(i, homepage); }
        }

        if (++threadCount == THREAD_SIZE) {
            System.out.println("*** CRAWLING END ***");

            Main.executorService.shutdown();
            Main.end();
        }
    }
}
