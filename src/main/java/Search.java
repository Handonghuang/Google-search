
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.springframework.http.*;
import org.springframework.web.client.RestTemplate;

import java.io.*;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Created by Administrator on 2018/8/9 0009.
 */
public class Search {

	//read excel file path
    private final static String SOURCE_FILE_PATH = "C:/Users/Administrator/Desktop/publist_ssci.xlsx";
    //sheet in excel
	private final static String SOURCE_SHEET_NAME = "Table 1";
	//result from google search page 1
    private final static String DEFAULT_PAGE_NUM = "1";
	//google search result
    private final static String DEFAULT_ROW_NUM = "20";
	//search start in row number in single sheet
    private final static Integer BEGIN_ROW = 4;
    //result save path
    private final static String GENERATE_FILE_PATH = "C:/Users/Administrator/Desktop/result.xlsx";
	//sheet name
    private final static String GENERATE_SHEET_NAME = "Sheet 1";

    public static void main(String[] args) {
		//read through the search name
        List<String> searchTitle = readFromExcel(SOURCE_FILE_PATH, SOURCE_SHEET_NAME);
		//restTemplate
        RestTemplate restTemplate = new RestTemplate();
        HttpHeaders headers = new HttpHeaders();
        headers.setAccept(Arrays.asList(MediaType.APPLICATION_JSON));
        headers.add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.99 Safari/537.36");
        HttpEntity<String> entity = new HttpEntity<>("parameters", headers);
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet(GENERATE_SHEET_NAME);
        String title = null;
        ResponseEntity<String> searchResult = null;
        Row row = null;
        Cell cell = null;
        Document document = null;
        Element searchDiv = null;
        Elements gDivs = null;
        Element h3 = null;
        Element a = null;
        for (int i = 0; i < searchTitle.size(); i++) {
            int j = 0;
            System.out.println(i);
            try {
				//control request speed to google otherwise may return http 503 error
                Thread.currentThread().sleep(1000L);// ms
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
//            try {
//                title = URLEncoder.encode(searchTitle.get(i),"utf-8");
//            } catch (UnsupportedEncodingException e) {
//                e.printStackTrace();
//            }
            title = searchTitle.get(i);
			//start search
            searchResult = restTemplate.exchange("https://www.google.com/search?q=" + title + "&start=" + DEFAULT_PAGE_NUM + "&num=" + DEFAULT_ROW_NUM, HttpMethod.GET, entity, String.class);
            row = sheet.createRow(i);
			//analysis the result from google
            document = Jsoup.parse(searchResult.getBody());
			//search div#search
            searchDiv = document.selectFirst("#search");
            if (searchDiv != null) {
				//check does it have div.g or not
                gDivs = searchDiv.select("div[class=g]");
                if (gDivs != null && gDivs.size() > 0) {
                    for (Element gDiv : gDivs) {
						//check does it have h3.r or not
                        h3 = gDiv.selectFirst("h3[class=r]");
                        if (h3 != null) {
							//check a
                            a = h3.selectFirst("a");
                            if (a != null) {
                                cell = row.createCell(j);
								//get the link from search result href and save into cell
                                cell.setCellValue(a.attr("href"));
                                j++;
                                if (j == 2) {
                                    break;
                                }
                            }
                        }
                    }
                }
            }
			//if there is no answer or not the one we need search with title with key words Journal
            if (j == 0) {
//                try {
//                    title = URLEncoder.encode(searchTitle.get(i)+" Journal","utf-8");
//                } catch (UnsupportedEncodingException e) {
//                    e.printStackTrace();
//                }
                title = searchTitle.get(i) + " Journal";
                searchResult = restTemplate.exchange("https://www.google.com/search?q=" + title + "&start=" + DEFAULT_PAGE_NUM + "&num=" + DEFAULT_ROW_NUM, HttpMethod.GET, entity, String.class);
                document = Jsoup.parse(searchResult.getBody());
                searchDiv = document.selectFirst("#search");
                if (searchDiv != null) {
                    gDivs = searchDiv.select("div[class=g]");
                    if (gDivs != null && gDivs.size() > 0) {
                        for (Element gDiv : gDivs) {
                            h3 = gDiv.selectFirst("h3[class=r]");
                            if (h3 != null) {
                                a = h3.selectFirst("a");
                                if (a != null) {
                                    cell = row.createCell(j);
                                    cell.setCellValue(a.attr("href"));
                                    j++;
                                    if (j == 2) {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        try {
			//save the result
            File generateFile = new File(GENERATE_FILE_PATH);
            if (generateFile.exists()) {
                generateFile.delete();
            }
            wb.write(new FileOutputStream(GENERATE_FILE_PATH));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

	//read through the excel and return a list of title
    private static List<String> readFromExcel(String filePath, String sheetName) {
        List<String> result = new ArrayList<>();
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(filePath);
            try {
                XSSFWorkbook wb = new XSSFWorkbook(fis);
                Sheet sheet = wb.getSheet(sheetName);
                Row row = null;
                for (int i = BEGIN_ROW; i < 104; i++) {//the numbers of the title need to find
                    row = sheet.getRow(i);
                    result.add(String.valueOf(row.getCell(0)));
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return result;
    }
}
