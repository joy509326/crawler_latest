package com.example.selenium.bean;
import com.example.selenium.service.EmailService;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.openqa.selenium.chrome.ChromeOptions;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import joinery.DataFrame;
import net.lingala.zip4j.ZipFile;
import net.lingala.zip4j.model.ZipParameters;
import net.lingala.zip4j.model.enums.EncryptionMethod;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.TimeUnit;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@Component
public class Crawler {
    @Autowired
    private EmailService emailService;
    private static final Logger logger = LoggerFactory.getLogger(Crawler.class);

    @Value("${file.save.path}")
    private String filePath;

    @Scheduled(cron = "0 01 09,15 ? * MON-FRI")
    public void run() throws Exception {
        System.out.println("START RUN!");
        int maxAttempt = 2; // 最大重試次數
        boolean errorEmailFlag = false;
        for (int attempt = 1; attempt <= maxAttempt; attempt++) {
            try {
                Calendar Today = Calendar.getInstance();
                System.out.println("START :　today is : " + Today.get(Calendar.DAY_OF_WEEK));
                if (Today.get(Calendar.DAY_OF_WEEK) == Calendar.SATURDAY || Today.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY) { //判斷是否為周末->不執行
                    logger.info("Today is weekend");
                } else {
                    //時間區間設定
                    Date startDate;
                    Date startDateRemain;
                    Date endDateRemain;
                    if (Today.get(Calendar.DAY_OF_WEEK) == Calendar.MONDAY) {
                        startDate = getTargetDate(-3, 9);
                        startDateRemain = getTargetDate(-3, 0);
                        endDateRemain = getTargetDate(-3, 9);
                    } else {
                        //星期一捕抓三天內的新聞
                        startDate = getTargetDate(-1, 9);
                        startDateRemain = getTargetDate(-1, 0);
                        endDateRemain = getTargetDate(-1, 9);
                    }
                    logger.info("Sheet1 抓取資料起始時間:" + startDate);
                    logger.info("Sheet1 抓取資料結束時間:" + getTargetDate(0, 9));
                    logger.info("Sheet2 抓取資料起始時間:" + startDateRemain);
                    logger.info("Sheet2 抓取資料結束時間:" + endDateRemain);

                    System.setProperty("webdriver.chrome.driver", "C:\\Users\\a5614\\Desktop\\Selenium2\\crawler\\chrome-windows\\chromedriver.exe");
                    // local的chrome driver
//                     System.setProperty("webdriver.chrome.driver", "./crawler/chrome-windows/chromedriver.exe");
                    // AWS-windows的chrome driver
                    // System.setProperty("webdriver.chrome.driver", "C:\\Users\\Administrator\\chrome-windows\\chromedriver.exe");
                    // AWS-ubuntu的chrome driver
//                    System.setProperty("webdriver.chrome.driver","/home/ubuntu/SCSB_Crawler/chrome-ubuntu/chromedriver");
                    WebDriver driver = getWebDriver();
                    TimeUnit.SECONDS.sleep(2);
                    List<String> urlEncoder = new ArrayList<>();
                    List<String> keyWordsList = new ArrayList<>(Arrays.asList("洗錢", "資恐", "制裁", "違反證券交易法", "違反證交法", "逃漏稅", "虛擬貨幣", "貪污", "詐欺", "詐騙", "詐貸", "毒品", "非法吸金", "違法吸金"));
                    DataFrame<Object> df = new DataFrame<>("關鍵字", "標題", "新聞連結", "發布時間");
                    DataFrame<Object> dfRemain = new DataFrame<>("關鍵字", "標題", "新聞連結", "發布時間");
                    DataFrame<Object> dfOriginal;
                    DataFrame<Object> dfOriginalRemain;
                    DataFrame<Object> dfKeyword = new DataFrame<>();
                    String urlTemplate = "https://udn.com/search/word/2/";
                    String tempTime;
                    Date tempDate;
                    int countNews;
                    int countRemainNews;

                    for (int i = 0; i < keyWordsList.size(); i++) {
                        countNews = 0;
                        countRemainNews = 0;
                        urlEncoder.add(urlTemplate + URLEncoder.encode(keyWordsList.get(i), StandardCharsets.UTF_8));
                        driver.get(urlEncoder.get(i));
                        JavascriptExecutor js = (JavascriptExecutor) driver;
                        js.executeScript("window.scrollBy(0,document.body.scrollHeight)", "");
                        TimeUnit.SECONDS.sleep(3);
                        js.executeScript("window.scrollBy(0,document.body.scrollHeight)", "");
                        TimeUnit.SECONDS.sleep(3);
                        js.executeScript("window.scrollBy(0,document.body.scrollHeight)", "");
                        TimeUnit.SECONDS.sleep(3);
                        js.executeScript("window.scrollBy(0,document.body.scrollHeight)", "");
                        TimeUnit.SECONDS.sleep(3);
                        js.executeScript("window.scrollBy(0,document.body.scrollHeight)", "");
                        TimeUnit.SECONDS.sleep(3);
                        js.executeScript("window.scrollBy(0,document.body.scrollHeight)", "");
                        TimeUnit.SECONDS.sleep(3);
                        js.executeScript("window.scrollBy(0,document.body.scrollHeight)", "");
                        TimeUnit.SECONDS.sleep(3);

                        //限制在左邊區域
                        WebElement leftPanel = driver.findElement(By.cssSelector(".context-box__content.story-list__holder.story-list__holder--full"));
                        List<WebElement> ElementsList = leftPanel.findElements(By.cssSelector(".story-list__text"));
                        for (WebElement Elements : ElementsList) {
                            tempTime = Elements.findElement(By.cssSelector(".story-list__text time")).getText();
                            tempDate = new SimpleDateFormat("yyyy-MM-dd HH:mm").parse(tempTime);
                            //在時間區間內的資料才寫入
                            //24, 72小時新聞
                            if (tempDate.after(startDate) && tempDate.before(getTargetDate(0, 9))) {
                                df.append(Arrays.asList(keyWordsList.get(i), Elements.findElement(By.cssSelector(".story-list__text h2 a")).getText(),
                                        Elements.findElement(By.cssSelector(".story-list__text h2 a")).getAttribute("href"), Elements.findElement(By.cssSelector(".story-list__text time")).getText()));
                                countNews++;
                            }
                            //保留新聞
                            if (tempDate.after(startDateRemain) && tempDate.before(endDateRemain)) {
                                dfRemain.append(Arrays.asList(keyWordsList.get(i), Elements.findElement(By.cssSelector(".story-list__text h2 a")).getText(),
                                        Elements.findElement(By.cssSelector(".story-list__text h2 a")).getAttribute("href"), Elements.findElement(By.cssSelector(".story-list__text time")).getText()));
                                countRemainNews++;
                            }
                        }
                        if (countNews != 0) {
                            dfKeyword.append(Arrays.asList(keyWordsList.get(i), countNews, "非保留新聞"));
                        }
                        if (countRemainNews != 0) {
                            dfKeyword.append(Arrays.asList(keyWordsList.get(i), countRemainNews, "保留新聞"));
                        }
                    }
                    driver.quit();

                    logger.info("=====Unreserved News=====");
                    dfOriginal = getContent(df);
                    logger.info("=====Reserved News=====");
                    dfOriginalRemain = getContent(dfRemain);

                    df = df.join(dfOriginal);
                    dfRemain = dfRemain.join(dfOriginalRemain);

                    SimpleDateFormat DateFormat = new SimpleDateFormat("yyyy-MM-dd");
                    String fileDate = DateFormat.format(getTargetDate(-1, 0));

                    DataFrame<Object> countMinus;
                    countMinus = saveContent(df, dfRemain, fileDate);

                    int newCount;
                    for (int i = 0; i < countMinus.length() - 1; i++) {
                        for (int z = 0; z < dfKeyword.length(); z++) {
                            if (countMinus.get(i, 0) == dfKeyword.get(z, 0) && countMinus.get(i, 2) == dfKeyword.get(z, 2)) {
                                newCount = (Integer) dfKeyword.get(z, 1) - (Integer) countMinus.get(i, 1);
                                dfKeyword.set(z, 1, newCount);
                            }
                        }
                    }
                    saveCount(dfKeyword, fileDate);

                    // 將檔案壓縮成Zip
                    ZipParameters zipParameters = new ZipParameters();
                    zipParameters.setEncryptFiles(false);
                    zipParameters.setEncryptionMethod(EncryptionMethod.AES);

                    List<File> filesToAdd = Arrays.asList(
                            new File(filePath + fileDate + "news.xlsx"),
                            new File(filePath + fileDate + "keyword_count.xlsx")
                    );

                    //ZipFile zipFile = new ZipFile(filePath + fileDate + "新聞.zip", "1234".toCharArray());
                    ZipFile zipFile = new ZipFile(filePath + fileDate + "news.zip");
                    zipFile.addFiles(filesToAdd, zipParameters);
                    zipFile.close();

                    // 將檔案刪除
                    Path pathNews = Paths.get(filePath + fileDate + "news.xlsx");
                    Path pathCount = Paths.get(filePath + fileDate + "keyword_count.xlsx");
                    try {
                        Files.delete(pathNews);
                        Files.delete(pathCount);
                        logger.info("Files are successfully deleted");
                    } catch (IOException e) {
                        logger.warn("Files deletion failed", e);
                    }
                    //                    String[] to = new String[]{"joy509326@gmail.com","yshuang0924@gmail.com","stanley38288@notes.scsb.com.tw"};
//                    String[] to = new String[] {"joy509326@gmail.com","chipmunklee@gmail.com"};
                    String[] to = new String[] {"joy509326@gmail.com"};
                    //                    String[] to = new String[] {"marscheng@notes.scsb.com.tw", "hlc@notes.scsb.com.tw", "riverwei@notes.scsb.com.tw",
                    //                            "jun@notes.scsb.com.tw", "ianscsb@notes.scsb.com.tw", "jan@notes.scsb.com.tw",
                    //                            "stanley38288@notes.scsb.com.tw", "leyway@notes.scsb.com.tw", "samliu861868@notes.scsb.com.tw",
                    //                            "joy509326@gmail.com", "scu04116102@gmail.com", "chipmunklee@gmail.com", "pig02010821@gmail.com", "a0955002221@gmail.com",
                    //                            "jean.ye1998@gmail.com", "yshuang0924@gmail.com", "bb.bibigrace@gmail.com"};
                    String subject = fileDate + "新聞(自動寄送)(Ubuntu測試)";
                    String text = "<html><body>" +
                            "<p style='font-size:16px; color: black'>您好，</p><p style='font-size:16px; color: black'>附件為" +
                            DateFormat.format(startDate) + "至" + DateFormat.format(getTargetDate(0, 9)) +
                            "新聞整理，供您參考。</p><p style='font-size: 16px'>" +
                            " <strong style='text-decoration: underline; color: #00008B'>備註: 下載dat檔後，點選右鍵 > 點選7-zip > 點選解壓縮至此，即可出現excel檔案</strong>" +
                            " </p><p style='font-size:16px'><br><br><br></p>" +
                            "<p style='font-size:16px; color: black'>感謝撥冗閱讀</p><p style='font-size: 16px; color: black'>" +
                            "若有問題請洽資訊研發處南區資訊開發中心黃育盛 ，連絡電話：(06)2648111 #5130" +
                            "</p>" + "</body></html>";

                    String attachmentPath = filePath + fileDate + "news.zip";
                    emailService.sendEmail(to, subject, text, attachmentPath);
                    Path pathZip = Paths.get(filePath + fileDate + "news.zip");
                    try {
                        Files.delete(pathZip);
                        logger.info("Zip file is successfully deleted");
                    } catch (IOException e) {
                        logger.warn("Zip file deletion failed", e);
                    }
                }
                break; // 如果一切成功，跳出迴圈
            } catch (RuntimeException e) {
                String errorMessage = getStackTrace(e);
                logger.warn(errorMessage);
                if (attempt < maxAttempt) {
                    logger.info("Retry...(No." + attempt + "/" + (maxAttempt - 1) + ")");
                    TimeUnit.MINUTES.sleep(5);
                    logger.info("After waiting for five minutes, execute again");
                } else {
                    if (!errorEmailFlag) {
                        String[] to = new String[] {"joy509326@gmail.com"};
                        String subject = "法尊處負面新聞爬蟲程式錯誤訊息";
                        String text = "<html><body>" +
                                "<p style='font-size:16px; color: black'>您好，</p>" +
                                "<p style='font-size:16px; color: black'>法尊處負面新聞爬蟲程式錯誤!</p>" +
                                "<p style='font-size:16px; color: black'>勞煩您檢查並處理該問題。</p>" +
                                "<p style='font-size:16px;font-weight:bold; color: black'>錯誤訊息 : </p>" +
                                "<p style='font-size:16px; color: black'>" + errorMessage + "</p>" +
                                "<p style='font-size:16px; color: black'>謝謝</p>" +
                                "</body></html>";
                        emailService.sendEmail(to, subject, text);
                    } else {
                        logger.warn("Maximum retry attempts reached, sending error email");
                    }
                    break;
                }
            }
        }
    }

    public Date getTargetDate(int day, int hour) {
        Calendar NowDate = Calendar.getInstance();
        NowDate.add(Calendar.DAY_OF_MONTH, +day);
        NowDate.set(Calendar.HOUR_OF_DAY, +hour);
        NowDate.set(Calendar.MINUTE, 0);
        NowDate.set(Calendar.SECOND, 0);
        Date date = NowDate.getTime();
        return date;
    }

    public DataFrame<Object> getContent(DataFrame<Object> dataFrame) throws IOException {
        List<Object> urlColumn = dataFrame.col("新聞連結");
        DataFrame<Object> dfContent = new DataFrame<>("內文");
        for (int index = 0; index < urlColumn.size(); index++) {
            URL urlConnect = new URL(urlColumn.get(index).toString());
            HttpURLConnection connection = (HttpURLConnection) urlConnect.openConnection();
            connection.setRequestMethod("GET");
            connection.connect();

            int code = connection.getResponseCode();
            if (code == 200) {
                logger.info(" No. " + (index + 1) + " round / " + urlColumn.size() + " round " + urlColumn.get(index).toString() + " - HTTP GET succeed");
            } else {
                logger.info("HTTP GET failed ：" + code);
            }

            try {
                Document document = Jsoup.connect(urlColumn.get(index).toString()).get();
                // 使用選擇器找到所有 <div class="article-content__paragraph"> 中的 <p> 標籤
                Elements paragraphDivs = document.select("div.article-content__paragraph p");
                List<String> content = new ArrayList<>();
                // 所有 <p> 標籤的內容
                for (Element paragraphDiv : paragraphDivs) {
                    String text = paragraphDiv.text().trim();
                    if (!text.isEmpty()) {
                        content.add(text);
                    }
                }
                String article = String.join("\n", content);
                dfContent.append(List.of(article));
            } catch (Exception e) {
                logger.warn("No." + (index + 1) + "unable to retrieve the content!");
            }
        }
        return dfContent;
    }

    public DataFrame<Object> saveContent(DataFrame<Object> dataFrame, DataFrame<Object> dataFrameRemain, String fileDate) throws Exception {
        Calendar Today = Calendar.getInstance();
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet;

        // 標題字型
        Font titleFont = workbook.createFont();
        titleFont.setFontHeightInPoints((short) 16);
        titleFont.setBold(true);

        // 內文字型
        Font contentFont = workbook.createFont();
        contentFont.setFontHeightInPoints((short) 13);
        contentFont.setBold(true);

        // 關鍵字字型
        Font keywordFont = workbook.createFont();
        keywordFont.setFontHeightInPoints((short) 15);
        keywordFont.setBold(true);

        // 自訂標題格式
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setFont(titleFont);
        titleStyle.setBorderTop(BorderStyle.THICK);
        titleStyle.setBorderBottom(BorderStyle.THICK);
        titleStyle.setBorderLeft(BorderStyle.THICK);
        titleStyle.setBorderRight(BorderStyle.THICK);
        titleStyle.setAlignment(HorizontalAlignment.CENTER);

        // 自定內文格式
        CellStyle contentStyle = workbook.createCellStyle();
        contentStyle.setFont(contentFont);
        contentStyle.setAlignment(HorizontalAlignment.LEFT);

        // 自訂關鍵字格式
        CellStyle keywordStyle = workbook.createCellStyle();
        keywordStyle.setFont(keywordFont);

        // 超連結字型+格式
        CreationHelper createHelper = workbook.getCreationHelper();
        CellStyle hlinkStyle = workbook.createCellStyle();
        Font hlinkFont = workbook.createFont();
        hlinkFont.setUnderline(Font.U_SINGLE);
        hlinkFont.setColor(IndexedColors.BLUE.getIndex());
        hlinkStyle.setFont(hlinkFont);

        System.out.println("TODAY IS : " + Today.get(Calendar.DAY_OF_WEEK));

        if (Today.get(Calendar.DAY_OF_WEEK) == Calendar.MONDAY) {
            sheet = workbook.createSheet("72小時新聞");
        } else {
            sheet = workbook.createSheet("24小時新聞");
        }
        Sheet sheet2 = workbook.createSheet("保留新聞");

        // 把df(24,72小時新聞)的內容傳到sheet1
        for (int rowIndex = 1; rowIndex < dataFrame.length() + 1; rowIndex++) {
            Row row = sheet.createRow(rowIndex);
            row.setHeightInPoints(30);
            for (int colIndex = 0; colIndex < 5; colIndex++) {
                Cell cell = row.createCell(colIndex);
                Object value = dataFrame.get(rowIndex - 1, colIndex);
                if (colIndex < 2) {
                    cell.setCellStyle(contentStyle);
                }
                cell.setCellValue(value.toString());
                if (colIndex == 2) {
                    XSSFHyperlink link = (XSSFHyperlink) createHelper.createHyperlink(HyperlinkType.URL);
                    link.setAddress(value.toString());
                    cell.setHyperlink(link);
                    cell.setCellStyle(hlinkStyle);
                }
            }
        }

        // 把dfRemain(保留新聞)的內容傳到sheet2
        for (int rowIndex = 1; rowIndex < dataFrameRemain.length() + 1; rowIndex++) {
            Row row = sheet2.createRow(rowIndex);
            row.setHeightInPoints(30);
            for (int colIndex = 0; colIndex < 5; colIndex++) {
                Cell cell = row.createCell(colIndex);
                Object value = dataFrameRemain.get(rowIndex - 1, colIndex);
                cell.setCellValue(value.toString());
                if (colIndex < 2) {
                    cell.setCellStyle(contentStyle);
                }
                cell.setCellValue(value.toString());
                if (colIndex == 2) {
                    XSSFHyperlink link = (XSSFHyperlink) createHelper.createHyperlink(HyperlinkType.URL);
                    link.setAddress(value.toString());
                    cell.setHyperlink(link);
                    cell.setCellStyle(hlinkStyle);
                }
            }
        }

        // 添加每個sheet的標題
        String[] title = {"關鍵字", "標題", "新聞連結", "發布時間", "內文"};
        Row row = sheet.createRow(0);
        Row row2 = sheet2.createRow(0);
        Cell cell;
        Cell cell2;
        for (int i = 0; i < title.length; i++) {
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(titleStyle);
        }
        for (int i = 0; i < title.length; i++) {
            cell2 = row2.createCell(i);
            cell2.setCellValue(title[i]);
            cell2.setCellStyle(titleStyle);
        }

        for (int i = 0; i <= title.length; i++) {
            sheet.setColumnWidth(i + 1, 14000);
            sheet2.setColumnWidth(i + 1, 14000);
        }
        sheet.setColumnWidth(0, 8000);
        sheet2.setColumnWidth(0, 8000);

        // 處理sheet1
        // sheet 若是新聞網開頭的新聞or禁止爬蟲名單就刪除那個row
        List<String> count = dealDuplicate(sheet);
        DataFrame<Object> countSum = new DataFrame<>();

        // 計算各關鍵字重複以及非udn新聞網出現的次數
        Map<String, Integer> countFreq = new HashMap<>();
        count.forEach(x -> countFreq.put(x, countFreq.getOrDefault(x, 0) + 1));
        countFreq.forEach((key, value) -> countSum.append(Arrays.asList(key, value, "非保留新聞")));

        // 處理sheet2
        // sheet2 若是新聞網開頭的新聞or禁止爬蟲名單就刪除那個row
        List<String> count2 = dealDuplicate(sheet2);

        // 計算各關鍵字重複以及非udn新聞網出現的次數
        Map<String, Integer> countFreq2 = new HashMap<>();
        count2.forEach(x -> countFreq2.put(x, countFreq2.getOrDefault(x, 0) + 1));
        countFreq2.forEach((key, value) -> countSum.append(Arrays.asList(key, value, "保留新聞")));

        try {
            FileOutputStream fileOutputStream = new FileOutputStream(filePath + fileDate + "news.xlsx");
            workbook.write(fileOutputStream);
            workbook.close();
            fileOutputStream.close();
        } catch (Exception e) {
            logger.warn("wrong path(news)", e);
        }
        return countSum;
    }

    public List<String> dealDuplicate(Sheet sheet) {
        List<String> count = new ArrayList<>();
        for (int i = sheet.getLastRowNum(); i > 1; i--) {
            if (!sheet.getRow(i).getCell(2).getStringCellValue().startsWith("https://udn.com/news/story") &&
                    !sheet.getRow(i).getCell(2).getStringCellValue().contains("/news/story/7240/1365555")) {
                if (!sheet.getRow(i).getCell(2).getStringCellValue().startsWith("https://vip.udn.com/vip/story") && !sheet.getRow(i).getCell(2).getStringCellValue().startsWith("https://sdgs")) {
                    count.add(sheet.getRow(i).getCell(0).getStringCellValue());
                    sheet.shiftRows(i + 1, sheet.getLastRowNum() + 2, -1);
                }
            }
        }

        for (int i = sheet.getLastRowNum(); i > 1; i--) {
            if (sheet.getRow(i).getCell(2).getStringCellValue().contains("vip")) {
                sheet.shiftRows(i + 1, sheet.getLastRowNum() + 2, -1);
            }
        }

        List<String> newsTitle = new ArrayList<>();
        for (int i = 1; i < sheet.getLastRowNum(); i++) {
            newsTitle.add(sheet.getRow(i).getCell(1).getStringCellValue());
        }

        Map<String, List<Integer>> indexMap = new HashMap<>();
        for (int i = 0; i < newsTitle.size(); i++) {
            String element = newsTitle.get(i);
            element = element.replaceAll("\\s", ""); // 把標題字串裡面空白的地方去掉，在進行比對
            if (indexMap.containsKey(element)) {
                indexMap.get(element).add(i);
            } else {
                List<Integer> indexList = new ArrayList<>();
                indexList.add(i);
                indexMap.put(element, indexList);
            }
        }

        List<List<Integer>> repeatIndex = new ArrayList<>();
        // 印出重複元素和索引
        // 取的Map中所有 key 和 value
        for (Map.Entry<String, List<Integer>> entry : indexMap.entrySet()) {
            List<Integer> indices = entry.getValue();
            if (indices.size() > 1) {
                repeatIndex.add(indices);
            }
        }

        List<Integer> sortValue = new ArrayList<>();
        // 取得重複新聞除了第一次出現外的所有index
        for (List<Integer> innerList : repeatIndex) {
            for (int i = 1; i < innerList.size(); i++) {
                sortValue.add(innerList.get(i));
                sortValue.sort(Collections.reverseOrder());
            }
        }

        // 判斷index如果一樣，就把第二次出現的那個row刪除
        for (Integer integer : sortValue) {
            for (int j = sheet.getLastRowNum(); j > 1; j--) {
                if (integer + 1 == j) {
                    count.add(sheet.getRow(j).getCell(0).getStringCellValue());
                    sheet.shiftRows(j + 1, sheet.getLastRowNum() + 2, -1);
                }
            }
        }
        return count;
    }

    // 將計數匯出成Xlsx
    public void saveCount(DataFrame<Object> dataFrame, String fileDate) throws Exception {
        try {
            // 標題
            String[] title = new String[] {"關鍵字", "計數", "工作表"};
            // 關鍵字總筆數
            int keywordLength = Integer.parseInt(dataFrame.count().get(0, 0).toString());
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet(fileDate + "關鍵字計數");
            Row row;
            //關鍵字計數內容
            for (int rowIndex = 0; rowIndex <= keywordLength; rowIndex++) {
                //建立新的一列
                row = sheet.createRow(rowIndex);
                if (rowIndex == 0) {
                    //標題
                    for (int i = 0; i < title.length; i++) {
                        row.createCell(i).setCellValue(title[i]);
                    }
                } else {
                    //內容
                    row = sheet.createRow(rowIndex);
                    for (int i = 0; i < title.length; i++) {
                        row.createCell(i).setCellValue(dataFrame.get(rowIndex - 1, i).toString());
                    }
                }
            }

            // 如果那個cell是0的話，就刪掉那個row
            for (int i = sheet.getLastRowNum(); i > 1; i--) {
                if (sheet.getRow(i).getCell(1).getStringCellValue().equals("0")) {
                    sheet.shiftRows(i + 1, sheet.getLastRowNum() + 2, -1);
                }
            }

            //設定檔案寫出指定位置
            FileOutputStream fileOutputStream = new FileOutputStream(filePath + fileDate + "keyword_count.xlsx");
            workbook.write(fileOutputStream);
            workbook.close();
            fileOutputStream.close();
        } catch (Exception e) {
            logger.warn("wrong path(count)", e);
        }
    }

    public String getStackTrace(Throwable throwable) {
        StringWriter stringWriter = new StringWriter();
        PrintWriter printWriter = new PrintWriter(stringWriter);
        throwable.printStackTrace(printWriter);
        return stringWriter.toString();
    }

    /* 初始 WebDriver & 設定 WebDriver 參數 */
    public WebDriver getWebDriver() {
        ChromeOptions chromeOption = new ChromeOptions();
        chromeOption.addArguments("--headless");
        /* 停用 sandbox 模式 */
        chromeOption.addArguments("--no-sandbox");
        /* 視窗最大化 */
        chromeOption.addArguments("--start-maximized");
        /* 禁止擴充功能 */
        chromeOption.addArguments("--disable-extensions");
        /* 不使用內存作為暫存區 */
        chromeOption.addArguments("--disable-dev-shm-usage");
        /* 固定執行視窗 Size */
        chromeOption.addArguments("--window-size=1920,1080");
        /* 禁止彈跳式訊息 */
        chromeOption.addArguments("--disable-popup-blocking");
        /* 不自動載入圖片檔案 */
        chromeOption.addArguments("blink-settings=imagesEnabled=false");
        /* 初始 chromeDriver */
        return new ChromeDriver(chromeOption);
    }
}
