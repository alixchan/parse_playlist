import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;


public class Parsing {

    public static final String URL_PLAYLISTS = "https://www.youtube.com/c/CompscicenterRu/playlists";
    public static final String COUNT_PLAYLIST = "//*[@id=\"overlays\"]" +
            "/ytd-thumbnail-overlay-side-panel-renderer/yt-formatted-string";
    public static final String TITLE_PLAYLIST = "//*[@id=\"video-title\"]";
    public static final String FILENAME = "res.xls";

    public static void main(String[] args) throws InterruptedException, IOException {
        System.setProperty("webdriver.chrome.driver", "./src/main/resources/chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();

        driver.get(URL_PLAYLISTS);

        new WebDriverWait(driver, 60)
                .until(dr -> ((JavascriptExecutor) dr).executeScript("return document.readyState").equals(
                        "complete"));
        var last_height = ((JavascriptExecutor) driver).executeScript("return document.documentElement" +
                ".scrollHeight");
        while (true) {
            ((JavascriptExecutor) driver).executeScript("window.scrollTo(0, document.documentElement" +
                    ".scrollHeight);");
            Thread.sleep(4000);
            var new_height = ((JavascriptExecutor) driver).executeScript("return document.documentElement" +
                    ".scrollHeight");
            if (new_height.equals(last_height)) {
                break;
            }
            last_height = new_height;
        }

        List<WebElement> count = driver.findElements(By.xpath(COUNT_PLAYLIST));
        List<WebElement> title = driver.findElements(By.xpath(TITLE_PLAYLIST));

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("data");

        for (int i = 0; i < count.size(); i++) {
            Row row = sheet.createRow(i);
            Cell cell = row.createCell(0);
            cell.setCellValue(count.get(i).getText());

            cell = row.createCell(1);
            cell.setCellValue(title.get(i).getText());
        }
        try (FileOutputStream outputStream = new FileOutputStream(FILENAME)) {
            workbook.write(outputStream);
        }
        System.out.println(".xlsx written successfully");

        driver.quit();
    }
}
