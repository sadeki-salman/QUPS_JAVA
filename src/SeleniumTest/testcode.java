package SeleniumTest;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class testcode {

	public static void main(String[] args) {
		// Set the path to the ChromeDriver executable
        System.setProperty("webdriver.chrome.driver", ".\\driver\\chromedriver.exe");

     // Get the current day
        DayOfWeek currentDay = LocalDate.now().getDayOfWeek();
        System.out.println("Today is " + currentDay);

        // Load the Excel workbook
        Workbook workbook;
        Sheet worksheet;
        try {
            FileInputStream fis = new FileInputStream(".\\file\\file.xlsx");
            workbook = new XSSFWorkbook(fis);
            worksheet = workbook.getSheet(currentDay.toString());
        } catch (Exception e) {
            e.printStackTrace();
            return;
        }

        // Find keywords list
        List<String> keywords = new ArrayList<>();
        Iterator<Row> rowIterator = worksheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(2);
            if (cell != null) {
                keywords.add(cell.getStringCellValue());
            }
        }
        System.out.println("Keywords are Collected Successfully: " + keywords);

        // Initialize ChromeDriver
        WebDriver driver = new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

        // Search for each keyword in Google and find the longest and shortest option
        int rowNumber = 2;
        for (String keyword : keywords) {
            driver.get("https://www.google.com/");
            WebElement searchBox = driver.findElement(By.name("q"));
            searchBox.sendKeys(keyword);
            try {
                Thread.sleep(1000); // Wait to load suggestions
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
            searchBox.sendKeys(Keys.ARROW_DOWN); // Move the cursor down to trigger every suggestion

            // Get the suggestion elements
            List<WebElement> suggestionElements = driver.findElements(By.cssSelector("li.sbct"));

            // Get the suggestions
            List<String> suggestionList = new ArrayList<>();
            for (WebElement suggestion : suggestionElements) {
                String suggestionText = suggestion.getText().split("\n")[0];
                if (!suggestionText.isEmpty()) {
                    suggestionList.add(suggestionText);
                }
            }
            System.out.println("Keyword: " + keyword);
            System.out.println("Suggestions are: " + suggestionList);

            String longestSuggestion = suggestionList.stream().max((s1, s2) -> s1.length() - s2.length()).orElse("");
            String shortestSuggestion = suggestionList.stream().min((s1, s2) -> s1.length() - s2.length()).orElse("");
            System.out.println("Longest Suggestion is '" + longestSuggestion + "'");
            System.out.println("Shortest Suggestion is '" + shortestSuggestion + "'");

            // Store the values in the worksheet
            Row row = worksheet.getRow(rowNumber);
            if (row == null) {
                row = worksheet.createRow(rowNumber);
            }
            Cell longestCell = row.createCell(3);
            Cell shortestCell = row.createCell(4);
            longestCell.setCellValue(longestSuggestion);
            shortestCell.setCellValue(shortestSuggestion);

            rowNumber++;
        }

        driver.quit();

        // Save changes to the worksheet
        try {
            FileOutputStream fos = new FileOutputStream(".\\file\\file.xlsx");
            workbook.write(fos);
            workbook.close();
            fos.close();
            System.out.println("Suggestions are Successfully Stored in the Worksheet");
        } catch (Exception e) {
            e.printStackTrace();
        }

	}

}
