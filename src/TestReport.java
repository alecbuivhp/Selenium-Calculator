import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.markuputils.MarkupHelper;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;
import data.DataLoader;
import org.openqa.selenium.By;
import org.openqa.selenium.Capabilities;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.Assert;
import org.testng.ITestResult;
import org.testng.annotations.*;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class TestReport {
    // Excel path
    final static String excelFilePath = "src/data/testcase.xlsx";
    List<WebDriver> driverList;
    // builds a new report using the html template
    ExtentSparkReporter reporter;

    ExtentReports extent;

    // helps to generate the logs in test report.
    ExtentTest test;

    @Parameters({"OS", "browser"})
    @BeforeTest
    public void startReport(String OS, String browser) {
        // initialize the Reporter
        reporter = new ExtentSparkReporter("./extent-reports/extent-report.html");

        extent = new ExtentReports();
        extent.attachReporter(reporter);

        extent.setSystemInfo("OS", OS);
        extent.setSystemInfo("Browser", browser);

        // configuration items to change the look and feel
        // add content, manage tests etc
        reporter.config().setDocumentTitle("Calculator Test Report");
        reporter.config().setReportName("Test Report");
        reporter.config().setTheme(Theme.STANDARD);
        reporter.config().setTimeStampFormat("EEEE, MMMM dd, yyyy, hh:mm a '('zzz')'");

        System.setProperty("webdriver.chrome.driver", "lib/webdriver/chromedriver");
        System.setProperty("webdriver.edge.driver", "lib/webdriver/msedgedriver");
        System.setProperty("webdriver.gecko.driver", "lib/webdriver/geckodriver");

        driverList = new ArrayList<>();
        driverList.add(new ChromeDriver());
        driverList.add(new FirefoxDriver());
        driverList.add(new EdgeDriver());

        for (WebDriver driver : driverList) {
            driver.get("https://testsheepnz.github.io/BasicCalculator.html");
        }
    }

    @BeforeTest
    @DataProvider(name = "data-provider")
    public Object[][] dpMethod() throws IOException {
        return DataLoader.readExcel(excelFilePath);
    }

    @Test(dataProvider = "data-provider")
    public void executeTest(int id, String description, Object num1, Object num2, int operator, boolean isInteger, Object expectedResult, int driverIndex, int build) {
        WebDriver driver = driverList.get(driverIndex);
        Capabilities cap = ((RemoteWebDriver) driver).getCapabilities();
        test = extent.createTest(cap.getBrowserName() + "(Build index " + build + "): Test case #" + id, description);

        driver.navigate().refresh();
        test.log(Status.INFO, "Refreshed page");

        Select dropDownBuild = new Select(driver.findElement(By.id("selectBuild")));
        dropDownBuild.selectByValue(Integer.toString(build));
        test.log(Status.INFO, "Selected build (index): " + build);

        driver.findElement(By.id("number1Field")).sendKeys(num1.toString());
        test.log(Status.INFO, "Sent number 1 value: " + num1);

        driver.findElement(By.id("number2Field")).sendKeys(num2.toString());
        test.log(Status.INFO, "Sent number 2 value: " + num2);

        Select dropDownOp = new Select(driver.findElement(By.id("selectOperationDropdown")));
        dropDownOp.selectByValue(Integer.toString(operator));
        test.log(Status.INFO, "Selected operation (index): " + operator);

        if (isInteger && operator != 4) {
            driver.findElement(By.id("integerSelect")).click();
            test.log(Status.INFO, "Checked 'Convert to integer'");
        }
        driver.findElement(By.id("calculateButton")).click();
        test.log(Status.INFO, "Clicked calculate button");

        String ans = driver.findElement(By.id("numberAnswerField")).getAttribute("value");
        if (ans.equals("")){
            ans = null;
        }
        if (expectedResult != null) {
            expectedResult = expectedResult.toString();
        }
        Assert.assertEquals(ans, expectedResult);
    }

    @AfterMethod
    public void getResult(ITestResult result) {
        if (result.getStatus() == ITestResult.FAILURE) {
            test.log(Status.FAIL, MarkupHelper.createLabel(result.getName() + " FAIL ", ExtentColor.RED));
            test.fail(result.getThrowable());
        } else if (result.getStatus() == ITestResult.SUCCESS) {
            test.log(Status.PASS, MarkupHelper.createLabel(result.getName() + " PASSED ", ExtentColor.GREEN));
        } else {
            test.log(Status.SKIP, MarkupHelper.createLabel(result.getName() + " SKIPPED ", ExtentColor.ORANGE));
            test.skip(result.getThrowable());
        }
    }

    @AfterClass
    public void tearDown() {
        // to write or update test information to reporter
        extent.flush();
    }

    @AfterClass
    public void stopDriver() {
        for (WebDriver driver : driverList) {
            driver.quit();
        }
    }
}
