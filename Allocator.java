package allocator;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.regex.Pattern;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.Platform;

import com.cognizant.framework.selenium.*;

import rerun.src_rerun.*;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;
import com.aventstack.extentreports.reporter.KlovReporter;
import com.aventstack.extentreports.reporter.configuration.ChartLocation;
import com.aventstack.extentreports.reporter.configuration.Theme;
import com.cognizant.framework.ExcelDataAccessforxlsm;
import com.cognizant.framework.FrameworkException;
import com.cognizant.framework.FrameworkParameters;
import com.cognizant.framework.IterationOptions;
import com.cognizant.framework.Settings;
import com.cognizant.framework.TimeStamp;
import com.cognizant.framework.Util;


/**
 * Class to manage the batch execution of test scripts within the framework
 * 
 * @author Cognizant-
 */
public class Allocator {
	private FrameworkParameters frameworkParameters = FrameworkParameters.getInstance();
	private Properties properties;
	private Properties mobileProperties;
	private ResultSummaryManager resultSummaryManager = ResultSummaryManager.getInstance();
	private static final XmlReader Rerun = new XmlReader();
	private static final TimeStamp time = new TimeStamp();
	private static ExtentHtmlReporter htmlReporter;
	private static ExtentReports extentReport;
	private static ExtentTest extentTest;
	private static KlovReporter klovReporter = new KlovReporter();
	public static String env_currentTestcase;
	public static String env_ApplicationName;
	public static String env_Bu;
	public static String executionMode;
	public static String url;

	/**private static final 
	 * The entry point of the test batch execution <br>
	 * Exits with a value of 0 if the test passes and 1 if the test fails
	 * 
	 * @param args
	 *            Command line arguments to the Allocator (Not applicable)
	 */
	public static void main(String[] args) {
		deleteFile("ApplicationExecuted.txt");
		deleteFile("FailedApplicationName.txt");
		deleteFile("Run_2.txt");
		deleteFile("ApplicationFailed.txt");
		deleteFile("ApplicationPassed.txt");
		deleteFile("Summary.txt");
		deleteFile("Summary_2.txt");
		Allocator allocator = new Allocator();
		allocator.driveBatchExecution();
	}

	@SuppressWarnings("static-access")
	private void driveBatchExecution() {
		try{
			int testBatchStatus = runExecution() ;

			resultSummaryManager.wrapUp(false);

			if (Boolean.parseBoolean(properties.getProperty("GenerateKlov"))) {
				extentReport.attachReporter(klovReporter);
			}
			extentReport.flush();
			resultSummaryManager.launchResultSummary();
			Rerun.mainClass();
			time.reportPathWithTimeStamp =null;
			if(Boolean.parseBoolean(properties.getProperty("Rerun"))) {
				testBatchStatus = runExecution() ;
				resultSummaryManager.wrapUp(false);
				if (Boolean.parseBoolean(properties.getProperty("GenerateKlov"))) {
					extentReport.attachReporter(klovReporter);
				}
				extentReport.flush();
				resultSummaryManager.launchResultSummary();
				Rerun.FailuremainClass();
			}
			
				System.exit(testBatchStatus);
		}catch(Exception e){
			System.out.println(e.toString());
		}
	}
	public int runExecution() {
		resultSummaryManager.setRelativePath();
		properties = Settings.getInstance();
		mobileProperties = Settings.getMobilePropertiesInstance();
		String runConfiguration;
		if (System.getProperty("RunConfiguration") != null) {
			runConfiguration = System.getProperty("RunConfiguration");
		} else {
			runConfiguration = properties.getProperty("RunConfiguration");
		}
		resultSummaryManager.initializeTestBatch(runConfiguration);

		int nThreads = Integer.parseInt(properties.getProperty("NumberOfThreads"));
		resultSummaryManager.initializeSummaryReport(nThreads);

		resultSummaryManager.setupErrorLog();

		generateExtentReports();

		int testBatchStatus = executeTestBatch(nThreads);

		
		return testBatchStatus;

	}
	private int executeTestBatch(int nThreads) {
		List<SeleniumTestParameters> testInstancesToRun = getRunInfo(frameworkParameters.getRunConfiguration());
		ExecutorService parallelExecutor = Executors.newFixedThreadPool(nThreads);
		ParallelRunner testRunner = null;

		for (int currentTestInstance = 0; currentTestInstance < testInstancesToRun.size(); currentTestInstance++) {
			testRunner = new ParallelRunner(testInstancesToRun.get(currentTestInstance));
			parallelExecutor.execute(testRunner);

			if (frameworkParameters.getStopExecution()) {
				break;
			}
		}

		parallelExecutor.shutdown();
		while (!parallelExecutor.isTerminated()) {
			try {
				Thread.sleep(3000);
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
		}

		if (testRunner == null) {
			return 0; // All tests flagged as "No" in the Run Manager
		} else {
			return testRunner.getTestBatchStatus();
		}
	}

	private List<SeleniumTestParameters> getRunInfo(String sheetName) {

		ExcelDataAccessforxlsm runManagerAccess = new ExcelDataAccessforxlsm(
				frameworkParameters.getRelativePath() + Util.getFileSeparator() + "src" + Util.getFileSeparator()
						+ "test" + Util.getFileSeparator() + "resources",
				"Run Manager");
		runManagerAccess.setDatasheetName(sheetName);

		runManagerAccess.setDatasheetName(sheetName);
		List<SeleniumTestParameters> testInstancesToRun = new ArrayList<SeleniumTestParameters>();
		String[] keys = { "Execute", "TestScenario", "TestCase",
				"TestInstance", "Description", "IterationMode",
				"StartIteration", "EndIteration", "TestConfigurationID","ApplicationName","BU","WiniumAppPath" };
		List<Map<String, String>> values = runManagerAccess.getValues(keys);

		for (int currentTestInstance = 0; currentTestInstance < values.size(); currentTestInstance++) {

			Map<String, String> row = values.get(currentTestInstance);
			String executeFlag = row.get("Execute");

			if (executeFlag.equalsIgnoreCase("Yes")) {
				String currentScenario = row.get("TestScenario");
				String currentTestcase = row.get("TestCase");
				env_ApplicationName = row.get("ApplicationName");
				env_Bu =row.get("BU");
				env_currentTestcase = currentTestcase;
				SeleniumTestParameters testParameters = new SeleniumTestParameters(
						currentScenario, currentTestcase);
				testParameters.setBU(row.get("BU"));
				testParameters
						.setCurrentTestDescription(row.get("Description"));
				testParameters
				.setWiniumAppPath(row.get("WiniumAppPath"));
				testParameters.setApplicationName(row.get("ApplicationName"));
				testParameters.setCurrentTestInstance("Instance"
						+ row.get("TestInstance"));
				testParameters.setExtentReport(extentReport);
				testParameters.setExtentTest(extentTest);
				String iterationMode = row.get("IterationMode");
				if (!iterationMode.equals("")) {
					testParameters.setIterationMode(IterationOptions
							.valueOf(iterationMode));
				} else {
					testParameters
							.setIterationMode(IterationOptions.RUN_ALL_ITERATIONS);
				}

				String startIteration = row.get("StartIteration");
				if (!startIteration.equals("")) {
					testParameters.setStartIteration(Integer
							.parseInt(startIteration));
				}
				String endIteration = row.get("EndIteration");
				if (!endIteration.equals("")) {
					testParameters.setEndIteration(Integer
							.parseInt(endIteration));
				}
				String testConfig = row.get("TestConfigurationID");
				if (!"".equals(testConfig)) {
					getTestConfigValues(runManagerAccess, "TestConfigurations",
							testConfig, testParameters);
				}

				testInstancesToRun.add(testParameters);
				runManagerAccess.setDatasheetName(sheetName);
			}
		}
		return testInstancesToRun;
	}

	private void getTestConfigValues(ExcelDataAccessforxlsm runManagerAccess, String sheetName, String testConfigName,
			SeleniumTestParameters testParameters) {

		runManagerAccess.setDatasheetName(sheetName);
		int rowNum = runManagerAccess.getRowNum(testConfigName, 0, 1);

		String[] keys = { "TestConfigurationID", "ExecutionMode",
				"MobileToolName", "MobileExecutionPlatform", "MobileOSVersion",
				"DeviceName", "Browser", "BrowserVersion", "Platform",
				"SeeTestPort","OSapiName"};
		Map<String, String> values = runManagerAccess.getValuesForSpecificRow(
				keys, rowNum);
		
		executionMode = values.get("ExecutionMode");
		if (!"".equals(executionMode)) {
			testParameters.setExecutionMode(ExecutionMode
					.valueOf(executionMode));
		} else {
			testParameters.setExecutionMode(ExecutionMode.valueOf(properties
					.getProperty("DefaultExecutionMode")));
		}
		String toolName = values.get("MobileToolName");
		if (!"".equals(toolName)) {
			testParameters.setMobileToolName(ToolName.valueOf(toolName));
		} else {
			testParameters.setMobileToolName(ToolName
					.valueOf(mobileProperties
							.getProperty("DefaultMobileToolName")));
		}

		String executionPlatform = values.get("MobileExecutionPlatform");
		if (!"".equals(executionPlatform)) {
			testParameters.setMobileExecutionPlatform(MobileExecutionPlatform
					.valueOf(executionPlatform));
		} else {
			testParameters.setMobileExecutionPlatform(MobileExecutionPlatform
					.valueOf(mobileProperties
							.getProperty("DefaultMobileExecutionPlatform")));
		}

		String mobileOSVersion = values.get("MobileOSVersion");
		if (!"".equals(mobileOSVersion)) {
			testParameters.setmobileOSVersion(mobileOSVersion);
		}

		String deviceName = values.get("DeviceName");
		if (!"".equals(deviceName)) {
			testParameters.setDeviceName(deviceName);
		}

		String browser = values.get("Browser");
		if (!"".equals(browser)) {
			testParameters.setBrowser(Browser.valueOf(browser));
		} else {
			testParameters.setBrowser(Browser.valueOf(properties
					.getProperty("DefaultBrowser")));
		}

		String browserVersion = values.get("BrowserVersion");
		if (!"".equals(browserVersion)) {
			testParameters.setBrowserVersion(browserVersion);
		}

		String platform = values.get("Platform");
		if (!"".equals(platform)) {
			testParameters.setPlatform(Platform.valueOf(platform));
		} else {
			testParameters.setPlatform(Platform.valueOf(properties
					.getProperty("DefaultPlatform")));
		}
		String OSapiName = values.get("OSapiName");
		if (!"".equals(OSapiName)) {
			testParameters.setOSapiName(OSapiName);
		} else {
			testParameters.setOSapiName(properties
                                .getProperty("DefaultOSapiName"));
		}				
		
		String seeTestPort = values.get("SeeTestPort");
		if (!"".equals(seeTestPort)) {
			testParameters.setSeeTestPort(seeTestPort);
		} else {
			testParameters.setSeeTestPort(mobileProperties
					.getProperty("SeeTestDefaultPort"));
		}
		
		
		if(!executionMode.contains("CROSSBROWSER")){
			String[] jobNamePath = System.getProperty("user.dir").split(Pattern.quote(File.separator));
    		String JobName = jobNamePath[jobNamePath.length-1];
			String Relativepath_c = "C:\\Driver\\"+JobName ;  
			//String Relativepath_c = "C:\\Driver" ;			//System.getProperty("user dir");
			properties.setProperty("ChromeDriverPath", Relativepath_c+"\\chromedriver.exe");
			properties.setProperty("EdgeDriverPath", Relativepath_c+"\\msedgedriver.exe");
			properties.setProperty("InternetExplorerDriverPath", Relativepath_c+"\\IEDriverServer.exe");
			boolean reportPathExists = new File(Relativepath_c).isDirectory();
			if (!reportPathExists) {
				boolean success = (new File(Relativepath_c)).mkdirs();
				if(!success){
					throw new FrameworkException("Folder creation Failed!");
				}
				driver_copy(Relativepath_c,browser);
			}else{
				driver_copy(Relativepath_c,browser);
			}
		}
			
			
	}

	private void generateExtentReports() {
		integrateWithKlov();
		htmlReporter = new ExtentHtmlReporter(resultSummaryManager.getReportPath() + Util.getFileSeparator()
				+ "Extent Result" + Util.getFileSeparator() + "ExtentReport.html");
		extentReport = new ExtentReports();
		extentReport.attachReporter(htmlReporter);
		extentReport.setSystemInfo("Project Name", properties.getProperty("ProjectName"));
		extentReport.setSystemInfo("Framework", "CRAFT Maven");
		extentReport.setSystemInfo("Framework Version", "3.2");
		extentReport.setSystemInfo("Author", "Cognizant");

		htmlReporter.config().setDocumentTitle("CRAFT Extent Report");
		htmlReporter.config().setReportName("Extent Report for CRAFT");
		htmlReporter.config().setTestViewChartLocation(ChartLocation.TOP);
		htmlReporter.config().setTheme(Theme.STANDARD);
	}

	private void integrateWithKlov() {
		String dbHost = properties.getProperty("DBHost");
		String dbPort = properties.getProperty("DBPort");
		if (Boolean.parseBoolean(properties.getProperty("GenerateKlov"))) {
			klovReporter.initMongoDbConnection(dbHost, Integer.valueOf(dbPort));
			klovReporter.setProjectName(properties.getProperty("GenerateKlov"));
			klovReporter.setReportName("CRAFT Reports");
			klovReporter.setKlovUrl(properties.getProperty("KlovURL"));
		}
	}
	public void driver_copy(String relativepath,String browser){
		switch(browser){
		case "CHROME":
			File rundriver = new File(relativepath
					+ Util.getFileSeparator()
					+ "chromedriver.exe");
			boolean driverexist =rundriver.isFile(); 
			if(!driverexist){
				file_copy(relativepath);
			}
		case "EDGE":
			File rundriver2 = new File(relativepath
					+ Util.getFileSeparator()
					+ "msedgedriver.exe");
			boolean driverexist2 =rundriver2.isFile(); 
			if(!driverexist2){
				file_copy(relativepath);
			}
		case "INTERNET_EXPLORER":
			File rundriver1 = new File(relativepath
					+ Util.getFileSeparator()
					+ "IEDriverServer.exe");
			boolean driverexist1 =rundriver1.isFile(); 
			if(!driverexist1){
				file_copy(relativepath);
			}
		}
				
	}
	public void file_copy(String Relative_path){
		try {
			try {
				Taskkill();
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
			File strSourcefile = new File(frameworkParameters.getRelativePath() + Util.getFileSeparator() + "src" + Util.getFileSeparator()
			+ "Driver");
			File strDestination =new File(Relative_path);
			FileUtils.copyDirectory(strSourcefile, strDestination);
		} catch (IOException e) {
			e.printStackTrace();
			throw new FrameworkException(
					"Error in creating run-time datatable: Copying the datatable failed...");
		}
	}
	public void Taskkill() throws InterruptedException {
		   try {
			   String script = frameworkParameters.getRelativePath() + Util.getFileSeparator() + "src" + Util.getFileSeparator()
				+ "test" + Util.getFileSeparator() + "resources"+Util.getFileSeparator()+"TaskKill.vbs";
			   String executable = "C:\\Windows\\System32\\wscript.exe"; 
			   String cmdArr [] = {executable, script};
			   Process p = Runtime.getRuntime ().exec (cmdArr);
		       p.waitFor();
		       Thread.sleep(4000);
			   
		   }
		   catch( IOException e ) {
		      e.printStackTrace();
		   }
		}
	public static void deleteFile(String fileName ){
		try {
			File ApplciationExecuted = new File(fileName);
			if(ApplciationExecuted.exists()) {
				ApplciationExecuted.delete();
			}
		}catch(Exception e) {
			e.printStackTrace();
		}
	}
}