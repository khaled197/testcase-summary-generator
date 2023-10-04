package summary_generator;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.xml.parsers.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xddf.usermodel.chart.*;
import org.w3c.dom.*;
import org.xml.sax.SAXException;


public class SummaryGenerator {


    private static int count = 0;
    private static Workbook workbook;
    private static Sheet sheet;
    private static int currentRow = 1;
    private static int passedCount = 0;
    private static int failedCount = 0;
    private static int otherCount = 0;

    
	public static void main(String[] args) {
        try {

            File dir = new File("D:\\Eclipse_IDE\\summary_generator\\Test reports");
            if (!(dir.exists() && dir.isDirectory())) {
            	System.err.println("Error while trying to access directory");
            	return;
            }
            
            File[] xmlFiles = dir.listFiles();

            if (xmlFiles.length == 0) {
                System.err.println("No files to parse");
                return;
            }

            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);

            DocumentBuilder builder = factory.newDocumentBuilder();

            
            workbook = new XSSFWorkbook();
            sheet = workbook.createSheet("Test Results");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Test Case Name");
            header.createCell(1).setCellValue("Result");
            header.createCell(2).setCellValue("Automatic/Manual");

            
            for (File file : xmlFiles) {
                Document doc = builder.parse(file);
                doc.getDocumentElement().normalize();
                System.out.println(file.getName());
                parseNode(doc.getDocumentElement());

            }
            
            createPieChart();
            
	        try (FileOutputStream fileOut = new FileOutputStream("TestResults.xlsx")) {
	            workbook.write(fileOut);
	        }

		} catch (ParserConfigurationException | SAXException | IOException e ) {
        	System.err.println("Error :" + e.getMessage());
		} catch (Exception e) {
		    e.printStackTrace();
			}
		}

    public static void parseNode(Node node) {
        if (node == null) {
            return;
        }

        if (node.getNodeName() == "testcase") {
        	System.out.println(++count + ": " + node.getNodeName());
            printNodeInfo((Element) node);
            writeTestCaseToExcel((Element) node);
        }

        NodeList childNodes = node.getChildNodes();
        for (int i = 0; i < childNodes.getLength(); i++) {
        	parseNode(childNodes.item(i));
        }
    }

    public static void printNodeInfo(Element element) {
		if (element.getNodeType() == Node.ELEMENT_NODE) {
			 Node titleNode = element.getElementsByTagName("title").item(0);
			 Node verdictNode = element.getElementsByTagName("verdict").item(0);
			 String title = "";
			 String result = "";
			 
			 if (titleNode != null) {
				 title = titleNode.getTextContent();
			 }
			 if (verdictNode != null) {
				 result = ((Element) verdictNode).getAttribute("result");
			 }
			 if ("pass".equalsIgnoreCase(result)) {
	                passedCount++;
	            } else if ("fail".equalsIgnoreCase(result)) {
	                failedCount++;
	            } else {
	                otherCount++;
	            }
			 System.out.println("Title: " + title);
			 System.out.println("Result: " + result);
		}}
    
    public static void writeTestCaseToExcel(Element element) {
        if (element.getNodeType() == Node.ELEMENT_NODE) {
            Node titleNode = element.getElementsByTagName("title").item(0);
            Node verdictNode = element.getElementsByTagName("verdict").item(0);

            String title = "";
            String result = "";

            if (titleNode != null) {
                title = titleNode.getTextContent();
            }
            if (verdictNode != null) {
                result = ((Element) verdictNode).getAttribute("result");
            }

            Row row = sheet.createRow(currentRow++);
            row.createCell(0).setCellValue(title);
            row.createCell(1).setCellValue(result);
            row.createCell(2).setCellValue("Automatic"); 
        }
    }
    
    public static void createPieChart() {
        XSSFSheet chartSheet = (XSSFSheet) workbook.createSheet("Test Result Chart");

        Row row = chartSheet.createRow(0);
        row.createCell(0).setCellValue("Passed");
        row.createCell(1).setCellValue(passedCount);
        row = chartSheet.createRow(1);
        row.createCell(0).setCellValue("Failed");
        row.createCell(1).setCellValue(failedCount);
        row = chartSheet.createRow(2);
        row.createCell(0).setCellValue("Other");
        row.createCell(1).setCellValue(otherCount);

        XSSFDrawing drawing = chartSheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 1, 1, 5, 15);

        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText("Test Results Pie Chart");
        chart.setTitleOverlay(false);

        XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);
        XDDFDataSource<String> cat = XDDFDataSourcesFactory.fromStringCellRange(chartSheet, new CellRangeAddress(0, 2, 0, 0));
        XDDFNumericalDataSource<Double> val = XDDFDataSourcesFactory.fromNumericCellRange(chartSheet, new CellRangeAddress(0, 2, 1, 1));
        data.addSeries(cat, val);
        chart.plot(data);

        try (FileOutputStream fileOut = new FileOutputStream("TestResultsWithChart.xlsx")) {
            workbook.write(fileOut);
        } catch(IOException ex) {
            ex.printStackTrace();
        }
    }
    

    
}



