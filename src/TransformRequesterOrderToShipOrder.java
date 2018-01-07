import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import utilities.Arith;
import utilities.FileUtils;

public class TransformRequesterOrderToShipOrder {
	
	//final static String excelFilePath = "d:/";
	final static String excelFilePath = "./";
	final static String excelFileExtension = "xlsx";
	//final static String excelFileOutputDestination = "d:/�w���/";
	final static String excelFileOutputDestination = "./";
	final static double countPerGiftBoxPackage = 2;
	static double[] giftBoxDetailCountPerRequest = new double[4];
	static List<String[]> shipOrderRecords = new ArrayList<String[]>();
	static File[] requestOrderFiles = null;
	static File shipOrderFileDestination = null;
	static File requestOrderFile = null;
	static String requestOrderFileName = "";
	static String shipOrderFileName = "";
	static String requesterName = "";
	static String requesterCellPhone = "";
	static String requesterPhone = "";
	static String recipientName = "";
	static String recipientCellPhone = "";
	static String recipientPhone = "";
	static String recipientAddress = "";
	static String arrivalDate = ""; 
	static String arrivalTimePeriod = "";
	static String paymentType = "";
	static String shipDate = "";
	static double countOf23GiftBox;
	static double countOf25GiftBox;
	static double countOf27GiftBox;
	static double countOf30GiftBox;
	static double countOfOrangeBox;
	static double countOfSweetOrangeBox;
	static double countOf23CommonBox;
	static double countOf25CommonBox;
	static double countOf27CommonBox;
	static double countOf30CommonBox;
	static double countOfAllSizeGiftBox;
	static double countOfAllSizeCommonBox;
	static double totalCountOfAllBox; 
	static double shipOrderCountOfAllGiftBox;
	static double totalShipOrderCount;
	static int shipOrderSeq;
	static String[] recordTemplate = new String[12];
	static Logger logger = Logger.getLogger(TransformRequesterOrderToShipOrder.class.getName());
	
	
	public TransformRequesterOrderToShipOrder(){}
	
	
	public static void main(String[] args) {
		logger.info("***�}�l���@�~ ");
		try {
			filterRequestOrderFiles();
			logger.info("�q���ɮ׼ƶq=[" + requestOrderFiles.length + "]" );
			makeFileFolderOfShipOrderFiles();
			transformRequestOrderToShipOrder();
		} catch(Exception ex) {
			logger.error("\t\t����: " + ex.getMessage() );
		}
		logger.info("�������@�~ ***");
	}
	
	
	public static void filterRequestOrderFiles() throws Exception {
		requestOrderFiles = FileUtils.listFiles(excelFilePath, excelFileExtension);
	}
	
	
	public static void makeFileFolderOfShipOrderFiles() throws Exception {
		shipOrderFileDestination = new File(excelFileOutputDestination);
    	if (!shipOrderFileDestination.exists()) shipOrderFileDestination.mkdirs();
	}
	
	public static void transformRequestOrderToShipOrder() throws Exception {
		for (int i=0; i<requestOrderFiles.length; i++) {
			shipOrderRecords = new ArrayList<String[]>();
			requestOrderFile = requestOrderFiles[i];
			requestOrderFileName = requestOrderFile.getName();
			shipOrderFileName = requestOrderFileName.substring(0, requestOrderFileName.length()-5)+"(�¿߱M��)."+excelFileExtension;
			
			logger.info("\t�B�z���#" + (i+1) + "=[" + requestOrderFileName + "]" );
			
			if (requestOrderFileName.contains("�¿�")) {
				logger.warn("\t\t����: ���B�z�w��n���q��!");
				continue;
			}
				
			readRequestOrderFile();
			writeShipOrderFile();
				
			logger.info("\t\t���\: "+shipOrderFileName);
		}
	}
	
	
	public static void readRequestOrderFile() throws Exception {
		InputStream fis = new FileInputStream(requestOrderFile);
	    Workbook wb = WorkbookFactory.create(fis);
	    Sheet sheet = wb.getSheetAt(0);
	    
	    for (Row row : sheet) {
	    	int rownum = row.getRowNum();
	    	if (rownum==0) continue;
	    	
	    	shipOrderSeq = 1;
	    	boolean skipThisRow = false;
	    	
	    	requesterName = getCellValue(row.getCell(11));
	    	requesterCellPhone = getCellValue(row.getCell(12));
	    	requesterPhone = getCellValue(row.getCell(13));
	    	recipientName = getCellValue(row.getCell(15));
	    	recipientCellPhone = getCellValue(row.getCell(16));
	    	recipientPhone = getCellValue(row.getCell(17));
	    	recipientAddress = getCellValue(row.getCell(18));
	    	arrivalDate = getCellValue(row.getCell(19)); 
	    	arrivalTimePeriod = getCellValue(row.getCell(20)); 
	    	paymentType = getCellValue(row.getCell(21)); 
	    	shipDate = getCellValue(row.getCell(23)); 
	    	
	    	recordTemplate[0] = requesterName;
	    	recordTemplate[1] = requesterPhone;
	    	recordTemplate[2] = requesterCellPhone;
	    	recordTemplate[3] = recipientName;
	    	recordTemplate[4] = recipientPhone;
	    	recordTemplate[5] = recipientCellPhone;
	    	recordTemplate[6] = recipientAddress;
	    	recordTemplate[7] = arrivalDate;
	    	recordTemplate[8] = shipDate;
	    	recordTemplate[9] = "";
	    	recordTemplate[10] = arrivalTimePeriod;
	    	recordTemplate[11] = "";
    			    	
	    	countOf23GiftBox = getCellValue(row.getCell(1)).equals("") ? (double)0 : Double.parseDouble(getCellValue(row.getCell(1)));
	    	countOf25GiftBox = getCellValue(row.getCell(2)).equals("") ? (double)0 : Double.parseDouble(getCellValue(row.getCell(2)));
	    	countOf27GiftBox = getCellValue(row.getCell(3)).equals("") ? (double)0 : Double.parseDouble(getCellValue(row.getCell(3)));
	    	countOf30GiftBox = getCellValue(row.getCell(4)).equals("") ? (double)0 : Double.parseDouble(getCellValue(row.getCell(4)));
	    	countOfOrangeBox = getCellValue(row.getCell(5)).equals("") ? (double)0 : Double.parseDouble(getCellValue(row.getCell(5)));
	    	countOfSweetOrangeBox = getCellValue(row.getCell(6)).equals("") ? (double)0 : Double.parseDouble(getCellValue(row.getCell(6)));
	    	countOf23CommonBox = getCellValue(row.getCell(7)).equals("") ? (double)0 : Double.parseDouble(getCellValue(row.getCell(7)));
	    	countOf25CommonBox = getCellValue(row.getCell(8)).equals("") ? (double)0 : Double.parseDouble(getCellValue(row.getCell(8)));
	    	countOf27CommonBox = getCellValue(row.getCell(9)).equals("") ? (double)0 : Double.parseDouble(getCellValue(row.getCell(9)));
	    	countOf30CommonBox = getCellValue(row.getCell(10)).equals("") ? (double)0 : Double.parseDouble(getCellValue(row.getCell(10)));
	    	
	    	countOfAllSizeGiftBox = Arith.add(Arith.add(Arith.add(countOf23GiftBox, countOf25GiftBox), countOf27GiftBox), countOf30GiftBox);
	    	countOfAllSizeCommonBox = Arith.add(Arith.add(Arith.add(countOf23CommonBox, countOf25CommonBox), countOf27CommonBox), countOf30CommonBox);
	    	totalCountOfAllBox = Arith.add(Arith.add(Arith.add(countOfAllSizeGiftBox, countOfAllSizeCommonBox), countOfOrangeBox), countOfSweetOrangeBox); 
	    	shipOrderCountOfAllGiftBox = calShipOrderCount(countOfAllSizeGiftBox);
	    	totalShipOrderCount = shipOrderCountOfAllGiftBox + countOfAllSizeCommonBox + countOfOrangeBox + countOfSweetOrangeBox;
    		
	    	giftBoxDetailCountPerRequest[0] = countOf23GiftBox;
	    	giftBoxDetailCountPerRequest[1] = countOf25GiftBox;
	    	giftBoxDetailCountPerRequest[2] = countOf27GiftBox;
	    	giftBoxDetailCountPerRequest[3] = countOf30GiftBox;
	    	
	    	if (requesterName.equals("") || requesterName.equals("n/a")) continue;
	    	if (arrivalTimePeriod.equals("")) arrivalTimePeriod = "�����w";
	    	
	    	if ((paymentType.equals("")) || paymentType.equals("�Ȧ�״�")) 
	    		paymentType = "�v��K";
	    	else if (paymentType.equals("�f��I��")) 
	    		paymentType = "�v��K�ȼֱo";
	    	
	    	if (recipientAddress.equals("")) {
	    		logger.warn("\t\t ����H�S���a�}!(��" + (rownum+1) + "�C)");
	    		skipThisRow = true;
	    	}
	    	if (totalCountOfAllBox==0 && !requesterName.equals("n/a")) {
	    		logger.warn("\t\t ���q��S�����w�ƶq��!(��"+ (rownum+1) +"�C)");
	    		skipThisRow = true;
	    	}
	    	if(skipThisRow) continue;
	    	
    		genShipOrderRecord(shipOrderCountOfAllGiftBox, "");
	    	genShipOrderRecord(countOf23CommonBox, "�Z��23#(�սc)x1");
	    	genShipOrderRecord(countOf25CommonBox, "�Z��25#(�սc)x1");
	    	genShipOrderRecord(countOf27CommonBox, "�Z��27#(�սc)x1");
	    	genShipOrderRecord(countOf30CommonBox, "�Z��30#(�սc)x1");
	    	genShipOrderRecord(countOfOrangeBox, "�h�B(�սc)x1");
	    	genShipOrderRecord(countOfSweetOrangeBox, "���B(�սc)x1");
	    	
  		}
		//������y
	    fis.close();
	   				
	}
	
	
	public static void genShipOrderRecord(double shipOrderCount, String itemName) throws Exception {
		for (int i=1; i<=shipOrderCount; i++) {
    		String[] record = Arrays.copyOf(recordTemplate, recordTemplate.length);
    		
    		if (itemName.equals("")) {
    			record[9] = (totalShipOrderCount==1) ? genShipOrderItemName() : (genShipOrderItemName() + "  (" + (int)totalShipOrderCount + "-" + shipOrderSeq + ")" );
        	} else {
        		record[9] = (totalShipOrderCount==1) ? itemName : (itemName + "  (" + (int)totalShipOrderCount + "-" + shipOrderSeq + ")") ;
        	}
    		
    		record[11] = (paymentType.equals("�v��K�ȼֱo") && shipOrderSeq==1) ? "�v��K�ȼֱo" : "�v��K";
    		
    		shipOrderRecords.add(record);
    		shipOrderSeq++;
    	}
	}
	
	
	public static void writeShipOrderFile() throws Exception {
		FileOutputStream fos = new FileOutputStream(excelFileOutputDestination+shipOrderFileName);
		Workbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet();
		wb.setSheetName(0, "sheet1" );
		
		Row header = sheet.createRow(0);
		Cell requester = header.createCell(0);
		Cell requesterPhone = header.createCell(1);
		Cell requesterCellPhone = header.createCell(2);
		Cell recipient = header.createCell(3);
		Cell recipientPhone = header.createCell(4);
		Cell recipientCellPhone = header.createCell(5);
		Cell recipientAddress = header.createCell(6);
		Cell arrivalDate = header.createCell(7);
		Cell shipDate = header.createCell(8);
		Cell itemName = header.createCell(9);
		Cell arrivalTimePeriod = header.createCell(10);
		Cell paymentType = header.createCell(11);
		
		requester.setCellValue("�H��H");
		requesterPhone.setCellValue("�H��H�q��");
		requesterCellPhone.setCellValue("�H��H���");
		recipient.setCellValue("����H");
		recipientPhone.setCellValue("����H�q��");
		recipientCellPhone.setCellValue("����H���");
		recipientAddress.setCellValue("����H�a�}");
		arrivalDate.setCellValue("�Ʊ�t�F��");
		shipDate.setCellValue("���f��");
		itemName.setCellValue("�~�W");
		arrivalTimePeriod.setCellValue("�t�e�ɬq");
		paymentType.setCellValue("�v�t�����");
		
		for (int i=0; i<shipOrderRecords.size(); i++) {
			String[] shipOrderRecord = shipOrderRecords.get(i);
			Row row = sheet.createRow(i+1);
			for (int x=0; x<shipOrderRecord.length; x++) {
				Cell cell = row.createCell(x);
				cell.setCellValue(shipOrderRecord[x]);
			}
		}
		
		wb.write(fos);
		fos.close();
		
	}
	
	
	public static double calShipOrderCount(double countOfAllGiftBox) throws Exception {
		double result = 0;
		try {
			result = Arith.roundup(Arith.div(countOfAllGiftBox,(double)2), 0);
		} catch(Exception ex) {
			throw ex;
		}
		return result;
	}
	
	
	public static String genShipOrderItemName() throws Exception {
		return  genShipOrderItemName("");
	}
	
	
	public static String genShipOrderItemName(String itemName) throws Exception {
		String[] specOfGiftBox = new String[]{"23#","25#","27#","30#"}; 
		String separatorOfItemName = "  ";
		String itemPrefixName = "�Z��";
		
		for (int i=0; i<=3; i++) {
			double someLevelGiftBoxCount = giftBoxDetailCountPerRequest[i];
			if (someLevelGiftBoxCount == 0) continue;
			
			if (someLevelGiftBoxCount == 1) {
				if (!specOfGiftBox[i].equals("30#")) {
					if (itemName.equals("")) {
						itemName += itemPrefixName + specOfGiftBox[i] + "x1" + separatorOfItemName;
						giftBoxDetailCountPerRequest[i] = 0;
						itemName = genShipOrderItemName(itemName);
						break;
					} else {
						itemName += itemPrefixName + specOfGiftBox[i] + "x1";
						giftBoxDetailCountPerRequest[i] = 0;
						break;
					}
				} else {
					itemName += itemPrefixName + specOfGiftBox[i] + "x1";
					break;
				}
			} else if (someLevelGiftBoxCount >= countPerGiftBoxPackage) {
				if (itemName.equals("")) {
					itemName += itemPrefixName + specOfGiftBox[i] + "x2";
					giftBoxDetailCountPerRequest[i] -= 2;
					break;
				} else {
					itemName += itemPrefixName + specOfGiftBox[i] + "x1";
					giftBoxDetailCountPerRequest[i] -= 1;
					break;
				}
			}
		}
		
		return itemName.trim();

	}
	
	
	public static String getCellValue(Cell cell) throws Exception {
		String cellValue = "";
		
        if(cell == null) return cellValue;
			
        switch (cell.getCellType()){
            case Cell.CELL_TYPE_STRING : {
                cellValue = cell.getRichStringCellValue().getString();
                break;
            }
            case Cell.CELL_TYPE_NUMERIC : {
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
                    cellValue = sdf.format(date);
                } else {
                    cellValue = Double.toString(cell.getNumericCellValue());
                }
                break;
            }
            default : {
            	//Do Nothing! Return empty string.
            	//System.out.println();
                //throw new Exception("Unspecified Cell Type!");
            }
        }
	    
		return cellValue;
	}

	
}
