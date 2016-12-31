import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import utilities.Arith;
import utilities.FileUtils;

public class Splitter 
{
	static double[] sSubCount = new double[6];
	static Logger logger = Logger.getLogger(Splitter.class.getName());
	
	public Splitter(){}
	
	/**
	 * @param args
	 */
	public static void main(String[] args){
		logger.info("********* �}�l���@�~ ************");
		try{
			/*
			 * �פJ��l�q��i��X�f����@�~�C
			 */
			String dir = "./";
			String extension = "xlsx";
			File[] files = FileUtils.listFiles(dir, extension);
			String destination = "./�w���/";
			
			File fileDir = new File(destination);
	    	if (!fileDir.exists()) 
	    	{
	    		fileDir.mkdirs();
	    	}
			
			logger.info("�q���ɮ׼ƶq=["+files.length+"]");
			
			for(int i=0; i<files.length; i++){
				File file = files[i];
				String filename = file.getName();
				String fileOutputName = filename.substring(0, filename.length()-5)+"(�¿߱M��)."+extension;
				
				logger.info("\t�B�z���#"+(i+1)+"=["+filename+"]");
				
				if(filename.contains("�¿�")) {
					logger.warn("\t\t��橿��: ���B�z�w��q��!");
					continue;
				}
				
				try{
					/*
					 * Ū�� & ���
					 */
					Splitter s = new Splitter();
					//Read
					List<String[]> list = s.readFile(dir+filename);
					//Split
					s.writeFile(list, destination+fileOutputName);
					//Successful and display the output file path
					logger.info("\t\t��榨�\: "+fileOutputName);
				}catch(Exception ex){
					logger.error("\t\t��楢��: "+ex.getMessage());
				}
			}
		}catch(Exception ex){
			logger.error(ex.getMessage());
		}
		logger.info("********* �������@�~ ************");
	}
	
	/*
	 * Ū��@�~
	 * @filePath �ɮ׸��|
	 * 
	 */
	public List<String[]> readFile(String filePath) throws Exception{
		List<String[]> list = new ArrayList<String[]>();
		
		InputStream inp = new FileInputStream(filePath);
	    Workbook wb = WorkbookFactory.create(inp);
	    //�T�wŪ���Ĥ@��Sheet
	    Sheet sheet = wb.getSheetAt(0);
	    //Ū���Ҧ��C���
	    for (Row row : sheet){
	    	int rownum = row.getRowNum();
	    	//���L�Ĥ@�C���Y���
	    	if(rownum==0) continue;
	    	//���o�Ȥ�W��(Sender)
	    	String sender = getCellValue(row, row.getCell(11));
	    	//�p��Ȥ᥼��h����
	    	if(sender.equals("")) continue;
	    	//���o�Ȥ�q�ܡB����A����H�m�W�B�q�ܡB����B�a�}�A�X�f��A�����
	    	String senderTel = getCellValue(row, row.getCell(12));
	    	String senderMobile = getCellValue(row, row.getCell(13));
	    	//String senderAddress = getCellValue(row, row.getCell(14));
	    	String senderEMail = getCellValue(row, row.getCell(14));
	    	
	    	String receiver = getCellValue(row, row.getCell(15));
	    	String receiverTel = getCellValue(row, row.getCell(16));
	    	String receiverMobile = getCellValue(row, row.getCell(17));
	    	String receiverAddress = getCellValue(row, row.getCell(18));
	    	
	    	String receiveDate = getCellValue(row, row.getCell(19)); //�Ʊ�e�F��
	    	String sentDate = getCellValue(row, row.getCell(24)); //�X�f��
	    	
	    	String receiveTime = getCellValue(row, row.getCell(20)); //����ɬq
	    	String paymentType = getCellValue(row, row.getCell(21)); //�I�ڤ覡
	    	if(receiveTime.equals("")) receiveTime = "�����w";
	    	if(paymentType.equals("")) paymentType = "�v��K";
	    	if(paymentType.equals("�Ȧ�״�")) paymentType = "�v��K";
	    	if(paymentType.contains("�f��I��")) paymentType = "�v��K�ȼֱo";
	    	
	    	
	    	//���o�q��ƶq:23#/25#/27#/30#/�h�B/���B
	    	//�p�G�ƶq�S����g�A�N��@ 0
	    	String s23GiftBox = getCellValue(row, row.getCell(1));
	    	String s25GiftBox = getCellValue(row, row.getCell(2));
	    	String s27GiftBox = getCellValue(row, row.getCell(3));
	    	String s30GiftBox = getCellValue(row, row.getCell(4));
	    	String sOrangeWhiteBox = getCellValue(row, row.getCell(5));
	    	String sSweetOrangeWhiteBox = getCellValue(row, row.getCell(6));
	    	String s23WhiteBox = getCellValue(row, row.getCell(7));
	    	String s25WhiteBox = getCellValue(row, row.getCell(8));
	    	String s27WhiteBox = getCellValue(row, row.getCell(9));
	    	String s30WhiteBox = getCellValue(row, row.getCell(10));
	    	
	    	double d23GiftBox = s23GiftBox.equals("")? (double)0 : Double.parseDouble(s23GiftBox) ;
	    	double d25GiftBox = s25GiftBox.equals("")? (double)0 : Double.parseDouble(s25GiftBox) ;
	    	double d27GiftBox = s27GiftBox.equals("")? (double)0 : Double.parseDouble(s27GiftBox) ;
	    	double d30GiftBox = s30GiftBox.equals("")? (double)0 : Double.parseDouble(s30GiftBox) ;
	    	double dOrangeWhiteBox = sOrangeWhiteBox.equals("")? (double)0 : Double.parseDouble(sOrangeWhiteBox) ;
	    	double dSweetOrangeWhiteBox = sSweetOrangeWhiteBox.equals("")? (double)0 : Double.parseDouble(sSweetOrangeWhiteBox) ;
	    	double d23WhiteBox = s23WhiteBox.equals("")? (double)0 : Double.parseDouble(s23WhiteBox) ;
	    	double d25WhiteBox = s25WhiteBox.equals("")? (double)0 : Double.parseDouble(s25WhiteBox) ;
	    	double d27WhiteBox = s27WhiteBox.equals("")? (double)0 : Double.parseDouble(s27WhiteBox) ;
	    	double d30WhiteBox = s30WhiteBox.equals("")? (double)0 : Double.parseDouble(s30WhiteBox) ;
	    	
	    	//�p��Z���a(§��)�`�ƶq
	    	double sCount = Arith.add(Arith.add(Arith.add(d23GiftBox, d25GiftBox), d27GiftBox), d30GiftBox);
	    	//�p��Z���a(�սc)�`�ƶq
	    	double sCount2 = Arith.add(Arith.add(Arith.add(d23WhiteBox, d25WhiteBox), d27WhiteBox), d30WhiteBox);
	    	boolean skipThisRow = false;
	    	//�p�G����H�S�a�}�A��� "����H�S���a�}!(��n�C)"�C
	    	if(receiverAddress.equals("")){
	    		logger.warn("\t\t\u6536\u4ef6\u4eba\u6c92\u6709\u5730\u5740!(\u7b2c"+(rownum+1)+"\u5217)");
	    		skipThisRow = true;
	    	}
	    	//�p�G�q��S�����w�ƶq�A���"���q��S�����w�ƶq��!(��n�C)"�C
	    	double totalCount = Arith.add(Arith.add(Arith.add(sCount, dOrangeWhiteBox), dSweetOrangeWhiteBox), sCount2); 
	    	if(totalCount == 0 && !sender.equals("n/a")){
	    		logger.warn("\t\t\u6709\u8a02\u55ae\u6c92\u6709\u6307\u5b9a\u6578\u91cf\u5594!(\u7b2c"+ (rownum+1) +"\u5217)");
	    		skipThisRow = true;
	    	}
	    	//�p�G���W�z���@�ر��ΡA�����o����ơC
	    	if(skipThisRow){
	    		continue;
	    	}
	    	
	    	//2014�p�⥴�]�覡�P�B�O���覡
	    	/*
	    	Object[] result = calCarriage2014(sCount);
	    	String packType = (String)result[0];
	    	double sSplitCount = (double)result[1];
	    	if(packType.equals("type1")){
	    		sSubCount[4] = (double)2;
	    		sSubCount[5] = (double)4;
	    	}
	    	else if(packType.equals("type2")){
	    		sSubCount[4] = (double)3;
	    		sSubCount[5] = (double)0;
	    	}
	    	*/
	    	//2015�p�⥴�]�覡�P�B�O���覡
	    	double sSplitCount = calCarriage2015(sCount);
	    	sSubCount[4] = (double)2;
    		sSubCount[5] = (double)0;
	    	//�P�@�a�}�X�f�`���
    		double sShipCount = sSplitCount + sCount2 + dOrangeWhiteBox + dSweetOrangeWhiteBox;
    		int shipIndex = 1;
	    	//�q�w�Z���a�X�f�檺�~�W
	    	//�]�w�U�ŭZ���a���c��(23#,@5#,27#,30#)
	    	sSubCount[0] = d23GiftBox;
	    	sSubCount[1] = d25GiftBox;
	    	sSubCount[2] = d27GiftBox;
	    	sSubCount[3] = d30GiftBox;
	    	//�]�w�����X�i�X�f��
	    	for(int i=1; i<=sSplitCount; i++){
	    		String[] record = new String[12];
	    		record[0] = sender;
	    		record[1] = senderTel;
	    		record[2] = senderMobile;
	    		record[3] = receiver;
	    		record[4] = receiverTel;
	    		record[5] = receiverMobile;
	    		record[6] = receiverAddress;
	    		record[7] = receiveDate;
	    		record[8] = sentDate;
	    		//record[9] = sSplitCount==1? this.genSplitItemName2014(sSubCount,packType) : this.genSplitItemName2014(sSubCount,packType)+"("+i+"/"+(int)sSplitCount+")";
	    		record[9] = sShipCount==1? this.genSplitItemName2015(sSubCount) : this.genSplitItemName2015(sSubCount) + "  ("+(int)sShipCount+"-"+shipIndex+")";
	    		
	    		record[10] = receiveTime;
	    		record[11] = paymentType.equals("�v��K�ȼֱo") && shipIndex==1? "�v��K�ȼֱo" : "�v��K";
	    		
	    		list.add(record);
	    		//�X�f�`��ƶ����[1
	    		shipIndex++;
	    	}
	    	/*
	    	 * �]�w�Z��23#(�սc)���~�W
	    	 * ��~�W: �Z��23#(�սc25��)x1
	    	 * �U��X: \u8302\u8c3723#(\u767d\u7bb125\u65a4)x1
	    	 */
	    	for(int x=1; x<=d23WhiteBox; x++){
	    		String[] record = new String[12];
	    		record[0] = sender;
	    		record[1] = senderTel;
	    		record[2] = senderMobile;
	    		record[3] = receiver;
	    		record[4] = receiverTel;
	    		record[5] = receiverMobile;
	    		record[6] = receiverAddress;
	    		record[7] = receiveDate;
	    		record[8] = sentDate;
	    		record[9] = sShipCount==1? "�Z��23#(�սc25��)x1" : "�Z��23#(�սc25��)x1" + "  ("+(int)sShipCount+"-"+shipIndex+")";
	    		
	    		record[10] = receiveTime;
	    		record[11] = paymentType.equals("�v��K�ȼֱo") && shipIndex==1? "�v��K�ȼֱo" : "�v��K";
	    		
	    		list.add(record);
	    		//�X�f�`��ƶ����[1
	    		shipIndex++;
	    	}
	    	/*
	    	 * �]�w�Z��25#(�սc)���~�W
	    	 * ��~�W: �Z��25#(�սc25��)x1
	    	 * �U��X: \u8302\u8c3725#(\u767d\u7bb125\u65a4)x1
	    	 */
	    	for(int x=1; x<=d25WhiteBox; x++){
	    		String[] record = new String[12];
	    		record[0] = sender;
	    		record[1] = senderTel;
	    		record[2] = senderMobile;
	    		record[3] = receiver;
	    		record[4] = receiverTel;
	    		record[5] = receiverMobile;
	    		record[6] = receiverAddress;
	    		record[7] = receiveDate;
	    		record[8] = sentDate;
	    		record[9] = sShipCount==1? "�Z��25#(�սc25��)x1" : "�Z��25#(�սc25��)x1" + "  ("+(int)sShipCount+"-"+shipIndex+")";

	    		record[10] = receiveTime;
	    		record[11] = paymentType.equals("�v��K�ȼֱo") && shipIndex==1? "�v��K�ȼֱo" : "�v��K";
	    		
	    		list.add(record);
	    		//�X�f�`��ƶ����[1
	    		shipIndex++;
	    	}
	    	/*
	    	 * �]�w�Z��27#(�սc)���~�W
	    	 * ��~�W: �Z��27#(�սc25��)x1
	    	 * �U��X: \u8302\u8c3727#(\u767d\u7bb125\u65a4)x1
	    	 */
	    	for(int x=1; x<=d27WhiteBox; x++){
	    		String[] record = new String[12];
	    		record[0] = sender;
	    		record[1] = senderTel;
	    		record[2] = senderMobile;
	    		record[3] = receiver;
	    		record[4] = receiverTel;
	    		record[5] = receiverMobile;
	    		record[6] = receiverAddress;
	    		record[7] = receiveDate;
	    		record[8] = sentDate;
	    		record[9] = sShipCount==1? "�Z��27#(�սc25��)x1" : "�Z��27#(�սc25��)x1" + "  ("+(int)sShipCount+"-"+shipIndex+")";

	    		record[10] = receiveTime;
	    		record[11] = paymentType.equals("�v��K�ȼֱo") && shipIndex==1? "�v��K�ȼֱo" : "�v��K";
	    		
	    		list.add(record);
	    		//�X�f�`��ƶ����[1
	    		shipIndex++;
	    	}
	    	/*
	    	 * �]�w�Z��30#(�սc)���~�W
	    	 * ��~�W: �Z��30#(�սc25��)x1
	    	 * �U��X: \u8302\u8c3730#(\u767d\u7bb125\u65a4)x1
	    	 */
	    	for(int x=1; x<=d30WhiteBox; x++){
	    		String[] record = new String[12];
	    		record[0] = sender;
	    		record[1] = senderTel;
	    		record[2] = senderMobile;
	    		record[3] = receiver;
	    		record[4] = receiverTel;
	    		record[5] = receiverMobile;
	    		record[6] = receiverAddress;
	    		record[7] = receiveDate;
	    		record[8] = sentDate;
	    		record[9] = sShipCount==1? "�Z��30#(�սc25��)x1" : "�Z��30#(�սc25��)x1" + "  ("+(int)sShipCount+"-"+shipIndex+")";

	    		record[10] = receiveTime;
	    		record[11] = paymentType.equals("�v��K�ȼֱo") && shipIndex==1? "�v��K�ȼֱo" : "�v��K";
	    		
	    		list.add(record);
	    		//�X�f�`��ƶ����[1
	    		shipIndex++;
	    	}
	    	/*
	    	 * �]�w�h�B���~�W
	    	 * ��~�W: �h�B(�սc25��)x1
	    	 * �U��X: \u67f3\u4e01(\u767d\u7bb125\u65a4)x1
	    	 */
	    	for(int x=1; x<=dOrangeWhiteBox; x++){
	    		String[] record = new String[12];
	    		record[0] = sender;
	    		record[1] = senderTel;
	    		record[2] = senderMobile;
	    		record[3] = receiver;
	    		record[4] = receiverTel;
	    		record[5] = receiverMobile;
	    		record[6] = receiverAddress;
	    		record[7] = receiveDate;
	    		record[8] = sentDate;
	    		record[9] = sShipCount==1? "�h�B(�սc25��)x1" : "�h�B(�սc25��)x1" + "  ("+(int)sShipCount+"-"+shipIndex+")";

	    		record[10] = receiveTime;
	    		record[11] = paymentType.equals("�v��K�ȼֱo") && shipIndex==1? "�v��K�ȼֱo" : "�v��K";
	    		
	    		list.add(record);
	    		//�X�f�`��ƶ����[1
	    		shipIndex++;
	    	}
	    	/*
	    	 * �]�w���B���~�W
	    	 * ��~�W: ���B(�սc25��)x1
	    	 * �U��X: \u751c\u4e01(\u767d\u7bb125\u65a4)x1
	    	 */
	    	for(int x=1; x<=dSweetOrangeWhiteBox; x++){
	    		String[] record = new String[12];
	    		record[0] = sender;
	    		record[1] = senderTel;
	    		record[2] = senderMobile;
	    		record[3] = receiver;
	    		record[4] = receiverTel;
	    		record[5] = receiverMobile;
	    		record[6] = receiverAddress;
	    		record[7] = receiveDate;
	    		record[8] = sentDate;
	    		record[9] = sShipCount==1? "���B(�սc25��)x1" : "���B(�սc25��)x1" + "  ("+(int)sShipCount+"-"+shipIndex+")";

	    		record[10] = receiveTime;
	    		record[11] = paymentType.equals("�v��K�ȼֱo") && shipIndex==1? "�v��K�ȼֱo" : "�v��K";
	    		
	    		list.add(record);
	    		//�X�f�`��ƶ����[1
	    		shipIndex++;
	    	}
  		}
		//������y
	    inp.close();
	    //�^�ǵ��G
	    return list;
				
	}
	
	/*
	 * �ץX������
	 */
	public void writeFile(List<String[]> recordList, String outputFilePath){
		try{
			//�إ��ɮ׶ץX��y
			FileOutputStream out = new FileOutputStream(outputFilePath);
			//�إ�Excel Workbook
			Workbook wb = new XSSFWorkbook();
			//�إ�Excel Sheet
			Sheet s = wb.createSheet();
			//�]�wSheet Name
			wb.setSheetName(0, "sheet1" );
			/*
			 * �]�w���Y���:
			 * �H��H, �H��H�q��, ,�H��H���, ����H, ����H�q��, ����H���, ����a�}, �����, �~�W
			 */
			Row header = s.createRow(0);
			Cell sender = header.createCell(0);
			Cell senderTel = header.createCell(1);
			Cell senderMobile = header.createCell(2);
			Cell receiver = header.createCell(3);
			Cell receiverTel = header.createCell(4);
			Cell receiverMobile = header.createCell(5);
			Cell receiverAddress = header.createCell(6);
			Cell receiveDate = header.createCell(7);
			Cell sentDate = header.createCell(8);
			Cell itemName = header.createCell(9);
			Cell receiveTime = header.createCell(10);
			Cell paymentType = header.createCell(11);
			
			sender.setCellValue("�H��H");
			senderTel.setCellValue("�H��H�q��");
			senderMobile.setCellValue("�H��H���");
			receiver.setCellValue("����H");
			receiverTel.setCellValue("����H�q��");
			receiverMobile.setCellValue("����H���");
			receiverAddress.setCellValue("����H�a�}");
			receiveDate.setCellValue("�����");
			sentDate.setCellValue("�H���");
			itemName.setCellValue("�~�W");
			receiveTime.setCellValue("�t�e�ɬq");
			paymentType.setCellValue("�v�t�����");
			
			//��J�����
			int size = recordList.size();
			for(int i=0; i<size; i++){
				String[] record = recordList.get(i);
				Row row = s.createRow(i+1);
				for(int x=0; x<record.length; x++){
					Cell cell = row.createCell(x);
					cell.setCellValue(record[x]);
				}
			}
			
			//�إ��ɮר�������y
			wb.write(out);
			out.close();
		}
		catch(Exception ex){
			ex.printStackTrace();
		}
	}
	
	/*
	 * 2014�p��Z���a�����]�覡�P�B�O:
	 * �@�c100,��c���@�]�]�O100,�T�c���@�]140
	 * 
	 * Type1: �Z���a�c��=N*3+1, �ӥB�`�ƶq��4�c�H�W
	 * Type2: �Z���a�c��=???
	 */
	/*
	public Object[] calCarriage2014(double totalBoxes) throws Exception{
		Object[] result = new Object[2];
		try{
			if(Arith.mod(totalBoxes, 3)==1 && totalBoxes >= 4){
				result[0] = "type1";
				result[1] = Arith.add(Arith.rounddown(Arith.div(Arith.sub(totalBoxes,4),3),0),2);
				return result;
			}
			else{
				result[0] = "type2";
				result[1] = Arith.add(Arith.rounddown(Arith.div(totalBoxes,3),0),Arith.mod(totalBoxes,3)!=0?1:0);
				return result;
			}
		}
		catch(Exception ex){
			throw ex;
		}
	}
	*/
	
	/*
	 * 2015�p��Z���a�����]�覡�P�B�O:
	 * �@�c90,��c���@�]110,�S���T�c���@�]�F.
	 */
	public double calCarriage2015(double totalBoxes) throws Exception{
		double result = 0;
		try{
			result = Arith.roundup(Arith.div(totalBoxes,(double)2), 0);
			return result;
		}
		catch(Exception ex){
			throw ex;
		}
	}
	
	/*
	public String genSplitItemName2014(double[] sSubCount, String packType) throws Exception{
		return  genSplitItemName(sSubCount, packType, "");
	}
	
	public String genSplitItemName(double[] sSubCount, String packType, String itemName) throws Exception{
		String[] spec = new String[]{"23#","25#","27#","30#"};
		try{
			double itemCountPerPackage = sSubCount[4];
			double remainCount2by2 = sSubCount[5];
			//The remain item count for packing...
			double itemCountRemain = itemCountPerPackage;
			
			//Do sub count analysis
			for(int i=0; i<=3; i++){
				double subCount = sSubCount[i];
				//If sub count == 0, continue the next loop.
				if(subCount == 0) continue;
				//Generate the Item Name
				if(subCount < itemCountPerPackage){
					if(subCount < itemCountRemain){
						//Gen partial Item Name
						itemName += spec[i] + "*" + (int)subCount + " ";
						//Set sub count = 0
						sSubCount[i] = 0;
						//itemCountRemain = itemCountRemain - subCount
						itemCountRemain = Arith.sub(itemCountRemain,subCount);
						
						//if itemCountPerPackage = 2
						//   remainCount2by2 = remainCount2by2 - subCount
						if(itemCountPerPackage == 2){
							//Update remainCount2by2
							remainCount2by2 = Arith.sub(remainCount2by2,subCount);
							sSubCount[5] = remainCount2by2;
							//if remainCount2by2 = 0
							//    set  itemCountPerPackage = 3
							if(remainCount2by2 == 0){
								itemCountPerPackage = 3;
								sSubCount[4] = 3;
							}
						}	
					}
					else if(subCount >= itemCountRemain){
						//Gen partial Item Name
						itemName += spec[i] + "*" + (int)itemCountRemain + " ";
						//Set sub count = 0
						sSubCount[i] = Arith.sub(subCount,itemCountRemain);
						//Set item count remain for packing = 0
						itemCountRemain = 0;
						
						//if itemCountPerPackage = 2
						//   remainCount2by2 = remainCount2by2 - subCount
						if(itemCountPerPackage == 2){
							//Update remainCount2by2
							remainCount2by2 = Arith.sub(remainCount2by2,itemCountRemain);
							sSubCount[5] = remainCount2by2;
							//if remainCount2by2 = 0
							//    set  itemCountPerPackage = 3
							if(remainCount2by2 == 0){
								itemCountPerPackage = 3;
								sSubCount[4] = 3;
							}
						}	
					}
				}
				else if(subCount >= itemCountPerPackage){
					//Gen partial Item Name
					itemName += spec[i] + "*" + (int)itemCountRemain + " ";
					//Set sub count = 0
					sSubCount[i] = Arith.sub(subCount,itemCountRemain);
					
					//if itemCountPerPackage = 2
					//   remainCount2by2 = remainCount2by2 - subCount
					if(itemCountPerPackage == 2){
						//Update remainCount2by2
						remainCount2by2 = Arith.sub(remainCount2by2,itemCountRemain);
						sSubCount[5] = remainCount2by2;
						//if remainCount2by2 = 0
						//    set  itemCountPerPackage = 3
						if(remainCount2by2 == 0){
							itemCountPerPackage = 3;
							sSubCount[4] = 3;
						}
					}
					
					//Set item count remain for packing = 0
					itemCountRemain = 0;
				}
				
				//If the remain count = 0, return the Item Name
				if(itemCountRemain == 0){
					return itemName;
				}
			}
		}
		catch(Exception ex){
			throw ex;
		}
		//Return the Item Name that the item count less than one standard package
		return itemName;
	}
	*/
	
	/*
	 * �~�W����
	 */
	public String genSplitItemName2015(double[] sSubCount) throws Exception{
		return  genSplitItemName2015(sSubCount, "");
	}
	/*
	 * �~�W����
	 */
	public String genSplitItemName2015(double[] sSubCount, String itemName) throws Exception{
		String[] spec = new String[]{"23#","25#","27#","30#"};
		String separator = "  ";
		try{
			//�i�H���@�]���c��
			double itemCountPerPackage = sSubCount[4];
						
			//�q23#~30#�̷Ӽƶq�B��c�@�]���W�h�A���~�X�f��~�W�C
			for(int i=0; i<=3; i++){
				//�Y�ӵ��Ū��c��
				double subCount = sSubCount[i];
				//�p�G�c�Ƭ�0�A�~��B�z�U�ӵ���
				if(subCount == 0) continue;
				
				/*
				 * ���ͫ~�W
				 */
				
				//�p�G�c��=1
				if(subCount == 1){
					//�c��=1 && �W�椣�O30#
					if(!spec[i].equals("30#")){
						//�p�G�~�W�٥��զ�
						if(itemName.equals("")){
							//���ͳ����~�W
							itemName += "\u8302\u8c37"+spec[i]+"x1"+separator;
							//�W��c���k�s
							sSubCount[i] = 0;
							//�~���ˬd�U�ӳW�檺�c��
							itemName = genSplitItemName2015(sSubCount, itemName);
							//����j��
							break;
						}
						//�p�G�~�W�w�������զ�
						else if(!itemName.equals("")){
							//���ͫ~�W
							itemName += "\u8302\u8c37"+spec[i]+"x1";
							//�W��c���k�s
							sSubCount[i] = 0;
							//����j��
							break;
						}
					}
					//�c��=1 && �W��O30#
					else{
						//�^�ǫ~�W
						itemName += "\u8302\u8c37"+spec[i]+"x1";
						//����j��
						break;
					}
				}
				//�p�G�c��>=2
				else if(subCount >= itemCountPerPackage){
					//�p�G�~�W�٥��զ�
					if(itemName.equals("")){
						//���ͫ~�W
						itemName += "\u8302\u8c37"+spec[i]+"x2";
						//�W��c�ƴ�2
						sSubCount[i] -= 2;
						//����j��
						break;
					}
					//�p�G�~�W�w�������զ�
					else{
						//���ͫ~�W
						itemName += "\u8302\u8c37"+spec[i]+"x1";
						//�W��c�ƴ�1
						sSubCount[i] -= 1;
						//����j��
						break;
					}
				}
			}
			
			/*
			 * �^�ǫ~�W
			 */
			return itemName.trim();
		}
		catch(Exception ex){
			throw ex;
		}
	}
	
	/*
	 * ���oExcel�������
	 */
	public String getCellValue(Row row, Cell cell) throws Exception{
		String cellValue = "";
		try{
	        //CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
            //System.out.print(cellRef.formatAsString());
            //System.out.print(" - ");
			if(cell == null)
				return cellValue;
				
            switch (cell.getCellType()){
                case Cell.CELL_TYPE_STRING:{
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                }
                case Cell.CELL_TYPE_NUMERIC:{
                    if (DateUtil.isCellDateFormatted(cell)){
                        Date date = cell.getDateCellValue();
                        //
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
                        cellValue = sdf.format(date);
                    } 
                    else{
                        cellValue = Double.toString(cell.getNumericCellValue());
                    }
                    break;
                }
                case Cell.CELL_TYPE_BOOLEAN:{
                    cellValue = Boolean.toString(cell.getBooleanCellValue());
                    break;
                }
                case Cell.CELL_TYPE_FORMULA:{
                    cellValue = cell.getCellFormula();
                    break;
                }
                default:{
                	//Do Nothing! Return empty string.
                	//System.out.println();
                    //throw new Exception("Unspecified Cell Type!");
                }
            }
	    }
		catch(Exception ex){
			throw ex;
		}
		return cellValue;
	}

}
