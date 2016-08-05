import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.*;

public class Splitter 
{
	/*
	 * @index0 23# �c��
	 * @index1 25# �c��
	 * @index2 27# �c��
	 * @index3 30# �c��
	 * @index4   �C�@���]�c��(2�c)
	 * @index5   �̫�n��c���@�]���c��
	 */
	static double[] sSubCount = new double[6];
	
	public Splitter(){}
	
	/**
	 * @param args
	 */
	public static void main(String[] args){
		try{
			/*
			 * �פJ��l�q��i��X�f����@�~�C
			 */
			System.out.println("�п�J�q���ɦW(���ݭn.xls)�A�Ҧporders�A��J������Enter: (�ɮ�Ū�����|�T�w��D:)");
			//Catch keyboard input
			BufferedReader bufferRead = new BufferedReader(new InputStreamReader(System.in));
			//Get file path
			String filePath = "d:/" + bufferRead.readLine() + ".xls";
			//String filePath = "D:/order.xls";
			File file = new File(filePath);
			if(!file.exists()){
				throw new Exception("���~!�ɮפ��s�b!�Э��s����ÿ�J���T���|�C");
			}
			
			String fileOutputPath = filePath.substring(0, filePath.length()-4)+"_printlist.xls";
			
			/*
			 * Ū�� & ���
			 */
			Splitter s = new Splitter();
			//Read
			List<String[]> list = s.readFile(filePath);
			//Split
			s.writeFile(list,fileOutputPath);
			//Successful and display the output file path
			System.out.println("���ɦ��\: "+fileOutputPath);
		}
		catch(Exception ex){
			System.out.println(ex.getMessage());
		}
	}
	
	/*
	 * Ū��@�~
	 * @filePath �ɮ׸��|
	 * 
	 */
	public List<String[]> readFile(String filePath) throws Exception{
		List<String[]> list = new ArrayList<String[]>();
		try	{
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
		    	String sender = getCellValue(row, row.getCell(1));
		    	//�p��Ȥ᥼��h����
		    	if(sender.equals("")) continue;
		    	//���o�Ȥ�q�ܡB����A����H�m�W�B�q�ܡB����B�a�}�A�X�f��A�����
		    	String senderTel = getCellValue(row, row.getCell(2));
		    	String senderMobile = getCellValue(row, row.getCell(3));
		    	String senderAddress = getCellValue(row, row.getCell(4));
		    	//String senderEMail = getCellValue(row, row.getCell(5));
		    	
		    	String receiver = getCellValue(row, row.getCell(5));
		    	String receiverTel = getCellValue(row, row.getCell(6));
		    	String receiverMobile = getCellValue(row, row.getCell(7));
		    	String receiverAddress = getCellValue(row, row.getCell(8));
		    	
		    	String sentDate = getCellValue(row, row.getCell(24));
		    	String receiveDate = getCellValue(row, row.getCell(23));
		    	//���o�q��ƶq:23#/25#/27#/30#/�h�B/���B
		    	//�p�G�ƶq�S����g�A�N��@ 0
		    	double s23GiftBox = getCellValue(row, row.getCell(9)).equals("")? (double)0 : Double.parseDouble(getCellValue(row, row.getCell(9))) ;
		    	double s25GiftBox = getCellValue(row, row.getCell(10)).equals("")? (double)0 : Double.parseDouble(getCellValue(row, row.getCell(10))) ;
		    	double s27GiftBox = getCellValue(row, row.getCell(11)).equals("")? (double)0 : Double.parseDouble(getCellValue(row, row.getCell(11))) ;
		    	double s30GiftBox = getCellValue(row, row.getCell(12)).equals("")? (double)0 : Double.parseDouble(getCellValue(row, row.getCell(12))) ;
		    	double s23WhiteBox = getCellValue(row, row.getCell(13)).equals("")? (double)0 : Double.parseDouble(getCellValue(row, row.getCell(13))) ;
		    	double s25WhiteBox = getCellValue(row, row.getCell(14)).equals("")? (double)0 : Double.parseDouble(getCellValue(row, row.getCell(14))) ;
		    	double s27WhiteBox = getCellValue(row, row.getCell(15)).equals("")? (double)0 : Double.parseDouble(getCellValue(row, row.getCell(15))) ;
		    	double s30WhiteBox = getCellValue(row, row.getCell(16)).equals("")? (double)0 : Double.parseDouble(getCellValue(row, row.getCell(16))) ;
		    	double orangeWhiteBox = getCellValue(row, row.getCell(17)).equals("")? (double)0 : Double.parseDouble(getCellValue(row, row.getCell(17))) ;
		    	double sweetOrangeWhiteBox = getCellValue(row, row.getCell(18)).equals("")? (double)0 : Double.parseDouble(getCellValue(row, row.getCell(18))) ;
		    	//�p��Z���a(§��)�`�ƶq
		    	double sCount = Arith.add(Arith.add(Arith.add(s23GiftBox, s25GiftBox), s27GiftBox), s30GiftBox);
		    	//�p��Z���a(�սc)�`�ƶq
		    	double sCount2 = Arith.add(Arith.add(Arith.add(s23WhiteBox, s25WhiteBox), s27WhiteBox), s30WhiteBox);
		    	boolean skipThisRow = false;
		    	//�p�G����H�S�a�}�A��� "����H�S���a�}!(��n�C)"�C
		    	if(receiverAddress.equals("")){
		    		System.out.println("<<< \u6536\u4ef6\u4eba\u6c92\u6709\u5730\u5740!(\u7b2c"+(rownum+1)+"\u5217) >>>");
		    		skipThisRow = true;
		    	}
		    	//�p�G�q��S�����w�ƶq�A���"���q��S�����w�ƶq��!(��n�C)"�C
		    	double totalCount = Arith.add(Arith.add(Arith.add(sCount, orangeWhiteBox), sweetOrangeWhiteBox), sCount2); 
		    	if(totalCount == 0){
		    		System.out.println("<<< \u6709\u8a02\u55ae\u6c92\u6709\u6307\u5b9a\u6578\u91cf\u5594!(\u7b2c"+ (rownum+1) +"\u5217) >>>");
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
	    		double sShipCount = sSplitCount + sCount2 + orangeWhiteBox + sweetOrangeWhiteBox;
	    		int shipIndex = 1;
		    	//�q�w�Z���a�X�f�檺�~�W
		    	//�]�w�U�ŭZ���a���c��(23#,@5#,27#,30#)
		    	sSubCount[0] = s23GiftBox;
		    	sSubCount[1] = s25GiftBox;
		    	sSubCount[2] = s27GiftBox;
		    	sSubCount[3] = s30GiftBox;
		    	//�]�w�����X�i�X�f��
		    	for(int i=1; i<=sSplitCount; i++){
		    		String[] record = new String[10];
		    		record[0] = sender;
		    		record[1] = senderTel;
		    		record[2] = senderMobile;
		    		record[3] = receiver;
		    		record[4] = receiverTel;
		    		record[5] = receiverMobile;
		    		record[6] = receiverAddress;
		    		record[7] = sentDate;
		    		record[8] = receiveDate;
		    		//record[9] = sSplitCount==1? this.genSplitItemName2014(sSubCount,packType) : this.genSplitItemName2014(sSubCount,packType)+"("+i+"/"+(int)sSplitCount+")";
		    		record[9] = sShipCount==1? this.genSplitItemName2015(sSubCount) : this.genSplitItemName2015(sSubCount) + "  ("+shipIndex+"/"+(int)sShipCount+")";
		    		list.add(record);
		    		//�X�f�`��ƶ����[1
		    		shipIndex++;
		    	}
		    	/*
		    	 * �]�w�Z��23#(�սc)���~�W
		    	 * ��~�W: �Z��23#(�սc25��)x1
		    	 * �U��X: \u8302\u8c3723#(\u767d\u7bb125\u65a4)x1
		    	 */
		    	for(int x=1; x<=s23WhiteBox; x++){
		    		String[] record = new String[10];
		    		record[0] = sender;
		    		record[1] = senderTel;
		    		record[2] = senderMobile;
		    		record[3] = receiver;
		    		record[4] = receiverTel;
		    		record[5] = receiverMobile;
		    		record[6] = receiverAddress;
		    		record[7] = sentDate;
		    		record[8] = receiveDate;
		    		record[9] = "\u8302\u8c3723#(\u767d\u7bb125\u65a4)x1" + "  ("+shipIndex+"/"+(int)sShipCount+")";
		    		list.add(record);
		    		//�X�f�`��ƶ����[1
		    		shipIndex++;
		    	}
		    	/*
		    	 * �]�w�Z��25#(�սc)���~�W
		    	 * ��~�W: �Z��25#(�սc25��)x1
		    	 * �U��X: \u8302\u8c3725#(\u767d\u7bb125\u65a4)x1
		    	 */
		    	for(int x=1; x<=s25WhiteBox; x++){
		    		String[] record = new String[10];
		    		record[0] = sender;
		    		record[1] = senderTel;
		    		record[2] = senderMobile;
		    		record[3] = receiver;
		    		record[4] = receiverTel;
		    		record[5] = receiverMobile;
		    		record[6] = receiverAddress;
		    		record[7] = sentDate;
		    		record[8] = receiveDate;
		    		record[9] = "\u8302\u8c3725#(\u767d\u7bb125\u65a4)x1" + "  ("+shipIndex+"/"+(int)sShipCount+")";
		    		list.add(record);
		    		//�X�f�`��ƶ����[1
		    		shipIndex++;
		    	}
		    	/*
		    	 * �]�w�Z��27#(�սc)���~�W
		    	 * ��~�W: �Z��27#(�սc25��)x1
		    	 * �U��X: \u8302\u8c3727#(\u767d\u7bb125\u65a4)x1
		    	 */
		    	for(int x=1; x<=s27WhiteBox; x++){
		    		String[] record = new String[10];
		    		record[0] = sender;
		    		record[1] = senderTel;
		    		record[2] = senderMobile;
		    		record[3] = receiver;
		    		record[4] = receiverTel;
		    		record[5] = receiverMobile;
		    		record[6] = receiverAddress;
		    		record[7] = sentDate;
		    		record[8] = receiveDate;
		    		record[9] = "\u8302\u8c3727#(\u767d\u7bb125\u65a4)x1" + "  ("+shipIndex+"/"+(int)sShipCount+")";
		    		list.add(record);
		    		//�X�f�`��ƶ����[1
		    		shipIndex++;
		    	}
		    	/*
		    	 * �]�w�Z��30#(�սc)���~�W
		    	 * ��~�W: �Z��30#(�սc25��)x1
		    	 * �U��X: \u8302\u8c3730#(\u767d\u7bb125\u65a4)x1
		    	 */
		    	for(int x=1; x<=s30WhiteBox; x++){
		    		String[] record = new String[10];
		    		record[0] = sender;
		    		record[1] = senderTel;
		    		record[2] = senderMobile;
		    		record[3] = receiver;
		    		record[4] = receiverTel;
		    		record[5] = receiverMobile;
		    		record[6] = receiverAddress;
		    		record[7] = sentDate;
		    		record[8] = receiveDate;
		    		record[9] = "\u8302\u8c3730#(\u767d\u7bb125\u65a4)x1" + "  ("+shipIndex+"/"+(int)sShipCount+")";
		    		list.add(record);
		    		//�X�f�`��ƶ����[1
		    		shipIndex++;
		    	}
		    	/*
		    	 * �]�w�h�B���~�W
		    	 * ��~�W: �h�B(�սc25��)x1
		    	 * �U��X: \u67f3\u4e01(\u767d\u7bb125\u65a4)x1
		    	 */
		    	for(int x=1; x<=orangeWhiteBox; x++){
		    		String[] record = new String[10];
		    		record[0] = sender;
		    		record[1] = senderTel;
		    		record[2] = senderMobile;
		    		record[3] = receiver;
		    		record[4] = receiverTel;
		    		record[5] = receiverMobile;
		    		record[6] = receiverAddress;
		    		record[7] = sentDate;
		    		record[8] = receiveDate;
		    		record[9] = "\u67f3\u4e01(\u767d\u7bb125\u65a4)x1" + "  ("+shipIndex+"/"+(int)sShipCount+")";
		    		list.add(record);
		    		//�X�f�`��ƶ����[1
		    		shipIndex++;
		    	}
		    	/*
		    	 * �]�w���B���~�W
		    	 * ��~�W: ���B(�սc25��)x1
		    	 * �U��X: \u751c\u4e01(\u767d\u7bb125\u65a4)x1
		    	 */
		    	for(int x=1; x<=sweetOrangeWhiteBox; x++){
		    		String[] record = new String[10];
		    		record[0] = sender;
		    		record[1] = senderTel;
		    		record[2] = senderMobile;
		    		record[3] = receiver;
		    		record[4] = receiverTel;
		    		record[5] = receiverMobile;
		    		record[6] = receiverAddress;
		    		record[7] = sentDate;
		    		record[8] = receiveDate;
		    		record[9] = "\u751c\u4e01(\u767d\u7bb125\u65a4)x1" + "  ("+shipIndex+"/"+(int)sShipCount+")";
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
		catch(Exception ex){
			ex.printStackTrace();
			throw ex;
		}
		
	}
	
	/*
	 * �ץX������
	 */
	public void writeFile(List<String[]> recordList, String outputFilePath){
		try{
			//�إ��ɮ׶ץX��y
			FileOutputStream out = new FileOutputStream(outputFilePath);
			//�إ�Excel Workbook
			Workbook wb = new HSSFWorkbook();
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
			Cell sentDate = header.createCell(7);
			Cell receiveDate = header.createCell(8);
			Cell itemName = header.createCell(9);
			sender.setCellValue("\u5bc4\u4ef6\u4eba");
			senderTel.setCellValue("\u5bc4\u4ef6\u4eba\u96fb\u8a71");
			senderMobile.setCellValue("\u5bc4\u4ef6\u4eba\u624b\u6a5f");
			receiver.setCellValue("\u6536\u4ef6\u4eba");
			receiverTel.setCellValue("\u6536\u4ef6\u4eba\u96fb\u8a71");
			receiverMobile.setCellValue("\u6536\u4ef6\u4eba\u624b\u6a5f");
			receiverAddress.setCellValue("\u6536\u4ef6\u5730\u5740");
			sentDate.setCellValue("\u5bc4\u4ef6\u65e5");
			receiveDate.setCellValue("\u6536\u4ef6\u65e5");
			itemName.setCellValue("\u54c1\u540d");
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
							sSubCount[i] -= 1;
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
							sSubCount[i] -= 1;
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
	
    public static File[] listFiles(String dir, String extension) 
    {
        final String ext = extension;
    	
        File directory = new File(dir);
        if (!directory.isDirectory()) 
        {
            //System.out.println("No directory provided");
            //return null;
        	directory.mkdir();
        }

        File[] files = directory.listFiles(fileFilter(ext));
        //The listFiles method, with or without a filter does not guarantee any order.
        Arrays.sort(files);
        return files;
    }

    public static FileFilter fileFilter(final String extension)
    {
        //create a FileFilter and override its accept-method
	    return new FileFilter() 
	    {
	        public boolean accept(File file) {
	            //if the file extension is .extension return true, else false
	            if (file.getName().endsWith("."+extension)) 
	            {
	                return true;
	            }
	            return false;
	        }
	    };
    }
}
