import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Splitter 
{
	static double[] sSubCount = new double[6];
	static Logger logger = Logger.getRootLogger();
	
	public Splitter(){}
	
	/**
	 * @param args
	 */
	public static void main(String[] args){
		try{
			/*
			 * 匯入原始訂單進行出貨單拆單作業。
			 */
			logger.info("********* 開始拆單作業 ************");
			
			String dir = "./";
			String extension = "xlsx";
			
			File[] files = FileUtils.listFiles(dir, extension);
			
			logger.info("檔案(*.xlsx)數量 is "+files.length);
			
			for(int i=0; i<files.length; i++){
				File file = files[i];
				String filename = file.getName();
				String fileOutputName = filename.substring(0, filename.length()-5)+"(黑貓專用)."+extension;
				
				logger.info("處理拆單#("+(i+1)+"): "+file.getName());
				
				if(filename.contains("黑貓")) {
					logger.warn("拆單忽略: 不處理已拆訂單!");
					continue;
				}
				
				try{
					/*
					 * 讀取 & 拆單
					 */
					Splitter s = new Splitter();
					//Read
					List<String[]> list = s.readFile(dir+filename);
					//Split
					s.writeFile(list, dir+fileOutputName);
					//Successful and display the output file path
					logger.info("拆單成功: "+fileOutputName);
				}catch(Exception ex){
					logger.error("拆單失敗: "+ex.getMessage());
				}
			}
			
		}
		catch(Exception ex){
			logger.error(ex.getMessage());
		}
	}
	
	/*
	 * 讀單作業
	 * @filePath 檔案路徑
	 * 
	 */
	public List<String[]> readFile(String filePath) throws Exception{
		List<String[]> list = new ArrayList<String[]>();
		
		InputStream inp = new FileInputStream(filePath);
	    Workbook wb = WorkbookFactory.create(inp);
	    //固定讀取第一個Sheet
	    Sheet sheet = wb.getSheetAt(0);
	    //讀取所有列資料
	    for (Row row : sheet){
	    	int rownum = row.getRowNum();
	    	//略過第一列表頭欄位
	    	if(rownum==0) continue;
	    	//取得客戶名稱(Sender)
	    	String sender = getCellValue(row, row.getCell(11));
	    	//如何客戶未填則忽略
	    	if(sender.equals("")) continue;
	    	//取得客戶電話、手機，收件人姓名、電話、手機、地址，出貨日，收件日
	    	String senderTel = getCellValue(row, row.getCell(12));
	    	String senderMobile = getCellValue(row, row.getCell(13));
	    	//String senderAddress = getCellValue(row, row.getCell(14));
	    	String senderEMail = getCellValue(row, row.getCell(14));
	    	
	    	String receiver = getCellValue(row, row.getCell(15));
	    	String receiverTel = getCellValue(row, row.getCell(16));
	    	String receiverMobile = getCellValue(row, row.getCell(17));
	    	String receiverAddress = getCellValue(row, row.getCell(18));
	    	
	    	String receiveDate = getCellValue(row, row.getCell(19));
	    	String sentDate = getCellValue(row, row.getCell(24));
	    	
	    	//取得訂單數量:23#/25#/27#/30#/柳丁/甜丁
	    	//如果數量沒有填寫，就當作 0
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
	    	
	    	//計算茂谷柑(禮盒)總數量
	    	double sCount = Arith.add(Arith.add(Arith.add(d23GiftBox, d25GiftBox), d27GiftBox), d30GiftBox);
	    	//計算茂谷柑(白箱)總數量
	    	double sCount2 = Arith.add(Arith.add(Arith.add(d23WhiteBox, d25WhiteBox), d27WhiteBox), d30WhiteBox);
	    	boolean skipThisRow = false;
	    	//如果收件人沒地址，顯示 "收件人沒有地址!(第n列)"。
	    	if(receiverAddress.equals("")){
	    		logger.warn("<<< \u6536\u4ef6\u4eba\u6c92\u6709\u5730\u5740!(\u7b2c"+(rownum+1)+"\u5217) >>>");
	    		skipThisRow = true;
	    	}
	    	//如果訂單沒有指定數量，顯示"有訂單沒有指定數量喔!(第n列)"。
	    	double totalCount = Arith.add(Arith.add(Arith.add(sCount, dOrangeWhiteBox), dSweetOrangeWhiteBox), sCount2); 
	    	if(totalCount == 0){
	    		logger.warn("<<< \u6709\u8a02\u55ae\u6c92\u6709\u6307\u5b9a\u6578\u91cf\u5594!(\u7b2c"+ (rownum+1) +"\u5217) >>>");
	    		skipThisRow = true;
	    	}
	    	//如果有上述任一種情形，忽略這筆資料。
	    	if(skipThisRow){
	    		continue;
	    	}
	    	
	    	//2014計算打包方式與運費的方式
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
	    	//2015計算打包方式與運費的方式
	    	double sSplitCount = calCarriage2015(sCount);
	    	sSubCount[4] = (double)2;
    		sSubCount[5] = (double)0;
	    	//同一地址出貨總件數
    		double sShipCount = sSplitCount + sCount2 + dOrangeWhiteBox + dSweetOrangeWhiteBox;
    		int shipIndex = 1;
	    	//訂定茂谷柑出貨單的品名
	    	//設定各級茂谷柑的箱數(23#,@5#,27#,30#)
	    	sSubCount[0] = d23GiftBox;
	    	sSubCount[1] = d25GiftBox;
	    	sSubCount[2] = d27GiftBox;
	    	sSubCount[3] = d30GiftBox;
	    	//設定分成幾張出貨單
	    	for(int i=1; i<=sSplitCount; i++){
	    		String[] record = new String[10];
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
	    		record[9] = sShipCount==1? this.genSplitItemName2015(sSubCount) : this.genSplitItemName2015(sSubCount) + "  ("+shipIndex+"/"+(int)sShipCount+")";
	    		list.add(record);
	    		//出貨總件數項次加1
	    		shipIndex++;
	    	}
	    	/*
	    	 * 設定茂谷23#(白箱)的品名
	    	 * 制式品名: 茂谷23#(白箱25斤)x1
	    	 * 萬國碼: \u8302\u8c3723#(\u767d\u7bb125\u65a4)x1
	    	 */
	    	for(int x=1; x<=d23WhiteBox; x++){
	    		String[] record = new String[10];
	    		record[0] = sender;
	    		record[1] = senderTel;
	    		record[2] = senderMobile;
	    		record[3] = receiver;
	    		record[4] = receiverTel;
	    		record[5] = receiverMobile;
	    		record[6] = receiverAddress;
	    		record[7] = receiveDate;
	    		record[8] = sentDate;
	    		record[9] = "\u8302\u8c3723#(\u767d\u7bb125\u65a4)x1" + "  ("+shipIndex+"/"+(int)sShipCount+")";
	    		list.add(record);
	    		//出貨總件數項次加1
	    		shipIndex++;
	    	}
	    	/*
	    	 * 設定茂谷25#(白箱)的品名
	    	 * 制式品名: 茂谷25#(白箱25斤)x1
	    	 * 萬國碼: \u8302\u8c3725#(\u767d\u7bb125\u65a4)x1
	    	 */
	    	for(int x=1; x<=d25WhiteBox; x++){
	    		String[] record = new String[10];
	    		record[0] = sender;
	    		record[1] = senderTel;
	    		record[2] = senderMobile;
	    		record[3] = receiver;
	    		record[4] = receiverTel;
	    		record[5] = receiverMobile;
	    		record[6] = receiverAddress;
	    		record[7] = receiveDate;
	    		record[8] = sentDate;
	    		record[9] = "\u8302\u8c3725#(\u767d\u7bb125\u65a4)x1" + "  ("+shipIndex+"/"+(int)sShipCount+")";
	    		list.add(record);
	    		//出貨總件數項次加1
	    		shipIndex++;
	    	}
	    	/*
	    	 * 設定茂谷27#(白箱)的品名
	    	 * 制式品名: 茂谷27#(白箱25斤)x1
	    	 * 萬國碼: \u8302\u8c3727#(\u767d\u7bb125\u65a4)x1
	    	 */
	    	for(int x=1; x<=d27WhiteBox; x++){
	    		String[] record = new String[10];
	    		record[0] = sender;
	    		record[1] = senderTel;
	    		record[2] = senderMobile;
	    		record[3] = receiver;
	    		record[4] = receiverTel;
	    		record[5] = receiverMobile;
	    		record[6] = receiverAddress;
	    		record[7] = receiveDate;
	    		record[8] = sentDate;
	    		record[9] = "\u8302\u8c3727#(\u767d\u7bb125\u65a4)x1" + "  ("+shipIndex+"/"+(int)sShipCount+")";
	    		list.add(record);
	    		//出貨總件數項次加1
	    		shipIndex++;
	    	}
	    	/*
	    	 * 設定茂谷30#(白箱)的品名
	    	 * 制式品名: 茂谷30#(白箱25斤)x1
	    	 * 萬國碼: \u8302\u8c3730#(\u767d\u7bb125\u65a4)x1
	    	 */
	    	for(int x=1; x<=d30WhiteBox; x++){
	    		String[] record = new String[10];
	    		record[0] = sender;
	    		record[1] = senderTel;
	    		record[2] = senderMobile;
	    		record[3] = receiver;
	    		record[4] = receiverTel;
	    		record[5] = receiverMobile;
	    		record[6] = receiverAddress;
	    		record[7] = receiveDate;
	    		record[8] = sentDate;
	    		record[9] = "\u8302\u8c3730#(\u767d\u7bb125\u65a4)x1" + "  ("+shipIndex+"/"+(int)sShipCount+")";
	    		list.add(record);
	    		//出貨總件數項次加1
	    		shipIndex++;
	    	}
	    	/*
	    	 * 設定柳丁的品名
	    	 * 制式品名: 柳丁(白箱25斤)x1
	    	 * 萬國碼: \u67f3\u4e01(\u767d\u7bb125\u65a4)x1
	    	 */
	    	for(int x=1; x<=dOrangeWhiteBox; x++){
	    		String[] record = new String[10];
	    		record[0] = sender;
	    		record[1] = senderTel;
	    		record[2] = senderMobile;
	    		record[3] = receiver;
	    		record[4] = receiverTel;
	    		record[5] = receiverMobile;
	    		record[6] = receiverAddress;
	    		record[7] = receiveDate;
	    		record[8] = sentDate;
	    		record[9] = "\u67f3\u4e01(\u767d\u7bb125\u65a4)x1" + "  ("+shipIndex+"/"+(int)sShipCount+")";
	    		list.add(record);
	    		//出貨總件數項次加1
	    		shipIndex++;
	    	}
	    	/*
	    	 * 設定甜丁的品名
	    	 * 制式品名: 甜丁(白箱25斤)x1
	    	 * 萬國碼: \u751c\u4e01(\u767d\u7bb125\u65a4)x1
	    	 */
	    	for(int x=1; x<=dSweetOrangeWhiteBox; x++){
	    		String[] record = new String[10];
	    		record[0] = sender;
	    		record[1] = senderTel;
	    		record[2] = senderMobile;
	    		record[3] = receiver;
	    		record[4] = receiverTel;
	    		record[5] = receiverMobile;
	    		record[6] = receiverAddress;
	    		record[7] = receiveDate;
	    		record[8] = sentDate;
	    		record[9] = "\u751c\u4e01(\u767d\u7bb125\u65a4)x1" + "  ("+shipIndex+"/"+(int)sShipCount+")";
	    		list.add(record);
	    		//出貨總件數項次加1
	    		shipIndex++;
	    	}
  		}
		//關閉串流
	    inp.close();
	    //回傳結果
	    return list;
				
	}
	
	/*
	 * 匯出拆單明細
	 */
	public void writeFile(List<String[]> recordList, String outputFilePath){
		try{
			//建立檔案匯出串流
			FileOutputStream out = new FileOutputStream(outputFilePath);
			//建立Excel Workbook
			Workbook wb = new XSSFWorkbook();
			//建立Excel Sheet
			Sheet s = wb.createSheet();
			//設定Sheet Name
			wb.setSheetName(0, "sheet1" );
			/*
			 * 設定表頭欄位:
			 * 寄件人, 寄件人電話, ,寄件人手機, 收件人, 收件人電話, 收件人手機, 收件地址, 收件日, 品名
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
			sender.setCellValue("\u5bc4\u4ef6\u4eba");
			senderTel.setCellValue("\u5bc4\u4ef6\u4eba\u96fb\u8a71");
			senderMobile.setCellValue("\u5bc4\u4ef6\u4eba\u624b\u6a5f");
			receiver.setCellValue("\u6536\u4ef6\u4eba");
			receiverTel.setCellValue("\u6536\u4ef6\u4eba\u96fb\u8a71");
			receiverMobile.setCellValue("\u6536\u4ef6\u4eba\u624b\u6a5f");
			receiverAddress.setCellValue("\u6536\u4ef6\u5730\u5740");
			receiveDate.setCellValue("\u6536\u4ef6\u65e5");
			sentDate.setCellValue("\u5bc4\u4ef6\u65e5");
			itemName.setCellValue("\u54c1\u540d");
			//輸入拆單資料
			int size = recordList.size();
			for(int i=0; i<size; i++){
				String[] record = recordList.get(i);
				Row row = s.createRow(i+1);
				for(int x=0; x<record.length; x++){
					Cell cell = row.createCell(x);
					cell.setCellValue(record[x]);
				}
			}
			
			//建立檔案並關閉串流
			wb.write(out);
			out.close();
		}
		catch(Exception ex){
			ex.printStackTrace();
		}
	}
	
	/*
	 * 2014計算茂谷柑的打包方式與運費:
	 * 一箱100,兩箱打一包也是100,三箱打一包140
	 * 
	 * Type1: 茂谷柑箱數=N*3+1, 而且總數量有4箱以上
	 * Type2: 茂谷柑箱數=???
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
	 * 2015計算茂谷柑的打包方式與運費:
	 * 一箱90,兩箱打一包110,沒有三箱打一包了.
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
	 * 品名產生
	 */
	public String genSplitItemName2015(double[] sSubCount) throws Exception{
		return  genSplitItemName2015(sSubCount, "");
	}
	/*
	 * 品名產生
	 */
	public String genSplitItemName2015(double[] sSubCount, String itemName) throws Exception{
		String[] spec = new String[]{"23#","25#","27#","30#"};
		String separator = "  ";
		try{
			//可以打一包的箱數
			double itemCountPerPackage = sSubCount[4];
						
			//從23#~30#依照數量、兩箱一包的規則，產品出貨單品名。
			for(int i=0; i<=3; i++){
				//某個等級的箱數
				double subCount = sSubCount[i];
				//如果箱數為0，繼續處理下個等級
				if(subCount == 0) continue;
				
				/*
				 * 產生品名
				 */
				
				//如果箱數=1
				if(subCount == 1){
					//箱數=1 && 規格不是30#
					if(!spec[i].equals("30#")){
						//如果品名還未組成
						if(itemName.equals("")){
							//產生部分品名
							itemName += "\u8302\u8c37"+spec[i]+"x1"+separator;
							//規格箱數歸零
							sSubCount[i] = 0;
							//繼續檢查下個規格的箱數
							itemName = genSplitItemName2015(sSubCount, itemName);
							//停止迴圈
							break;
						}
						//如果品名已有部分組成
						else if(!itemName.equals("")){
							//產生品名
							itemName += "\u8302\u8c37"+spec[i]+"x1";
							//規格箱數歸零
							sSubCount[i] = 0;
							//停止迴圈
							break;
						}
					}
					//箱數=1 && 規格是30#
					else{
						//回傳品名
						itemName += "\u8302\u8c37"+spec[i]+"x1";
						//停止迴圈
						break;
					}
				}
				//如果箱數>=2
				else if(subCount >= itemCountPerPackage){
					//如果品名還未組成
					if(itemName.equals("")){
						//產生品名
						itemName += "\u8302\u8c37"+spec[i]+"x2";
						//規格箱數減2
						sSubCount[i] -= 2;
						//停止迴圈
						break;
					}
					//如果品名已有部分組成
					else{
						//產生品名
						itemName += "\u8302\u8c37"+spec[i]+"x1";
						//規格箱數減1
						sSubCount[i] -= 1;
						//停止迴圈
						break;
					}
				}
			}
			
			/*
			 * 回傳品名
			 */
			return itemName.trim();
		}
		catch(Exception ex){
			throw ex;
		}
	}
	
	/*
	 * 取得Excel表格欄位值
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
