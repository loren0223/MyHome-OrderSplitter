import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.*;

public class Splitter 
{
	/*
	 * @index0 23# 箱數
	 * @index1 25# 箱數
	 * @index2 27# 箱數
	 * @index3 30# 箱數
	 * @index4   每一打包箱數(2箱)
	 * @index5   最後要兩箱打一包的箱數
	 */
	static double[] sSubCount = new double[6];
	
	public Splitter(){}
	
	/**
	 * @param args
	 */
	public static void main(String[] args){
		try{
			/*
			 * 匯入原始訂單進行出貨單拆單作業。
			 */
			System.out.println("請輸入訂單檔名(不需要.xls)，例如orders，輸入完畢按Enter: (檔案讀取路徑固定為D:)");
			//Catch keyboard input
			BufferedReader bufferRead = new BufferedReader(new InputStreamReader(System.in));
			//Get file path
			String filePath = "d:/" + bufferRead.readLine() + ".xls";
			//String filePath = "D:/order.xls";
			File file = new File(filePath);
			if(!file.exists()){
				throw new Exception("錯誤!檔案不存在!請重新執行並輸入正確路徑。");
			}
			
			String fileOutputPath = filePath.substring(0, filePath.length()-4)+"_printlist.xls";
			
			/*
			 * 讀取 & 拆單
			 */
			Splitter s = new Splitter();
			//Read
			List<String[]> list = s.readFile(filePath);
			//Split
			s.writeFile(list,fileOutputPath);
			//Successful and display the output file path
			System.out.println("轉檔成功: "+fileOutputPath);
		}
		catch(Exception ex){
			System.out.println(ex.getMessage());
		}
	}
	
	/*
	 * 讀單作業
	 * @filePath 檔案路徑
	 * 
	 */
	public List<String[]> readFile(String filePath) throws Exception{
		List<String[]> list = new ArrayList<String[]>();
		try	{
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
		    	String sender = getCellValue(row, row.getCell(1));
		    	//如何客戶未填則忽略
		    	if(sender.equals("")) continue;
		    	//取得客戶電話、手機，收件人姓名、電話、手機、地址，出貨日，收件日
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
		    	//取得訂單數量:23#/25#/27#/30#/柳丁/甜丁
		    	//如果數量沒有填寫，就當作 0
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
		    	//計算茂谷柑(禮盒)總數量
		    	double sCount = Arith.add(Arith.add(Arith.add(s23GiftBox, s25GiftBox), s27GiftBox), s30GiftBox);
		    	//計算茂谷柑(白箱)總數量
		    	double sCount2 = Arith.add(Arith.add(Arith.add(s23WhiteBox, s25WhiteBox), s27WhiteBox), s30WhiteBox);
		    	boolean skipThisRow = false;
		    	//如果收件人沒地址，顯示 "收件人沒有地址!(第n列)"。
		    	if(receiverAddress.equals("")){
		    		System.out.println("<<< \u6536\u4ef6\u4eba\u6c92\u6709\u5730\u5740!(\u7b2c"+(rownum+1)+"\u5217) >>>");
		    		skipThisRow = true;
		    	}
		    	//如果訂單沒有指定數量，顯示"有訂單沒有指定數量喔!(第n列)"。
		    	double totalCount = Arith.add(Arith.add(Arith.add(sCount, orangeWhiteBox), sweetOrangeWhiteBox), sCount2); 
		    	if(totalCount == 0){
		    		System.out.println("<<< \u6709\u8a02\u55ae\u6c92\u6709\u6307\u5b9a\u6578\u91cf\u5594!(\u7b2c"+ (rownum+1) +"\u5217) >>>");
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
	    		double sShipCount = sSplitCount + sCount2 + orangeWhiteBox + sweetOrangeWhiteBox;
	    		int shipIndex = 1;
		    	//訂定茂谷柑出貨單的品名
		    	//設定各級茂谷柑的箱數(23#,@5#,27#,30#)
		    	sSubCount[0] = s23GiftBox;
		    	sSubCount[1] = s25GiftBox;
		    	sSubCount[2] = s27GiftBox;
		    	sSubCount[3] = s30GiftBox;
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
		    		record[7] = sentDate;
		    		record[8] = receiveDate;
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
		    		//出貨總件數項次加1
		    		shipIndex++;
		    	}
		    	/*
		    	 * 設定茂谷25#(白箱)的品名
		    	 * 制式品名: 茂谷25#(白箱25斤)x1
		    	 * 萬國碼: \u8302\u8c3725#(\u767d\u7bb125\u65a4)x1
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
		    		//出貨總件數項次加1
		    		shipIndex++;
		    	}
		    	/*
		    	 * 設定茂谷27#(白箱)的品名
		    	 * 制式品名: 茂谷27#(白箱25斤)x1
		    	 * 萬國碼: \u8302\u8c3727#(\u767d\u7bb125\u65a4)x1
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
		    		//出貨總件數項次加1
		    		shipIndex++;
		    	}
		    	/*
		    	 * 設定茂谷30#(白箱)的品名
		    	 * 制式品名: 茂谷30#(白箱25斤)x1
		    	 * 萬國碼: \u8302\u8c3730#(\u767d\u7bb125\u65a4)x1
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
		    		//出貨總件數項次加1
		    		shipIndex++;
		    	}
		    	/*
		    	 * 設定柳丁的品名
		    	 * 制式品名: 柳丁(白箱25斤)x1
		    	 * 萬國碼: \u67f3\u4e01(\u767d\u7bb125\u65a4)x1
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
		    		//出貨總件數項次加1
		    		shipIndex++;
		    	}
		    	/*
		    	 * 設定甜丁的品名
		    	 * 制式品名: 甜丁(白箱25斤)x1
		    	 * 萬國碼: \u751c\u4e01(\u767d\u7bb125\u65a4)x1
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
		    		//出貨總件數項次加1
		    		shipIndex++;
		    	}
      		}
    		//關閉串流
		    inp.close();
		    //回傳結果
		    return list;
		}
		catch(Exception ex){
			ex.printStackTrace();
			throw ex;
		}
		
	}
	
	/*
	 * 匯出拆單明細
	 */
	public void writeFile(List<String[]> recordList, String outputFilePath){
		try{
			//建立檔案匯出串流
			FileOutputStream out = new FileOutputStream(outputFilePath);
			//建立Excel Workbook
			Workbook wb = new HSSFWorkbook();
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
							sSubCount[i] -= 1;
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
							sSubCount[i] -= 1;
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
