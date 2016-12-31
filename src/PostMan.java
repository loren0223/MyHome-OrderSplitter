import java.io.File;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.log4j.Logger;

import utilities.FileUtils;
import utilities.MailPropReader;
/**
 * GmailSendMailviaTLS
 * @author loren
 *
 */
public class PostMan {
	// 取得作業參數
	static final String HOST = MailPropReader.readProperty("host");
	static final String PORT = MailPropReader.readProperty("port");
	static final String USER = MailPropReader.readProperty("user");
	static final String PWD = MailPropReader.readProperty("pwd");
	static final String FROM = MailPropReader.readProperty("from");
	static final String SENDER = MailPropReader.readProperty("sender");
	static final String TO = MailPropReader.readProperty("to");
	static final String SUBJECT = MailPropReader.readProperty("subject");
	static final String CONTENT = MailPropReader.readProperty("content");
	// 取得log4j logger
	static Logger logger = Logger.getLogger(PostMan.class.getName());
	
	public static void main(String args[]) {
		logger.info("****** 開始郵寄黑貓作業 *******");
		
		logger.debug("HOST="+HOST);
		logger.debug("PORT="+PORT);
		logger.debug("USER="+USER);
		logger.debug("PWD="+PWD);
		logger.debug("FROM="+FROM);
		logger.debug("SENDER="+SENDER);
		logger.debug("TO="+TO);
		logger.debug("SUBJECT="+SUBJECT);
		logger.debug("CONTENT="+CONTENT);
		
		// 設定javamail參數		
		Properties props = new Properties();
		props.put("mail.smtp.host", HOST);
		props.put("mail.smtp.auth", "true");
		props.put("mail.smtp.starttls.enable", "true");
		props.put("mail.smtp.port", Integer.parseInt(PORT));
		// 建立Gmail session
		Session session = Session.getInstance(props, new Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(USER, PWD);
			}
		});
		
		logger.debug("Gmail session is done.");
		
		try{
			// 取得Mail附件
			String dir = "./已拆單/";
			String extension = "xlsx";
			File[] files = FileUtils.listFiles(dir, extension);
			logger.info("拆單檔案數量=["+ files.length +"]");
			// 一個附件發一封mail
			for(int i=0; i<files.length; i++){
				File file = files[i];
				// 取得附件檔名
				String filename = file.getName();
				logger.info("\t郵寄拆單#"+(i+1)+"=["+filename+"]");
				// 設定主旨: 等於檔名
				// 設定內容: 用檔名取代部分內容
				String mailSubject = SUBJECT;
				String mailContent = CONTENT;
				mailSubject = mailSubject.replaceAll("@filename", filename);
				mailContent = mailContent.replaceAll("@filename", filename);
				
				try{
					// 夾帶附件
					Multipart mp = new MimeMultipart();
					MimeBodyPart mbp1 = new MimeBodyPart();
				    MimeBodyPart mbp2 = new MimeBodyPart();
				    mbp1.setText(mailContent);
				    mbp2.attachFile(file);
				    mp.addBodyPart(mbp1);
				    mp.addBodyPart(mbp2);
				    
				    // 設定Mail Message
					Message message = new MimeMessage(session);
				    message.setFrom(new InternetAddress(FROM, SENDER));
				    message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(TO));
				    message.setSubject(mailSubject);
				    message.setContent(mp);
				    // 連線SMTP服務器
				    Transport transport = session.getTransport("smtp");
				    transport.connect(HOST, Integer.parseInt(PORT), USER, PWD);
				    // 寄出mail
				    Transport.send(message);
				    
				    // 刪除拆單
				    file.delete();
				    
				    logger.info("\t\t郵寄完成。");
				    
				    // 將宅配單搬移到"已郵寄"目錄
				    /*
				    try {
				    	FileUtils.moveFile2(file, "./已郵寄/");
				    }catch(IOException ioe){
				    	logger.error("\t\t搬移檔案到[已郵寄]目錄失敗: "+ioe.getMessage());
				    }
				    */
				}catch(MessagingException ex){
					logger.error("\t\t郵寄失敗: "+ex.getMessage());
					
					// 將宅配單搬移到"郵寄失敗"目錄
					try {
				    	FileUtils.moveFile2(file, "./郵寄失敗/");
				    }catch(IOException ioe){
				    	logger.error("\t\t搬移檔案到[郵寄失敗]目錄失敗: "+ioe.getMessage());
				    }
				}
			}
		}catch(IOException e) {
			logger.error("取得宅配單檔案失敗: "+e.getMessage());
			e.printStackTrace();
		}
		logger.info("****** 郵寄作業結束 *******");
	}
	
}