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
import javax.mail.internet.MimeUtility;

import org.apache.log4j.Logger;

import utilities.FileUtils;
import utilities.MailPropReader;
/**
 * GmailSendMailviaTLS
 * @author loren
 *
 */
public class PostMan {
	// ���o�@�~�Ѽ�
	static final String HOST = MailPropReader.readProperty("host");
	static final String PORT = MailPropReader.readProperty("port");
	static final String USER = MailPropReader.readProperty("user");
	static final String PWD = MailPropReader.readProperty("pwd");
	static final String FROM = MailPropReader.readProperty("from");
	static final String SENDER = MailPropReader.readProperty("sender");
	static final String TO = MailPropReader.readProperty("to");
	static final String SUBJECT = MailPropReader.readProperty("subject");
	static final String CONTENT = MailPropReader.readProperty("content");
	static final String BACKUPFOLDER = MailPropReader.readProperty("backupfolder");
	// ���olog4j logger
	static Logger logger = Logger.getLogger(PostMan.class.getName());
	
	public static void main(String args[]) {
		logger.info("****** �}�l�l�H�¿ߧ@�~ *******");
		
		logger.debug("HOST="+HOST);
		logger.debug("PORT="+PORT);
		logger.debug("USER="+USER);
		logger.debug("PWD="+PWD);
		logger.debug("FROM="+FROM);
		logger.debug("SENDER="+SENDER);
		logger.debug("TO="+TO);
		logger.debug("SUBJECT="+SUBJECT);
		logger.debug("CONTENT="+CONTENT);
		
		// �]�wjavamail�Ѽ�		
		Properties props = new Properties();
		props.put("mail.smtp.host", HOST);
		props.put("mail.smtp.auth", "true");
		props.put("mail.smtp.starttls.enable", "true");
		props.put("mail.smtp.port", Integer.parseInt(PORT));
		// �إ�Gmail session
		Session session = Session.getInstance(props, new Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(USER, PWD);
			}
		});
		
		logger.debug("Gmail session is done.");
		
		try{
			// ���oMail����
			String dir = "./�w���/";
			String extension = "xlsx";
			File[] files = FileUtils.listFiles(dir, extension);
			logger.info("����ɮ׼ƶq=["+ files.length +"]");
			// �@�Ӫ���o�@��mail
			for(int i=0; i<files.length; i++){
				File file = files[i];
				// ���o�����ɦW
				String filename = file.getName();
				logger.info("\t�l�H���#"+(i+1)+"=["+filename+"]");
				// �]�w�D��: �����ɦW
				// �]�w���e: ���ɦW���N�������e
				String mailSubject = SUBJECT;
				String mailContent = CONTENT;
				mailSubject = mailSubject.replaceAll("@filename", filename);
				mailContent = mailContent.replaceAll("@filename", filename);
				
				try{
					// ���a����
					Multipart mp = new MimeMultipart();
					
					MimeBodyPart mbp1 = new MimeBodyPart();
					mbp1.setText(mailContent);
				    
					MimeBodyPart mbp2 = new MimeBodyPart();
				    mbp2.attachFile(file);
				    /*
				     * A note on RFC 822 and MIME headers
				     * 
				     * RFC 822 header fields must contain only US-ASCII characters. 
				     * MIME allows non ASCII characters to be present in certain portions 
				     * of certain headers, by encoding those characters. RFC 2047 
				     * specifies the rules for doing this. The MimeUtility class provided 
				     * in this package can be used to to achieve this. Callers of the 
				     * setHeader, addHeader, and addHeaderLine methods are responsible 
				     * for enforcing the MIME requirements for the specified headers. 
				     * In addition, these header fields must be folded (wrapped) before 
				     * being sent if they exceed the line length limitation for the 
				     * transport (1000 bytes for SMTP). Received headers may have been 
				     * folded. The application is responsible for folding and unfolding 
				     * headers as appropriate. 
				     */
				    mbp2.setFileName(MimeUtility.encodeText(filename)); //�קK�����ɦW�|�y��Outlook���쪺�����ɦW�ܦ�XXX.dat 
				    
				    mp.addBodyPart(mbp1);
				    mp.addBodyPart(mbp2);
				    
				    // �]�wMail Message
					Message message = new MimeMessage(session);
				    message.setFrom(new InternetAddress(FROM, SENDER));
				    message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(TO));
				    message.setSubject(mailSubject);
				    message.setContent(mp);
				    // �s�uSMTP�A�Ⱦ�
				    Transport transport = session.getTransport("smtp");
				    transport.connect(HOST, Integer.parseInt(PORT), USER, PWD);
				    // �H�Xmail
				    Transport.send(message);
				    
				    // �N�v�t��h����"�ƥ�"�ؿ�
				    try {
				    	FileUtils.moveFile2(file, BACKUPFOLDER);
				    }catch(IOException ioe){
				    	logger.error("\t\t�h���ɮר�[�ƥ�]�ؿ�����: "+ioe.getMessage());
				    }
				    
				    logger.info("\t\t�l�H�����C");
				    
				    
				}catch(MessagingException ex){
					logger.error("\t\t�l�H����: "+ex.getMessage());
					
					// �N�v�t��h����"�l�H����"�ؿ�
					try {
				    	FileUtils.moveFile2(file, "./�l�H����/");
				    }catch(IOException ioe){
				    	logger.error("\t\t�h���ɮר�[�l�H����]�ؿ�����: "+ioe.getMessage());
				    }
				}
			}
		}catch(IOException e) {
			logger.error("���o�v�t���ɮץ���: "+e.getMessage());
			e.printStackTrace();
		}
		logger.info("****** �l�H�@�~���� *******");
	}
	
}