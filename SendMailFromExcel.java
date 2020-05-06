import java.io.*;
import java.util.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import javax.mail.*;
import javax.mail.internet.*;
import javax.activation.*;
import javax.mail.Session;
import javax.mail.Transport;


public class SendMailFromExcel  {
	private String username, sender, password, host, msg, subject, attachmentFiles;
	String recipient;
	private Properties properties;
	private Session session;
	private Transport transport;
	Boolean attachmentAvailable = false;
	public static void main(String[] args){
		String name = "";
		String mailid = "";
		int count = 0;
		File file = null;
		Console console = System.console();
		SendMailFromExcel smfx = new SendMailFromExcel();
		smfx.setConfig();
		try{
			Boolean check = false;
			String filename = console.readLine("Excel filename : ");
			file = new File(filename);
			if((filename.endsWith(".xls") || filename.endsWith(".xlsx")) && file.exists()){
				check = true;
			}
			else{
				check = false;
			}
			while(!check){
				console.printf("Error: The specified file either does not exists or is not an excel file!. Please re-enter the file name.\n");
				filename = console.readLine("Excel filename: ");
				file = new File(filename);
				if((filename.endsWith(".xls") || filename.endsWith(".xlsx")) && file.exists()){
					check = true;
				}
				else{
					check = false;
				}
			}
      			FileInputStream fis = new FileInputStream(file);
			Iterator<Row> iterator = null;
			if(filename.endsWith(".xls")){
				HSSFWorkbook workbook = new HSSFWorkbook(fis);
				HSSFSheet sheet = workbook.getSheetAt(0);
				iterator = sheet.iterator();
			}
			else if(filename.endsWith(".xlsx")){
				XSSFWorkbook workbook = new XSSFWorkbook(fis);
				XSSFSheet sheet = workbook.getSheetAt(0);
				iterator = sheet.iterator();
			}
			while (iterator.hasNext()){
				Row row = iterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				count = 0;
				while (cellIterator.hasNext()){
					Cell cell = cellIterator.next();
					count++;
					if(count == 1){
						name = cell.getStringCellValue();
					}
					else if (count == 2){
						mailid = cell.getStringCellValue();
						smfx.sendmail(name, mailid);
					}
				}
				System.out.println("");
			}
      System.out.println("Mail has been sent to every one in the specified excel sheet!");
		}
		catch(Exception e){
			e.printStackTrace();
		}
	}


	private void setConfig(){
		Console console = System.console();
		sender = console.readLine("Your Mail-Id: ");
		username = sender;
		char[] pass = console.readPassword("Password: ");
		password = new String(pass);
		subject = console.readLine("Enter Subject: ");
		host = "smtp.office365.com";
		properties = System.getProperties();
		properties.setProperty("mail.smtp.host", host);
		properties.setProperty("mail.smtp.port", "587");
		properties.setProperty("mail.smtp.auth", "true");
		properties.setProperty("mail.smtp.starttls.enable", "true");
		session = Session.getInstance(properties,
        		new javax.mail.Authenticator() {
        			protected PasswordAuthentication getPasswordAuthentication() {
            				return new PasswordAuthentication(username, password);
            			}
          		});
        	session.setDebug(false);
		try{
			transport = session.getTransport();
			msg = readMessage();
			attachmentFiles = console.readLine("Enter attachment filename: (Skip, if you don't want to sent any attachments | If you need to attach more than one attachments seperate file names with ',') ");
  			if(attachmentFiles.isEmpty()){
  				attachmentAvailable = false;
  			}
  			else{
  				attachmentAvailable = true;
  			}
  			String[] attachments = attachmentFiles.split(",");
  			for(String attachmentFilename : attachments){
  				File f = new File(attachmentFilename);
  				if(!f.exists()){
  					System.out.println("Error: One or more files in the attachment does not exists!. (Tip: Check whether your attachments exists in the working current directory)");
  					System.exit(1);
  				}
  			}
  		}
		catch(NoSuchProviderException nspex){
			nspex.printStackTrace();
		}
		catch(IOException ioe){
			ioe.printStackTrace();
		}
	}

	public void sendmail(String name, String mailid){
		try
		{
			recipient =  mailid;
			if(!attachmentAvailable){
				MimeMessage message = new MimeMessage(session);
				message.setFrom(new InternetAddress(sender));
				message.addRecipient(Message.RecipientType.TO, new InternetAddress(recipient));
				message.setSubject(subject);
				String finalmsg = "Dear " + name + ",\n" + msg + "\n";
				message.setText(finalmsg);
				transport.send(message);
				System.out.println("Mail successfully sent to " + name + " @ " + mailid);
			}
			else{
				MimeMessage message = new MimeMessage(session);
				message.setFrom(new InternetAddress(sender));
				message.addRecipient(Message.RecipientType.TO, new InternetAddress(recipient));
				message.setSubject(subject);
				String finalmsg = "Dear " + name + ",\n" + msg + "\n";

				BodyPart textBodyPart = new MimeBodyPart();
				textBodyPart.setText(finalmsg);

				Multipart multipart = new MimeMultipart();
        			multipart.addBodyPart(textBodyPart);
    				String[] attachments = attachmentFiles.split(",");
    				for(String attachmentFilename : attachments){
    					multipart.addBodyPart(addAttachment(attachmentFilename));
    				}

				message.setContent(multipart);
				transport.send(message);
				System.out.println("Mail successfully sent to " + name + " @ " + mailid);
			}
		}
		catch (MessagingException me){
			me.printStackTrace();
		}
	}

	public String readMessage() throws IOException {
		String line = "";
		String paragraph = "";
		System.out.println("Enter the text: (Type 'exit' on a new line to finish)");
		InputStreamReader inputStreamReader = new InputStreamReader(System.in);
		BufferedReader bufferedReader = new BufferedReader(inputStreamReader);
		do{
			line = bufferedReader.readLine();
			if(!line.equals("exit")){
				paragraph = paragraph + line + "\n";
			}
		}while(!line.equals("exit"));
		return paragraph;
	}

	private MimeBodyPart addAttachment(String attachmentFilename){
		MimeBodyPart attachmentBodyPart = new MimeBodyPart();
		DataSource source = new FileDataSource(attachmentFilename);
		try{
				attachmentBodyPart.setDataHandler(new DataHandler(source));
				attachmentBodyPart.setFileName(attachmentFilename);
		}
		catch(MessagingException e){
			e.printStackTrace();
		}
		return attachmentBodyPart;
	}
}
