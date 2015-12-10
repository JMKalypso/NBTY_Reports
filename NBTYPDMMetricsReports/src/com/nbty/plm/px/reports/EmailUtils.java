package com.nbty.plm.px.reports;

import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Authenticator;
import javax.mail.BodyPart;
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

public class EmailUtils {
	
	private static final Logger logger = Logger.getLogger("GenerateReportsPXLog");
	
	public static void sendEmail(String to, String from, String subject, String messageBody, String attachmentPath, String attachmentFilename,
			final String username, final String password, Properties smtpProps) {
		
		logger.info("******* sendEmail *******");
		
		// Get the Session object.
		Session session = null;

		if ((username != null && !username.trim().equals("")) || (password != null && !password.trim().equals(""))) {
			// username/password is provided
			logger.info("Username/password is provided.");
			session = Session.getInstance(smtpProps, new javax.mail.Authenticator() {
				protected PasswordAuthentication getPasswordAuthentication() {
					return new PasswordAuthentication(username, password);
				}
			});
		} else {
			// either username or password is blank
			logger.info("Either username or password is blank.");
			session = Session.getInstance(smtpProps, new Authenticator() {
			});
		}

		try {
			// Create a default MimeMessage object.
			Message message = new MimeMessage(session);
			logger.info("Message created.");
			// Set From: header field of the header.
			message.setFrom(new InternetAddress(from));
			logger.info("From added: " + from);
			// Set To: header field of the header.
			message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(to));
			logger.info("Recipients added: " + to);
			// Set Subject: header field
			message.setSubject(subject);
			logger.info("Subject added: " + subject);
			// Create the message part
			BodyPart messageBodyPart = new MimeBodyPart();

			// Now set the actual message
			messageBodyPart.setText(messageBody);
			logger.info("Body added: " + messageBody);
			// Create a multipart message
			Multipart multipart = new MimeMultipart();

			// Set text message part
			multipart.addBodyPart(messageBodyPart);
			logger.info("Body part added to message. ");
			// Part two is attachment
			messageBodyPart = new MimeBodyPart();
			String filename = attachmentPath;
			logger.info("Filename: " + filename);
			DataSource source = new FileDataSource(filename);
			logger.info("Data source created.");
			messageBodyPart.setDataHandler(new DataHandler(source));
			logger.info("Data handler created.");
			messageBodyPart.setFileName(attachmentFilename);
			logger.info("Filename set.");
			multipart.addBodyPart(messageBodyPart);
			logger.info("Attachment added.");

			// Send the complete message parts
			message.setContent(multipart);
			logger.info("Content added.");
			// Send message
			Transport.send(message);

		} catch (MessagingException e) {
			logger.info(e.getMessage());
			throw new RuntimeException(e);
		}
	}
	
	public static void sendEmail(String to, String from, String subject, String messageBody, 
			final String username, final String password, Properties smtpProps) {
		
		// Get the Session object.
		Session session = null;

		if ((username != null && !username.trim().equals("")) || (password != null && !password.trim().equals(""))) {
			// username/password is provided
			logger.info("Username/password is provided.");
			session = Session.getInstance(smtpProps, new javax.mail.Authenticator() {
				protected PasswordAuthentication getPasswordAuthentication() {
					return new PasswordAuthentication(username, password);
				}
			});
		} else {
			// either username or password is blank
			logger.info("Either username or password is blank.");
			session = Session.getInstance(smtpProps, new Authenticator() {
			});
		}

		try {
			// Create a default MimeMessage object.
			Message message = new MimeMessage(session);
			logger.info("Message created.");
			// Set From: header field of the header.
			message.setFrom(new InternetAddress(from));
			logger.info("From added: " + from);
			// Set To: header field of the header.
			message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(to));
			logger.info("Recipients added: " + to);
			// Set Subject: header field
			message.setSubject(subject);
			logger.info("Subject added: " + subject);
			// Create the message part
			BodyPart messageBodyPart = new MimeBodyPart();

			// Now set the actual message
			messageBodyPart.setText(messageBody);
			logger.info("Body added: " + messageBody);

			// Send message
			Transport.send(message);

		} catch (MessagingException e) {
			logger.info(e.getMessage());
			throw new RuntimeException(e);
		}
	
	}
}
