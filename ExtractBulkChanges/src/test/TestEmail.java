package test;

import java.util.Properties;

import com.nbty.plm.px.extbulkchgs.EmailUtils;

public class TestEmail {

	public static void main(String[] args) {
		Properties smtpProps = new Properties();
		smtpProps.put("mail.smtp.starttls.enable", "true");
		smtpProps.put("mail.smtp.host", "smtp.gmail.com");
		smtpProps.put("mail.smtp.port", "587");
		smtpProps.put("mail.smtp.auth", "true");
		
		try {
			EmailUtils.sendEmail("juan.lozano@kalypso.com", "juanmlov@gmail.com", "Test", "Body", "C:\\Users\\Kalypso-47\\Desktop\\test.txt", "test.txt", "juanmlov@gmail.com", "", smtpProps);
		} catch(Exception e) {
			e.printStackTrace();
		}
	}

}
