package cn.ac.qibebt.util;


import cn.ac.qibebt.gaoqian.AutoSendEmail;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;

public class MailServer {

	public void send(String[] mail_to_list, String mail_body, String mail_subject) {
		SendMail sendMail = new SendMail(mail_to_list, mail_body, mail_subject);
		sendMail.start();
	}
	public void send(String[] mail_to_list, String mail_body, String mail_subject, String[] mail_cc_list,String[] attachFiles, String personName) {
		SendMail sendMail = new SendMail(mail_to_list, mail_body, mail_subject, mail_cc_list,attachFiles,personName);
		sendMail.start();
	}

	public static void main(String[] args) {
		MailServer sendmail = new MailServer();
		try {
			String bodyFile = "/tmp/mail_body.md";
			String source= new String(Files.readAllBytes(Paths.get(bodyFile)), StandardCharsets.UTF_8);
			String mail_body = MarkDownUtil.toHtml(source);
			System.out.println(mail_body);
//			String mail_body = "您好！  <br/>    " + " 附件 " ;
			String[] attachFiles = new String[]{"/tmp/中文.zip"};
			sendmail.send(new String[]{"qiaoyh@qibebt.ac.cn"}, mail_body, "附件");
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

}
