package cn.ac.qibebt.util;

import javax.mail.*;
import javax.mail.internet.*;
import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import java.util.Date;
import java.util.Properties;

/**
 * Created by luke on 5/19/16.
 */
public class SendMail extends Thread{
    // 邮箱服务器
    private String host = "smtp.cstnet.cn";
    // 这个是你的邮箱用户名
    private String username = "XXXX@qibebt.ac.cn";
    // 你的邮箱密码
    private String password = "*****";

    private String[] mail_to_list;
    private String[] mail_cc_list;
    private String mail_cc = "XXXXX@qibebt.ac.cn";
    private String mail_from = "XXXXX@qibebt.ac.cn";
    private String mail_subject = "邮件自动分发";
    private String mail_body = "邮件测试";
    private String personalName = "标题";
    private String[] attachFiles;

    public SendMail(String[] mail_to_list, String mail_body, String mail_subject) {
        this(mail_to_list, mail_body, mail_subject, null, null, null);
    }
    public SendMail(String[] mail_to_list, String mail_body, String mail_subject, String[] mail_cc_list, String[] attachFiles,String personalName) {
        this.mail_to_list = mail_to_list;
        this.mail_subject = mail_subject;
        this.mail_body = mail_body;
        this.mail_cc_list = mail_cc_list;
        this.attachFiles = attachFiles;
        this.personalName = personalName;
    }
    /**
     * 此段代码用来发送普通电子邮件
     */
    @Override
    public void run() {
        try {
            Properties props = new Properties(); // 获取系统环境
            Authenticator auth = new Email_Autherticator(); // 进行邮件服务器用户认证
            props.put("mail.smtp.host", host);
            props.put("mail.smtp.auth", "true");
            Session session = Session.getDefaultInstance(props, auth);
            // 设置session,和邮件服务器进行通讯。
            MimeMessage message = new MimeMessage(session);
            // message.setContent("foobar, "application/x-foobar"); // 设置邮件格式
            message.setSubject(mail_subject); // 设置邮件主题
            message.setText(mail_body,"utf-8", "html");; // 设置邮件正文
            // message.setHeader(mail_head_name, mail_head_value); // 设置邮件标题
            message.setSentDate(new Date()); // 设置邮件发送日期
            Address address = new InternetAddress(mail_from, personalName);
            message.setFrom(address); // 设置邮件发送者的地址
            message.setDescription("描述");

            for (String mail_to : mail_to_list) {
                Address toAddress = new InternetAddress(mail_to); // 设置邮件接收方的地址
                message.addRecipient(Message.RecipientType.TO, toAddress);
            }
            if (this.mail_cc_list != null) {
                for (String mail_cc : mail_cc_list) {
                    Address ccAddress = new InternetAddress(mail_cc); // 设置邮件抄送方的地址
                    message.addRecipient(Message.RecipientType.CC, ccAddress);
                }
            }

            // creates message part
            MimeBodyPart messageBodyPart = new MimeBodyPart();
            messageBodyPart.setContent(mail_body, "text/html;charset=UTF-8");

            // creates multi-part
            Multipart multipart = new MimeMultipart();
            multipart.addBodyPart(messageBodyPart);
            // adds attachments
            if (attachFiles != null && attachFiles.length > 0) {
                for (String filePath : attachFiles) {
                    MimeBodyPart attachPart = new MimeBodyPart();
//                    System.out.println(filePath);
                    try {
                        attachPart.attachFile(filePath);
                        attachPart.setFileName(MimeUtility.encodeText(new File(filePath).getName()));
                    } catch (IOException ex) {
                        ex.printStackTrace();
                    }

                    multipart.addBodyPart(attachPart);
                }
            }
            Address ccAddress = new InternetAddress(mail_cc); // 设置邮件接收方的地址
            message.addRecipient(Message.RecipientType.CC, ccAddress);
            message.setContent(multipart);
            try{
                Transport.send(message); // 发送邮件
            }catch (Exception e){
                System.out.println("to address fail:"+ Arrays.toString(mail_to_list));
                throw new RuntimeException(e);
            }

            // System.out.println("send ok!");
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
    /**
     * 用来进行服务器对用户的认证
     */
    public class Email_Autherticator extends Authenticator {
        public Email_Autherticator() {
            super();
        }

        public Email_Autherticator(String user, String pwd) {
            super();
            username = user;
            password = pwd;
        }

        public PasswordAuthentication getPasswordAuthentication() {
            return new PasswordAuthentication(username, password);
        }
    }
}
