package cn.ac.qibebt.gaoqian;

public class EmailConfig {
    private String emailFile; //联系人文件路径
    private String email;  //邮箱列
    private String dep; //部门列
    private String subject; //邮件标题
    private String[] mail_cc ; //抄送列表
    private String personName; //发件人姓名
    private String mailBody;

    public EmailConfig() {
    }

    /*public EmailConfig(String emailFile, String email) {
        this.emailFile = emailFile;
        this.email = email;
    }*/

    public String getEmailFile() {
        return emailFile;
    }

    public void setEmailFile(String emailFile) {
        this.emailFile = emailFile;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public String getDep() {
        return dep;
    }

    public void setDep(String dep) {
        this.dep = dep;
    }

    public String getSubject() {
        return subject;
    }

    public void setSubject(String subject) {
        this.subject = subject;
    }

    public String[] getMail_cc() {
        return mail_cc;
    }

    public void setMail_cc(String[] mail_cc) {
        this.mail_cc = mail_cc;
    }

    public String getPersonName() {
        return personName;
    }

    public void setPersonName(String personName) {
        this.personName = personName;
    }

    public String getMailBody() {
        return mailBody;
    }

    public void setMailBody(String mailBody) {
        this.mailBody = mailBody;
    }
}
