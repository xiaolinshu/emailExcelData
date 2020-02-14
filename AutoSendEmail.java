package cn.ac.qibebt.gaoqian;

import cn.ac.qibebt.util.LicenseUtil;
import cn.ac.qibebt.util.MailServer;
import cn.ac.qibebt.util.MarkDownUtil;
import com.aspose.cells.*;

import java.io.File;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;

public class AutoSendEmail {
    private static final String dir = "/tmp/11";
    private MapValueList attachFiles = new MapValueList(); //部门名称：[附件1，2]
    private MapValueList depEmail = new MapValueList(); //部门名称：[联系人邮箱1，2]
    public static void main(String[] args) throws Exception {
        final String testEmail = "qiaoyh@qibebt.ac.cn";
        AutoSendEmail t = new AutoSendEmail();
        EmailConfig emailConfig = t.getEmailTec();
        t.getFile();
        t.getEmail(emailConfig);
//        t.depEmail.print();
        t.process(emailConfig);
//        t.process(emailConfig, testEmail);  //测试,加上该参数后，发送地址和抄送地址都是testEmail
    }
    private EmailConfig getEmailSingfle(){
        EmailConfig finance = new EmailConfig();
        finance.setEmailFile("/backup1t/project/excel_depSplit_email/data/singlefile/邮件数据.xlsx");
        finance.setEmail("K");
        finance.setDep("B");
        finance.setSubject("关于2019冬季毕业学生的学位论文评阅劳务费");
        finance.setMail_cc(new String[]{"sufx@qibebt.ac.cn"});
        finance.setPersonName("青能所人事教育处");
        finance.setMailBody(AutoSendEmail.class.getClassLoader().getResource("mail_body_single.md").getPath());
        return finance;
    }
    private EmailConfig getEmailFinance(){
        EmailConfig finance = new EmailConfig();
        finance.setEmailFile("/backup1t/project/excel_depSplit_email/data/lilb/研究组发送切分邮件清单20191206.xlsx");
        finance.setEmail("F");
        finance.setDep("C");
        finance.setSubject("2019年管理费计提工作-请周一结束前返回高茜");
        finance.setMail_cc(new String[]{"lilb@qibebt.ac.cn"});
        finance.setPersonName("青能所财务处");
        finance.setMailBody(AutoSendEmail.class.getClassLoader().getResource("mail_body_finance.md").getPath());
        return finance;
    }
    private EmailConfig getEmailTec(){
        EmailConfig tec = new EmailConfig();
//        tec.setEmailFile(AutoSendEmail.class.getClassLoader().getResource("dep-email.xlsx").getPath());
        tec.setEmailFile("/backup1t/project/excel_depSplit_email/data/gaoqian/email.xlsx");
        tec.setEmail("F");
        tec.setDep("C");
        tec.setMail_cc( new String[]{"gaoqian@qibebt.ac.cn"});
        tec.setSubject("2019年管理费补充（科研经费到款时间为2018年12月和2019年12月部分）计提工作-请周二结束前返回高茜");
        tec.setPersonName("青能所科技处");
        tec.setMailBody(AutoSendEmail.class.getClassLoader().getResource("mail_body.md").getPath());
        return tec;
    }

    public void process(EmailConfig emailConfig, String... testEmail) throws IOException {
//        String resource = AutoSendEmail.class.getClassLoader().getResource(emailConfig.getMailBody()).getPath();
        String mail_body_md= new String(Files.readAllBytes(Paths.get(emailConfig.getMailBody())), StandardCharsets.UTF_8);
        String mail_body = MarkDownUtil.toHtml(mail_body_md);
        MailServer sendmail = new MailServer();
        String[] mail_cc = emailConfig.getMail_cc();
        int flag = 0;
        int test_mailsend_number = 0;
        for (String dep : depEmail.keySet()) {
            List<String> emailList = depEmail.get(dep);
            String[] emailarr =  emailList.toArray(new String[emailList.size()]);
            List<String> fileList = attachFiles.get(dep);
            if (fileList == null  || fileList.size()==0)
                continue;
            else{
                //所有人都发的附件
//                fileList.add("/tmp/财政项目收支结余表2019年.xlsx");
            }
            String[] filearr = fileList.toArray(new String[fileList.size()]);
            if (testEmail.length == 0) {
                sendmail.send(emailarr, mail_body, emailConfig.getSubject(), mail_cc, filearr, emailConfig.getPersonName());
                flag ++;
            }else{ //测试，只发一封，发多了，会被封。单次100人，15分钟400人，（已经单独给我调整到800人）
                mail_cc = testEmail;
                if (test_mailsend_number == 0) {
                    sendmail.send(testEmail, mail_body, emailConfig.getSubject(), mail_cc, filearr,emailConfig.getPersonName());
                    test_mailsend_number++;
                }
            }
            attachFiles.remove(dep); //删除已经发送邮件的
            System.out.println(Arrays.toString(emailarr)+" : "+Arrays.toString(filearr));
        }
        if(!attachFiles.isEmpty()){
            //如果有剩余的附件没有发送
            for (String dep : attachFiles.keySet()) {
                System.out.println(" 没有对应的邮件:"+dep);
            }
            System.out.println(">>> 已经共发送 "+flag+" 封邮件！");
        }else{
            System.out.println(">>> 附件已经全部发送,共发送 "+flag+" 封邮件！");
        }

    }
    private void getEmail(EmailConfig emailConfig) throws Exception {
        String fileEmail = emailConfig.getEmailFile();
        System.out.println(fileEmail);
        LicenseUtil.getLicense();
        Workbook workbook = new Workbook(fileEmail);
        WorksheetCollection sheets = workbook.getWorksheets();
        Worksheet worksheet = sheets.get(0);
        Cells cells = worksheet.getCells();

        for (int r = 0; r <= cells.getMaxRow(); r++) {
            String dep = cells.get(r, CellsHelper.columnNameToIndex( emailConfig.getDep() )).getStringValue();
//            dep = dep.replaceAll("^\\d+-", ""); //把团队前面的数字删除，例如1179-规划战略中心
            dep = DepUtil.filterName(dep);

            String emails = cells.get(r, CellsHelper.columnNameToIndex( emailConfig.getEmail() )).getStringValue();
            String emails2 = emails.replaceAll("；", ";");
            String[] arr = emails2.split(";");
            for (String email : arr) {
                if (email.contains("@")) {
                    depEmail.putOne(dep, email);
                }
            }

        }
    }

    /**
     * 分析文件名，获取部门名称和文件附件的对应关系
     */
    public void getFile(){
        File dir = new File(this.dir);
        if (dir.exists()) {
            File[] filelist = dir.listFiles();
            for (File f : filelist) {
                String dep = getDep( f.getName() );
//                addAttach(dep, f.getName());
                attachFiles.putOne(dep, this.dir+File.separator+f.getName());
            }
        }
    }

    /*private void addAttach(String dep, String filename) {
        if (this.attachFiles.containsKey(dep)) {
            attachFiles.get(dep).add(filename);
        }else{
            List arr = new ArrayList();
            arr.add(filename);
            attachFiles.put(dep, arr);
        }
    }*/
    private String getDep(String filename) {
        String[] s1 = filename.split("_");
        String s2 = s1[s1.length - 1];
        return s2.split("\\.")[0];
    }

}
class MapValueList extends HashMap<String, List<String>>{
    public void putOne(String key, String value) {
        if (this.containsKey(key)) {
            List<String> list = this.get(key);
            //如果已经存在，不添加
            if (!list.contains(value)) {
                this.get(key).add(value);
            }
        }else{
            List<String> arr = new ArrayList<String>();
            arr.add(value);
            this.put(key, arr);
        }
    }

    public void print() {
        for (String key : this.keySet()) {
            List<String> value = this.get(key);
            String[] arr = value.toArray(new String[value.size()]);
            System.out.println(key+ Arrays.toString(arr));
        }
    }
}
