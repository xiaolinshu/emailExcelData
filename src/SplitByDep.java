package cn.ac.qibebt.gaoqian;

import cn.ac.qibebt.util.LicenseUtil;
import com.aspose.cells.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by luke on 12/27/17.
 * excel文件的要求：第一行必须是标题行，且不能合并单元格。部门的标题是关键词，必须一致。
 */
public class SplitByDep {
    private String splitFlag = "研究组";
    private String dirpath = "/backup1t/project/excel_depSplit_email/data/gaoqian/";
    private String xlsfile = "表1-2018年12月和2019年12月科研经费到位尚未计提管理费清单.xlsx";
    private int sheetIndex = 0;
    private String sheetname = null;
    private int depIndex = -1;
    private List<Integer> nosupport = new ArrayList<Integer>();
    private Map<String,List<Integer>> map = new HashMap<String, List<Integer>>();

    // 根据分割标志，查找所在的列，只在sheet的第一行查找
    private void findDep(Cells cells){
        for(int i=0; i<=cells.getMaxColumn();i++) {
            Cell cell = cells.get(0, i);
            if (cell.getValue() == null) {
                continue;
            }
            String value = cell.getValue().toString();
            if (value.contains(this.splitFlag)) {
                this.depIndex = i;
            }
        }
        if (this.depIndex == -1) {
            throw new RuntimeException("没有部门列");
        }
    }
    //把分割结果保存到map 中，key：部门名称，value：[行号列表]
    private void split(Cells cells){
        int log_number = 0; // 记录数据量大小
        for(int r=1;r<=cells.getMaxRow(); r++) {
            Object depo = cells.get(r, this.depIndex).getValue();
            if (depo == null) {
                this.nosupport.add(r);
                continue;
            }
            String dep = depo.toString();
//            dep = dep.replaceAll("^\\d+-", ""); //把团队前面的数字删除，例如1179-规划战略中心
            dep = DepUtil.filterName(dep);

            if(dep.trim().equals("#N/A") || dep.trim().equals("")){
                this.nosupport.add(r);
                continue;
            }
            log_number++;
            List<Integer> list = map.get(dep);
            if (list == null) {
                list = new ArrayList<Integer>();
                list.add(r);
                this.map.put(dep, list);
            }else{
                list.add(r);
            }
        }
        System.out.println("文件："+this.sheetname+"  行数："+log_number);
    }
    //根据map的key，生成一个excel文件
    private void createFile(Cells cells) throws Exception {
//        File dir = new File(this.dirpath + this.xlsfile.split("\\.")[0]);
        File dir = new File("/tmp/11");
        if (!dir.exists()) {
            dir.mkdir();//以excel的文件名创建 文件夹
        }
        for (String dep : this.map.keySet()) {
            List<Integer> list = this.map.get(dep); //dep 对应的行号列表
            String fileName = dir + File.separator + this.xlsfile.split("\\.xls")[0]+"_"+dep + ".xls";
            File file = new File(fileName);
            Workbook workbook = null;
            // 创建第二个sheet时，该文件已经创建。如果文件已经存在，只需读取。
            if (file.exists()) {
                workbook = new Workbook(fileName);
            } else {
                workbook = new Workbook();
            }
            //新创建的excel文件，默认只有一个sheet，当大于1个是，先创建。 用for循环解决大于3个sheet时，第二个sheet为空问题
            for (int k = workbook.getWorksheets().getCount(); k <= this.sheetIndex; k++) {
                workbook.getWorksheets().add();
            }
            /*if (this.sheetIndex > 0) {
                workbook.getWorksheets().add();
            }*/
            Worksheet  sheet = workbook.getWorksheets().get(this.sheetIndex);

            sheet.setName(this.sheetname);
            int n = 0; //新创建的excel 对应的行号
            //写入标题行
            for(int c = 0; c<=cells.getMaxColumn();c++) {
                Cell srcCell = cells.get(0,c);
                Cell desCell = sheet.getCells().get(n, c);
                copyCell(srcCell, desCell);
//                sheet.getCells().get(n, c).setValue(cells.get(0,c).getValue());
                //设置列宽度
                sheet.getCells().setColumnWidthPixel(c, cells.getColumnWidthPixel(c));

            }
            n++; //写入一行，自动加1
            //写入行号列表对应的数据
            for (int r : list) {
                for(int c = 0; c<=cells.getMaxColumn();c++) {
                    Cell srcCell = cells.get(r,c);
                    Cell desCell = sheet.getCells().get(n, c);
                    copyCell(srcCell, desCell);

                    // 日期会变成字符串
//                     sheet.getCells().get(n, c).setValue(cells.get(r,c).getValue());
                }
                n++;
            }
            //自动调整列宽度和 行宽度，必须放在最后，生成数据之后。
            sheet.autoFitColumns();
            sheet.autoFitRows();
            //保存
            workbook.save(fileName);
        }
    }

    private void copyCell(Cell src, Cell des) {
        if (src.isFormula()) {
//            des.setValue(src.getValue());// 日期会变成字符串
            des.copy(src); //cell复制，格式、内容保持原样。
        }else{
            des.copy(src); //cell复制，格式、内容保持原样。
        }
        //设置单元格格式
/*//      wrap text
        Style style = src.getStyle();
        style.setTextWrapped(true);
        StyleFlag flag = new StyleFlag();
        flag.setWrapText(true);
        des.setStyle(style, flag);*/

    }
    public void printNoSupport(Cells cells){
        System.out.println("不支持的行数" + this.nosupport.size());
        for (int n : this.nosupport) {
            System.out.print(n+1+"、");
        }
    }
    public void process() throws Exception {
        LicenseUtil.getLicense();
        Workbook workbook = new Workbook(dirpath+xlsfile);
        WorksheetCollection sheets = workbook.getWorksheets();
        for (Object sheeto : sheets) {
            Worksheet sheet = (Worksheet)sheeto;
            //重置参数,避免多个sheet时出错
            this.sheetIndex = sheet.getIndex();
            if(sheet.isVisible()==false){
                continue;//默认不处理隐藏sheet
            }
            this.depIndex = -1;
            this.sheetname = sheet.getName();
            System.out.println("===Sheet:"+this.sheetname);
            this.map = new HashMap<String, List<Integer>>();
            this.nosupport = new ArrayList<Integer>();
            Cells cells = sheet.getCells();
            findDep(cells);
            if (this.depIndex != -1) {
                split(cells);
                createFile(cells);
                printNoSupport(cells);
            }
        }
//        Worksheet sheet = sheets.get(this.sheetIndex);

    }



    public static void main(String[] args) throws Exception {
        SplitByDep test = new SplitByDep();
        test.process();

    }

}
