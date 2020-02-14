package cn.ac.qibebt.gaoqian;

public class DepUtil {
    public static String filterName(String depName) {
        String dep = depName.replaceAll("^\\d+-", ""); //把团队前面的数字删除，例如1179-规划战略中心
        dep = dep.replaceAll("团队$", "");
        dep = dep.replaceAll("研究组$", "");
        return dep;
    }

    public static void main(String[] args) {
        String name = "单细胞中心研究组团队";
        System.out.println(filterName(name));
    }
}
