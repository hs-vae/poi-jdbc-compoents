package Choose;


import Entity.ComponentsOne;
import Entity.ComponentsSum;
import Entity.ComponentsTwo;
import Entity.Pathname;
import Generate.XlsxOne;
import Generate.XlsxSum;
import Generate.XlsxTwo;

import java.io.IOException;
import java.util.Scanner;

/**
 * 零件表,零件进库表,零件汇总表的生成方式选择类
 */
public class ChooseGenerate {
    static {
        System.out.println("------------------------");
        System.out.println("|  欢迎进入零件库房系统!   |");
        System.out.println("------------------------");
    }
    public static void generateSum(ComponentsOne one, ComponentsTwo two, ComponentsSum three , Pathname pn) throws IOException {
        Scanner sc=new Scanner(System.in);
        int mark=0;
        while (mark!=1){
            System.out.println("输入你想生成零件汇总表的方式代号:");
            System.out.println("1.通过创建零件表和零件进库表生成  2.通过已经准备好的零件汇总表导入  3.生成完毕,退出生成功能");
            int choose=sc.nextInt();
            if(choose==1){
                creatComponents(one,pn);
                creatComponentsEntry(two,pn);
                XlsxSum.creatSumXlsx1(three,pn);
                System.out.println("请输入下一步操作代号:");
            }else if(choose==2){
                XlsxSum.creatSumXlsx2(three,pn);
                System.out.println("为了后续的功能请导入零件表和零件进库表");
                XlsxOne.creatXlsx2(one,pn);
                XlsxTwo.creatXlsx2(two,pn);
                System.out.println("请输入下一步操作代号:");
            }else if(choose==3){
                System.out.println("已退出生成零件汇总表功能!");
                mark=1;
            }else {
                System.out.println("输入生成方式代号错误,请重新输入:");
            }
        }
    }


    public static void creatComponents(ComponentsOne one, Pathname pn){
        Scanner sc=new Scanner(System.in);
        int mark=0;
        while (mark!=1){
            System.out.println("输入你想生成零件表的方式代号:");
            System.out.println("1.手动输入零件表信息  2.从外界导入零件表信息 ");
            int choose=sc.nextInt();
            if(choose==1){
                XlsxOne.creatXlsx1(one,pn);
                mark=1;
            }else if(choose==2){
                XlsxOne.creatXlsx2(one,pn);
                mark=1;
            }else{
                System.out.println("输入生成代号错误,请重新输入!");
            }
        }
    }
    public static void creatComponentsEntry(ComponentsTwo two,Pathname pn){
        Scanner sc=new Scanner(System.in);
        int mark=0;
        while (mark!=1){
            System.out.println("输入你想生成零件进库表的方式代号:");
            System.out.println("1.手动输入零件进库表信息  2.从外界导入零件进库表信息");
            int choose=sc.nextInt();
            if (choose==1){
                XlsxTwo.creatXlsx1(two, pn);
                mark=1;
            }else if(choose==2){
                XlsxTwo.creatXlsx2(two, pn);
                mark=1;
            }else {
                System.out.println("输入生成方式代号错误,请重新输入:");
            }
        }
    }
}
