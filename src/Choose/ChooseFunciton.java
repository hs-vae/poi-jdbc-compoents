package Choose;


import Entity.ComponentsOne;
import Entity.ComponentsSum;
import Entity.ComponentsTwo;
import Entity.Pathname;
import Function.XlsxAdd;
import Function.XlsxChange;
import Function.XlsxDelete;
import Function.XlsxSelect;

import java.io.IOException;
import java.util.Scanner;

/**
 * Poi技术增删改查功能选择类
 */
public class ChooseFunciton {

    static {
        System.out.println("------------------------");
        System.out.println("| 欢迎进入Poi增删改查功能! |");
        System.out.println("------------------------");
    }
    public static void Merit(ComponentsOne one, ComponentsTwo two, ComponentsSum three, Pathname pn) throws IOException {
        System.out.println("已进入Poi操作零件信息功能!");
        Scanner sc=new Scanner(System.in);
        System.out.println("请选择你想进行的功能代号:");
        int mark=0;
        while (mark!=1){
            System.out.println("1.零件表 2.零件进库表 3.零件汇总表 4.退出Poi操作零件信息功能");
            int choose=sc.nextInt();
            if(choose==1){
                XlsxOneMerite(one, pn);
            }else if(choose==2){
                XlsxTwoMerite(two, pn);
            }else if(choose==3){
                XlsxSumMerite(three, pn);
            }else if(choose==4){
                mark=1;
            }else {
                System.out.println("输入代号错误,请重新输入:");
            }
        }
    }

    public static void XlsxOneMerite (ComponentsOne one,Pathname pn) throws IOException {
        int mark=0;
        while (mark!=1){
            System.out.println("1.插入单条信息 2.删除信息 3.修改信息 4.查询信息 5.退出操作零件表");
            Scanner sc=new Scanner(System.in);
            int choose=sc.nextInt();
            if(choose==1){
                XlsxAdd.addXlsxOne(one,pn);
            }else if (choose==2){
                XlsxDelete.deleteXlsxOne(one,pn);
            }else if(choose==3){
                XlsxChange.changeXlsxOne(one, pn);
            }else if (choose==4){
                XlsxSelect.selectXlsxOne(one, pn);
            } else if (choose == 5) {
                mark=1;
            }else {
                System.out.println("输入错误,请重新输入:");
            }
        }
    }
    public static void XlsxTwoMerite(ComponentsTwo two,Pathname pn) throws IOException {
        int mark=0;
        while (mark!=1){
            System.out.println("1.插入单条信息 2.删除信息 3.修改信息 4.查询信息 5.退出操作零件进库表");
            Scanner sc=new Scanner(System.in);
            int choose=sc.nextInt();
            if(choose==1){
                XlsxAdd.addXlsxTwo(two, pn);
            }else if (choose==2){
                XlsxDelete.deleteXlsxTwo(two, pn);
            }else if(choose==3){
                XlsxChange.changeXlsxTwo(two, pn);
            }else if (choose==4){
                XlsxSelect.selectXlsxTwo(two, pn);
            }else if (choose==5){
                mark=1;
            }else {
                System.out.println("输入错误,请重新输入:");
            }
        }
    }
    public static void XlsxSumMerite(ComponentsSum three,Pathname pn) throws IOException {
        int mark=0;
        while (mark!=1){
            System.out.println("1.插入单条信息 2.删除信息 3.修改信息 4.查询信息 5.退出操作零件汇总表");
            Scanner sc=new Scanner(System.in);
            int choose=sc.nextInt();
            if(choose==1){
                XlsxAdd.addXlsxThree(three, pn);
            }else if (choose==2){
                XlsxDelete.deleteXlsxSum(three, pn);
            }else if(choose==3){
                XlsxChange.changeXlsxSum(three, pn);
            }else if (choose==4){
                XlsxSelect.selectXlsxSum(three, pn);
            }else if (choose==5){
                mark=1;
            }else {
                System.out.println("输入错误,请重新输入:");
            }
        }
    }
}
