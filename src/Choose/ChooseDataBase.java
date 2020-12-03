package Choose;

import Entity.ComponentsOne;
import Entity.ComponentsSum;
import Entity.ComponentsTwo;
import Entity.Pathname;
import jdbc.PartDB;
import jdbc.PartEntryDB;
import jdbc.PartSumDB;

import java.io.IOException;
import java.sql.SQLException;
import java.util.Scanner;

/**
 * JDBC操作三个表的零件信息功能选择类
 * 事先请建立一个compoents数据库
 * 在compoents数据库中建立三个表:
 * 1.part表
 *     字段为: number(varchar) name(varchar) size(varchar)
 * 2.partentry表
 *     字段为: number(varchar) amount(int)
 * 3.partsum表
 *     字段为: number(varchar) name(varchar) size(varchar) amount(int)
 */
public class ChooseDataBase {
static {
    System.out.println("--------------------------");
    System.out.println("| 欢迎进入数据库增删改查功能! |");
    System.out.println("--------------------------");
}

    public static void dataBase(ComponentsOne one, ComponentsTwo two, ComponentsSum three, Pathname pn) throws IOException, SQLException {
        System.out.println("已进入数据库操作表信息功能!");
        System.out.println("输入你想操作的表:");
        int mark=0;
        while (mark!=1){
            System.out.println("1.零件表  2.零件进库表 3.零件汇总表 4.退出数据库操作功能");
            Scanner sc=new Scanner(System.in);
            int choose=sc.nextInt();
            if (choose==1){
                functionPartDB(one,pn);
            }else if (choose==2){
                functionPartEntryDB(two, pn);
            }else if (choose==3){
                functionPartSumDB(three, pn);
            }else if (choose==4){
                mark=1;
            }else {
                System.out.println("输入操作代号错误,请重新输入:");
            }
        }
    }

    //JDBC操作零件表part的方法
    public static void functionPartDB(ComponentsOne one, Pathname pn) throws SQLException, IOException {
        System.out.println("已连接到compoents数据库,请输入操作part表的代号:");
        int mark=0;
        while (mark!=1){
            System.out.println("1.将零件表.xlsx信息导入到part表  2.插入单条信息 3.查询信息 4.修改信息 5.删除信息 6.退出");
            Scanner sc=new Scanner(System.in);
            int choose=sc.nextInt();
            if (choose==1){
                PartDB.insertAllPart(one);
            }else if (choose==2){
                PartDB.insertPart(one,pn);
            }else if (choose==3){
                PartDB.selectPart(one);
            }else if (choose==4){
                PartDB.updatePart(one,pn);
            }else if (choose==5){
                PartDB.deletePrat(one, pn);
            }else if (choose==6){
                mark=1;
            }else {
                System.out.println("输入代号错误,请重新输入!");
            }
        }
    }

    //JDBC操作零件进库表partentry的方法
    public static void functionPartEntryDB(ComponentsTwo two, Pathname pn) throws SQLException, IOException {
        System.out.println("已连接到compoents数据库,请输入操作partentry表的代号:");
        int mark=0;
        while (mark!=1){
            System.out.println("1.将零件进库表.xlsx信息导入到partentry表  2.插入单条信息 3.查询信息 4.修改信息 5.删除信息 6.退出");
            Scanner sc=new Scanner(System.in);
            int choose=sc.nextInt();
            if (choose==1){
                PartEntryDB.insertAllPartEntry(two);
            }else if (choose==2){
                PartEntryDB.insertPartEntry(two,pn);
            }else if (choose==3){
                PartEntryDB.selectPartEntry(two);
            }else if (choose==4){
                PartEntryDB.updatePartEntry(two, pn);
            }else if (choose==5){
                PartEntryDB.deletePratEntry(two, pn);
            } else if (choose==6){
                mark=1;
            }else {
                System.out.println("输入代号错误,请重新输入!");
            }
        }
    }

    //JDBC操作零件汇总表partsum的方法
    public static void functionPartSumDB(ComponentsSum three, Pathname pn) throws SQLException, IOException {
        System.out.println("已连接到compoents数据库,请输入操作partsum表的代号:");
        int mark=0;
        while (mark!=1){
            System.out.println("1.将零件汇总表.xlsx信息导入到partsum表  2.插入单条信息 3.查询信息 4.修改信息 5.删除信息 6.退出");
            Scanner sc=new Scanner(System.in);
            int choose=sc.nextInt();
            if (choose==1){
                PartSumDB.insertAllPartSum(three);
            }else if (choose==2){
                PartSumDB.insertPartSum(three, pn);
            }else if (choose==3){
                PartSumDB.selectPartSum();
            }else if (choose==4){
                PartSumDB.updatePartSum(three, pn);
            }else if (choose==5){
                PartSumDB.deletePratSum(three, pn);
            }else if (choose==6){
                mark=1;
            }else {
                System.out.println("输入代号错误,请重新输入!");
            }
        }
    }
}
