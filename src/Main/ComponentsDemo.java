package Main;

import Choose.ChooseDataBase;
import Choose.ChooseFunciton;
import Choose.ChooseGenerate;
import Entity.ComponentsOne;
import Entity.ComponentsSum;
import Entity.ComponentsTwo;
import Entity.Pathname;

import java.io.IOException;
import java.sql.SQLException;
import java.util.Scanner;

/**
 * 零件库房系统主类
 */
public class ComponentsDemo {
    public static void main(String[] args) throws IOException, SQLException {
        //创建零件表信息类对象
        ComponentsOne cp1=new ComponentsOne();
        //创建零件进库表信息类对象
        ComponentsTwo cp2=new ComponentsTwo();
        //创建零件表汇总表类对象
        ComponentsSum cp3=new ComponentsSum();
        //创建三个表文件路径对象
        Pathname pn=new Pathname();
        //生成汇总表
        ChooseGenerate.generateSum(cp1,cp2,cp3,pn);
        /*
            1.利用poi技术增删改查功能
            2.利用JDBC将三个表零件信息导入数据库对应的表中,操作数据库同时将零件信息写入到对应的文件中
         */
        choose(cp1,cp2,cp3,pn);
    }

    public static void choose(ComponentsOne one,ComponentsTwo two,ComponentsSum three,Pathname pn) throws IOException, SQLException {
        int mark=0;
        while (mark!=1){
            System.out.println("1.Poi技术操作三个表的零件信息  2.数据库JDBC驱动操作三个表的零件信息  3.退出零件库房系统");
            Scanner sc=new Scanner(System.in);
            int choose=sc.nextInt();
            if (choose==1){
                //利用poi技术增删改查功能
                ChooseFunciton.Merit(one,two,three,pn);
            }else if(choose==2){
                //利用JDBC将三个表零件信息导入数据库对应的表中,操作数据库同时将零件信息写入到对应的文件中
                ChooseDataBase.dataBase(one, two, three, pn);
            }else if(choose==3){
                System.out.println("已退出零件库房系统");
                mark=1;
            }else {
                System.out.println("输入代号错误,请重新输入:");
            }
        }
    }
}
