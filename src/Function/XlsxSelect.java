package Function;


import Entity.ComponentsOne;
import Entity.ComponentsSum;
import Entity.ComponentsTwo;
import Entity.Pathname;
import Generate.XlsxOne;
import Generate.XlsxSum;
import Generate.XlsxTwo;

import java.io.*;
import java.util.Scanner;

/**
 * 查询功能:(零件表,零件进库表,零件汇总表)根据零件编号查询该对应的零件信息
 */
public class XlsxSelect {

    //根据零件编号查询零件表对应的零件信息的方法
    public static void selectXlsxOne(ComponentsOne one, Pathname pn) throws IOException {
        System.out.println("------零件表------");
        XlsxOne.printXlsxOne(pn);
        Scanner sc = new Scanner(System.in);
        System.out.println("请输入需要查询零件表的零件编号:");
        String idSelect = sc.next();
        for (int i = 0; i < one.number.length; i++) {
            if (idSelect.equals(one.getNumber()[i])) {
                System.out.println("零件编号:" + one.getNumber()[i]);
                System.out.println("零件名称:" + one.getName()[i]);
                System.out.println("零件规格:" + one.getSize()[i]);
            }
        }
    }

    //根据零件编号查询零件进库表对应的零件信息的方法
    public static void selectXlsxTwo(ComponentsTwo two, Pathname pn) throws IOException {
        System.out.println("------零件进库表------");
        XlsxTwo.printXlsxTwo(pn);
        Scanner sc = new Scanner(System.in);
        System.out.println("请输入需要查询零件进库表的零件编号:");
        String idSelect = sc.next();
        for (int i = 0; i < two.number.length; i++) {
            if (idSelect.equals(two.getNumber()[i])) {
                System.out.println("零件编号:" + two.getNumber()[i]);
                System.out.println("零件数量:" + two.getAmount()[i]);
            }
        }
    }

    //根据零件编号查询零件汇总表对应的零件信息的方法
    public static void selectXlsxSum(ComponentsSum three, Pathname pn) throws IOException {
        System.out.println("------零件汇总表------");
        XlsxSum.PrintSumXlsx(pn);
        Scanner sc = new Scanner(System.in);
        System.out.println("请输入需要查询零件汇总表的零件编号:");
        String idSelect = sc.next();
        for (int i = 0; i < three.number.length; i++) {
            if (idSelect.equals(three.getNumber()[i])) {
                System.out.println("零件编号:" + three.getNumber()[i]);
                System.out.println("零件名称:" + three.getName()[i]);
                System.out.println("零件规格:" + three.getSize()[i]);
                System.out.println("零件数量:" + three.getAmount()[i]);
            }
        }
    }
}
