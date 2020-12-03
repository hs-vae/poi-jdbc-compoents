package Generate;

import Entity.ComponentsSum;
import Entity.Pathname;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Arrays;
import java.util.Scanner;

/*
    根据零件表和零件进库表的编号判断是否相同,
    抽取相同的零件信息并且排列编号为汇总表
 */

class E implements Comparable<E>{ //实现comparable接口，复写comparato方法
    String number;
    String name;
    String size;
    @Override
    public int compareTo(E e) {
        if (this.number.compareTo(e.number) <0 )
            return 1;
        else
            return -1;
    }
}
class H implements Comparable<H>{ //实现comparable接口，复写comparato方法
    String number;
    Double amount;
    @Override
    public int compareTo(H d) {
        if (this.number.compareTo(d.number) <0 )
            return 1;
        else
            return -1;
    }
}

/**
 * 零件汇总表的生成类
 */
public class XlsxSum {
    public  static void creatSumXlsx1(ComponentsSum three, Pathname pn) throws IOException {
        //获取两个表的信息
        FileInputStream fis1 = new FileInputStream(pn.getPathnameOne());
        FileInputStream fis2 = new FileInputStream(pn.getPathnameTwo());
        Workbook sheetsOne=new XSSFWorkbook(fis1);
        Sheet sheet1=sheetsOne.getSheetAt(0);
        Workbook sheetsTwo=new XSSFWorkbook(fis2);
        Sheet sheet2=sheetsTwo.getSheetAt(0);
        //产生汇总表
        XSSFWorkbook workbook=new XSSFWorkbook();
        System.out.println("已自动生成汇总表!");
        XSSFSheet sheet=workbook.createSheet("零件汇总表");
        System.out.println("输入零件汇总表.xlsx生成的位置:");
        Scanner sc=new Scanner(System.in);
        String pathnameThree=sc.nextLine();
        pn.setPathnameSum(pathnameThree);
        FileOutputStream fos=new FileOutputStream(pn.getPathnameSum());
        XSSFRow row=null;
        XSSFCell cell=null;
        //用String类型数组存放输入的零件编号,名字,规格
        int num1=sheet1.getLastRowNum()-1;
        int num2=sheet2.getLastRowNum()-1;
        int a=Math.min(num1,num2);
        //定义一个arr数组用来存储零件表.xlsx所有零件信息
        String[] numberOne=new String[num1]; //编号
        String[] name=new String[num1];
        String[] size=new String[num1];
        //定义一个brr数组用来存储零件进库表.xlsx所有零件信息
        String[] numberTwo=new String[num2]; //编号
        double[] amount=new double[num2];
        //定义零件汇总表对象的成员变量
        three.number=new String[a+1];
        three.name=new String[a+1];
        three.size=new String[a+1];
        three.amount=new int[a+1];
        //先将零件表.xlsx按编号排序好零件信息,并存储在对应的数组里
        for (int i = 2; i < num1+2; i++) {
            Row row1=sheet1.getRow(i);
            numberOne[i-2]=row1.getCell(0).getStringCellValue();
            name[i-2]=row1.getCell(1).getStringCellValue();
            size[i-2]=row1.getCell(2).getStringCellValue();
        }
        //先将零件进库表.xlsx按编号排序好零件信息,并存储在对应的数组里
        for (int i = 2; i < num2+2; i++) {
            Row row2=sheet2.getRow(i);
            numberTwo[i-2]=row2.getCell(0).getStringCellValue();
            amount[i-2]=row2.getCell(1).getNumericCellValue();
        }

        E e [] = new E[num1];
        for (int i = 0; i < num1; i++) {
            e[i] = new E();
            e[i].number = numberOne[i];
            e[i].name = name[i];
            e[i].size = size[i];
        }
        Arrays.sort(e);//从大到小排序
        int e1=0;
        for (int i = num1-1; i >=0; i--) {
            numberOne[e1]=e[i].number;
            name[e1]=e[i].name;
            size[e1]=e[i].size;
            e1++;
        }

        H h [] = new H[num2];
        for (int i = 0; i < num2; i++) {
            h[i] = new H();
            h[i].number= numberTwo[i];
            h[i].amount = amount[i];
        }
        Arrays.sort(h);
        int h1=0;
        for (int i = num2-1; i >=0; i--) {
            numberTwo[h1]=h[i].number;
            amount[h1]=h[i].amount;
            h1++;
        }

        //设置零件汇总表.xlsx标题和表头
        for (int i = 0; i < 2; i++) {
            row = sheet.createRow(i);
            cell = row.createCell(i);
            if(i==0){
                //设置标题
                //设置单元格范围地址,参数分别表示为(起始行号,终止行号,起始列号,终止列号)
                CellRangeAddress region = new CellRangeAddress(0, 0, 0, 3);
                row.createCell(0).setCellValue("零件汇总表信息");
                //设置水平居中
                CellStyle style = workbook.createCellStyle();
                style.setAlignment(HorizontalAlignment.CENTER);
                cell.setCellStyle(style);
                sheet.addMergedRegion(region);
            }else{
                //设置表头:编号,名称,规格,进库数量
                row.createCell(0).setCellValue("零件编号");
                row.createCell(1).setCellValue("零件名称");
                row.createCell(2).setCellValue("零件规格");
                row.createCell(3).setCellValue("零件数量");
                three.number[0]="零件编号";
                three.name[0]="零件名称";
                three.size[0]="零件规格";
                three.amount[0]=0;
                three.setNumber(three.number);
                three.setName(three.name);
                three.setSize(three.size);
                three.setAmount(three.amount);
            }
        }
        //采用归并思想,判断两个(排列好编号)表的编号是否相同,若相同则放入汇总表里
        int i=0,j=0,d1=0,d2=0;
        int k=2;
        while(i<numberOne.length && j<numberTwo.length){
            if(numberOne[i].equals(numberTwo[j])){
                ++d2;
                row=sheet.createRow(k);
                row.createCell(0).setCellValue(numberOne[i]);
                row.createCell(1).setCellValue(name[i]);
                row.createCell(2).setCellValue(size[i]);
                row.createCell(3).setCellValue(amount[j]);
                three.number[d2]=numberOne[d1];
                three.name[d2]=name[d1];
                three.size[d2]=size[d1];
                three.amount[d2]=(int)amount[d1];
                three.setNumber(three.number);
                three.setName(three.name);
                three.setSize(three.size);
                three.setAmount(three.amount);
                ++d1;
                k++;
                i++;
                j++;
            }else if((numberOne[i].compareTo(numberTwo[j]))==-1){
                i++;
            }else{
                j++;
            }
        }
        workbook.write(fos);
        workbook.close();
        fos.close();
        PrintSumXlsx(pn);
    }

    //从外面导入汇总表
    public static void creatSumXlsx2(ComponentsSum three,Pathname pn) throws IOException {
        System.out.println("请输入要导入的零件汇总表路径:");
        Scanner sc=new Scanner(System.in);
        String pathname = sc.next();
        pn.setPathnameSum(pathname);
        FileInputStream fis = new FileInputStream(pn.getPathnameSum());
        Workbook sheets=new XSSFWorkbook(fis);
        Sheet sheet1 = sheets.getSheetAt(0);
        three.number=new String[sheet1.getLastRowNum()];
        three.name=new String[sheet1.getLastRowNum()];
        three.size=new String[sheet1.getLastRowNum()];
        three.amount=new int[sheet1.getLastRowNum()];
        System.out.println("------零件汇总表------");
        for (int i = 1; i < sheet1.getLastRowNum() + 1; i++) {
            //获取行数
            Row row1 = sheet1.getRow(i);
            //获取单元值并且取值
            three.number[i-1]=row1.getCell(0).getStringCellValue();
            three.name[i-1]=row1.getCell(1).getStringCellValue();
            three.size[i-1]=row1.getCell(2).getStringCellValue();
            if(i==1){
                three.amount[i-1]=0;
                System.out.println(row1.getCell(0).getStringCellValue()
                        + "   " + row1.getCell(1).getStringCellValue()
                        + "   " + row1.getCell(2).getStringCellValue()
                        + "   " + row1.getCell(3).getStringCellValue());
            }else{
                double c=row1.getCell(3).getNumericCellValue();
                System.out.println(row1.getCell(0).getStringCellValue()
                        + "   " + row1.getCell(1).getStringCellValue()
                        + "   " + row1.getCell(2).getStringCellValue()
                        + "   " + (int)c);
                three.amount[i-1]=(int)c;
            }
        }
        System.out.println("-----------------");
        fis.close();
        sheets.close();
    }



    //打印输出汇总表
    public  static void PrintSumXlsx(Pathname pn) throws IOException {
        FileInputStream fis=new FileInputStream(pn.getPathnameSum());
        Workbook sheets=new XSSFWorkbook(fis);
        Sheet sheet=sheets.getSheetAt(0);
        System.out.println("--------零件汇总表(排序好编号后)----------");
        for (int i = 1; i < sheet.getLastRowNum()+1;i++){
            //获取行数
            Row row=sheet.getRow(i);
            //获取单元值并且取值
            if(i==1){
                System.out.println(row.getCell(0).getStringCellValue()
                        +"   "+row.getCell(1).getStringCellValue()
                        +"   "+row.getCell(2).getStringCellValue()
                        +"   "+row.getCell(3).getStringCellValue());
            }else{
                double c= row.getCell(3).getNumericCellValue();
                System.out.println(row.getCell(0).getStringCellValue()
                        +"   "+row.getCell(1).getStringCellValue()
                        +"   "+row.getCell(2).getStringCellValue()
                        +"   "+(int)c);
            }
        }
        System.out.println("--------------------------");
        sheets.close();
        fis.close();
    }
}
