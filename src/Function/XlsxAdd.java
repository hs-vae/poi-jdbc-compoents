package Function;

import Entity.ComponentsOne;
import Entity.ComponentsSum;
import Entity.ComponentsTwo;
import Entity.Pathname;
import Generate.XlsxOne;
import Generate.XlsxSum;
import Generate.XlsxTwo;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.*;
import java.util.Arrays;
import java.util.Scanner;

class G implements Comparable<G>{ //实现comparable接口,复写comparato方法
    String number;
    String name;
    String size;
    int amount;
    @Override
    public int compareTo(G g) {
        if (this.number.compareTo(g.number) <0 )
            return 1;
        else
            return -1;
    }
}
/**
 *  添加功能类:
 *  添加一行零件(零件表,零件进库表,零件汇总表)信息
 */
public class XlsxAdd {

    //零件表末尾添加一行信息的方法
    public  static void addXlsxOne(ComponentsOne one, Pathname pn) throws IOException {
        System.out.println("------零件表(添加前)------");
        XlsxOne.printXlsxOne(pn);
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("零件表");
        FileOutputStream fos=new FileOutputStream(new File(pn.getPathnameOne()));
        Workbook sheets=null;
        XSSFRow row=null;
        XSSFCell cell=null;
        Scanner sc=new Scanner(System.in);
        //增加零件表每个零件信息数组的长度+1
        one.number = Arrays.copyOf(one.number,one.number.length+1);
        one.name = Arrays.copyOf(one.name, one.name.length + 1);
        one.size = Arrays.copyOf(one.size,one.size.length+1);
        System.out.println("依次输入需要添加的零件编号,零件名称,零件规格");
        String addNumber=sc.next();
        String addName=sc.next();
        String addSize=sc.next();
        for (int i = 0; i < one.number.length+1; i++) { //4+1=5
            row = sheet.createRow(i);
            cell = row.createCell(i);
            if (i == 0) {
                //设置单元格范围地址,参数分别表示为(起始行号,终止行号,起始列号,终止列号)
                CellRangeAddress region = new CellRangeAddress(0, 0, 0, 2);
                row.createCell(0).setCellValue("零件表信息");
                //设置水平居中
                CellStyle style = workbook.createCellStyle();
                style.setAlignment(HorizontalAlignment.CENTER);
                cell.setCellStyle(style);
                sheet.addMergedRegion(region);
            }  else {
                //在数组尾部添加新增的零件信息
                if(i==one.number.length){
                    row.createCell(0).setCellValue(addNumber);
                    row.createCell(1).setCellValue(addName);
                    row.createCell(2).setCellValue(addSize);
                    one.number[i-1]=addNumber;
                    one.name[i-1]=addName;
                    one.size[i-1]=addSize;
                    one.setNumber(one.number);
                    one.setName(one.name);
                    one.setSize(one.size);
                }else {
                    row.createCell(0).setCellValue(one.getNumber()[i-1]);
                    row.createCell(1).setCellValue(one.getName()[i-1]);
                    row.createCell(2).setCellValue(one.getSize()[i-1]);
                }
            }
        }
        System.out.println("添加成功,已存入" + pn.getPathnameOne() + "目录中!");
        workbook.write(fos);
        workbook.close();
        fos.close();
        //打印上面输入的零件表信息
        System.out.println("------零件表(添加后)------");
        XlsxOne.printXlsxOne(pn);
    }


   //零件进库表末尾添加一行信息的方法
    public  static void addXlsxTwo(ComponentsTwo two, Pathname pn) throws IOException {
        System.out.println("------零件进库表(添加前)------");
        XlsxTwo.printXlsxTwo(pn);
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("零件进库表");
        FileOutputStream fos=new FileOutputStream(new File(pn.getPathnameTwo()));
        Workbook sheets=null;
        XSSFRow row=null;
        XSSFCell cell=null;
        Scanner sc=new Scanner(System.in);
        //增加零件表每个零件信息数组的长度+1
        two.number = Arrays.copyOf(two.number,two.number.length+1);
        two.amount = Arrays.copyOf(two.amount,two.amount.length+1);
        System.out.println(two.number.length);
        System.out.println("依次输入需要添加的零件编号,零件数量");
        String addNumber=sc.next();
        int addAmount=sc.nextInt();
        for (int i = 0; i < two.number.length+1; i++) { //4+1=5
            row = sheet.createRow(i);
            cell = row.createCell(i);
            if (i == 0) {
                //设置单元格范围地址,参数分别表示为(起始行号,终止行号,起始列号,终止列号)
                CellRangeAddress region = new CellRangeAddress(0, 0, 0, 1);
                row.createCell(0).setCellValue("零件进库表信息");
                //设置水平居中
                CellStyle style = workbook.createCellStyle();
                style.setAlignment(HorizontalAlignment.CENTER);
                cell.setCellStyle(style);
                sheet.addMergedRegion(region);
            }  else {
                //在数组尾部添加新增的零件信息
                if(i==two.number.length){
                    row.createCell(0).setCellValue(addNumber);
                    row.createCell(1).setCellValue(addAmount);
                    two.number[i-1]=addNumber;
                    two.amount[i-1]=addAmount;
                    two.setNumber(two.number);
                    two.setAmount(two.amount);
                }else if(i==1){
                    row.createCell(0).setCellValue(two.getNumber()[i-1]);
                    row.createCell(1).setCellValue("零件数量");
                }else{
                    row.createCell(0).setCellValue(two.getNumber()[i-1]);
                    row.createCell(1).setCellValue(two.getAmount()[i-1]);
                }
            }
        }
        System.out.println("添加成功,已存入" + pn.getPathnameTwo() + "目录中!");
        workbook.write(fos);
        workbook.close();
        fos.close();

        //打印上面输入的零件表信息
        System.out.println("------零件进库表(添加后)------");
        XlsxTwo.printXlsxTwo(pn);
    }

    //零件汇总表末尾添加一行信息的方法,并且排序好编号
    public static void addXlsxThree(ComponentsSum three, Pathname pn) throws IOException {
        System.out.println("------零件汇总表(添加前)------");
        XlsxSum.PrintSumXlsx(pn);
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("零件汇总表");
        FileOutputStream fos=new FileOutputStream(new File(pn.getPathnameSum()));
        Workbook sheets=null;
        XSSFRow row=null;
        XSSFCell cell=null;
        Scanner sc=new Scanner(System.in);
        //增加零件表每个零件信息数组的长度+1
        three.number = Arrays.copyOf(three.number,three.number.length+1);
        three.name = Arrays.copyOf(three.name,three.name.length+1);
        three.size = Arrays.copyOf(three.size,three.size.length+1);
        three.amount = Arrays.copyOf(three.amount,three.amount.length+1);
        System.out.println("依次输入需要添加的零件编号,零件名称,零件规格,零件数量");
        String addNumber=sc.next();
        String addName=sc.next();
        String addSize=sc.next();
        int addAmount=sc.nextInt();

        //将汇总表的零件信息全部存储在数组g里面
        G g [] = new G[three.number.length-1];
        for (int i = 0; i < g.length; i++) {   //g.length==6
            g[i] = new G();
            g[i].number= three.getNumber()[i+1];
            g[i].name = three.getName()[i+1];
            g[i].size= three.getSize()[i+1];
            g[i].amount = three.getAmount()[i+1];
            if(i==(g.length-1)){
                g[i].number = addNumber;
                g[i].name = addName;
                g[i].size = addSize;
                g[i].amount = addAmount;
            }
        }
        //从大到小对编号进行排序
        Arrays.sort(g);
        //想得到从小到大的编号,所以倒置循环一下,将零件信息又存储在three对象的成员变量数组中
        int e=1;
        for (int i = g.length; i >= 1; i--) {
            three.number[e]=g[i-1].number;
            three.name[e]=g[i-1].name;
            three.size[e]=g[i-1].size;
            three.amount[e]=g[i-1].amount;
            ++e;
        }
        for (int i = 0; i < three.number.length+1; i++) {
            row = sheet.createRow(i);
            cell = row.createCell(i);
            if (i == 0) {
                //设置单元格范围地址,参数分别表示为(起始行号,终止行号,起始列号,终止列号)
                CellRangeAddress region = new CellRangeAddress(0, 0, 0, 3);
                row.createCell(0).setCellValue("零件汇总表信息");
                //设置水平居中
                CellStyle style = workbook.createCellStyle();
                style.setAlignment(HorizontalAlignment.CENTER);
                cell.setCellStyle(style);
                sheet.addMergedRegion(region);
                three.number[0]="零件编号";
                three.name[0]="零件名称";
                three.size[0]="零件规格";
                three.amount[0]=0;
            }  else {
                three.setNumber(three.number);
                three.setName(three.name);
                three.setSize(three.size);
                three.setAmount(three.amount);
                //在数组尾部添加新增的零件信息
                row.createCell(0).setCellValue(three.getNumber()[i-1]);
                row.createCell(1).setCellValue(three.getName()[i-1]);
                row.createCell(2).setCellValue(three.getSize()[i-1]);
                if(i==1){
                    row.createCell(3).setCellValue("零件数量");
                }else{
                    row.createCell(3).setCellValue(three.getAmount()[i-1]);
                }
            }
        }
        System.out.println("添加成功,已存入" + pn.getPathnameSum() + "目录中!");
        workbook.write(fos);
        workbook.close();
        fos.close();

        //打印上面输入的零件表信息
        System.out.println("------零件进库表(添加后)------");
        XlsxSum.PrintSumXlsx(pn);
    }
}
