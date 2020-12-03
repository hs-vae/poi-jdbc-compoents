package Generate;

import Entity.ComponentsTwo;
import Entity.Pathname;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.text.NumberFormat;
import java.util.Scanner;

/**
 * 零件进库表的生成类
 */
public class XlsxTwo {

    //生成手动输入零件信息的零件进库表.xlsx
    public  static void creatXlsx1(ComponentsTwo cp, Pathname pn)  {
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("零件进库表");
        FileOutputStream fos= null;
        FileInputStream fis=null;
        Workbook sheets=null;
        try {
            Scanner sn = new Scanner(System.in);
            System.out.println("输入零件进库表.xlsx文件生成的位置:");
            String pathname = sn.nextLine();
            pn.setPathnameTwo(pathname);
            fos = new FileOutputStream(pn.getPathnameTwo());
            System.out.println("输入零件进库表的行数:");
            cp.setNumTwo(sn.nextInt());
            XSSFRow row = null;
            XSSFCell cell = null;
            Scanner sc = new Scanner(System.in);
            cp.number=new String[cp.getNumTwo()+1];
            cp.amount=new int[cp.getNumTwo()+1];
            //用String类型数组存放输入的零件编号,名字,规格
            for (int i = 0; i < cp.getNumTwo()+1; i++) {
                if (i == 0) {
                    cp.number[0]= "零件编号";
                    cp.amount[0]=0;
                } else {
                    System.out.println("请输入" + i  +"行零件编号:");
                    cp.number[i]= sc.next();
                    System.out.println("请输入" + i  + "行零件数量:");
                    cp.amount[i] = sc.nextInt();
                }
            }
            cp.setNumber(cp.number);
            cp.setAmount(cp.amount);
            for (int i = 0; i < cp.getNumTwo()+2; i++) {
                if (i == 0) {
                    row = sheet.createRow(i);
                    cell = row.createCell(i);
                    //设置单元格范围地址,参数分别表示为(起始行号,终止行号,起始列号,终止列号)
                    CellRangeAddress region = new CellRangeAddress(0, 0, 0, 1);
                    row.createCell(0).setCellValue("零件进库表信息");
                    //设置水平居中
                    CellStyle style = workbook.createCellStyle();
                    style.setAlignment(HorizontalAlignment.CENTER);
                    cell.setCellStyle(style);
                    sheet.addMergedRegion(region);
                } else {
                    row = sheet.createRow(i);
                    cell = row.createCell(i);
                    row.createCell(0).setCellValue(cp.getNumber()[i-1]);
                    if(i==1){
                        row.createCell(1).setCellValue("零件数量");
                    }else {
                        row.createCell(1).setCellValue(cp.getAmount()[i-1]);
                        //由于Excel设置数字默认类型为浮点型,需要将小数点.0去掉
                        double a=row.getCell(1).getNumericCellValue();
                        NumberFormat nf=NumberFormat.getInstance();
                        String s=nf.format(a);
                        int c=Integer.parseInt(s);
                        row.createCell(1).setCellValue(c);
                    }
                }
            }
            System.out.println("生成成功,已存入" + pathname + "目录中!");
            System.out.println("------零件进库表------");
            workbook.write(fos);
            workbook.close();
            fos.close();

            //打印上面输入的零件表信息
            fis = new FileInputStream(pathname);
            sheets = new XSSFWorkbook(fis);
            Sheet sheet1 = sheets.getSheetAt(0);
            for (int i = 1; i < sheet1.getLastRowNum() + 1; i++) {
                //获取行数
                Row row1 = sheet1.getRow(i);
                //获取单元值并且取值
                if(i==1){
                    System.out.println(row1.getCell(0).getStringCellValue()
                            + "   " + row1.getCell(1).getStringCellValue());
                }else {
                    double a=row1.getCell(1).getNumericCellValue();
                    System.out.println(row1.getCell(0).getStringCellValue()
                            + "   " + (int)a);
                }
            }
            System.out.println("-----------------");
        } catch (FileNotFoundException e) {
            System.out.println("输入的文件路径有误");
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if(sheets!=null){
                try {
                    sheets.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if(fis!=null){
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    //从外面导入已经准备好的零件进库表.xlsx
    public static void creatXlsx2(ComponentsTwo two,Pathname pn)  {
        FileInputStream fis= null;
        Workbook sheets=null;
        try {
            System.out.println("输入你要导入的零件进库表文件路径:");
            Scanner sn=new Scanner(System.in);
            String pathname=sn.nextLine();
            pn.setPathnameTwo(pathname);
            fis = new FileInputStream(pn.getPathnameTwo());
            sheets=new XSSFWorkbook(fis);
            Sheet sheet1=sheets.getSheetAt(0);
            System.out.println("导入零件进库表.xlsx成功!");
            System.out.println("------零件进库表-------");
            two.number=new String[sheet1.getLastRowNum()];
            two.amount=new int[sheet1.getLastRowNum()];
            for (int i = 1; i < sheet1.getLastRowNum()+1;i++){
                //获取行数
                Row row1=sheet1.getRow(i);
                //获取单元值并且取值
                two.number[i-1]=row1.getCell(0).getStringCellValue();
                if(i==1){
                    System.out.println(row1.getCell(0).getStringCellValue()
                            +"   "+row1.getCell(1).getStringCellValue());
                    two.amount[i-1]=0;
                }else{
                    double a=row1.getCell(1).getNumericCellValue();
                    System.out.println(row1.getCell(0).getStringCellValue()
                            +"   "+(int)a);
                    two.amount[i-1]=(int)a;
                }
                two.setNumber(two.number);
                two.setAmount(two.amount);
            }
            System.out.println("---------------------");
        } catch (FileNotFoundException e) {
            System.out.println("输入的导入文件路径有误");
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if(sheets!=null){
                try {
                    sheets.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if(fis!=null){
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
    //打印输出零件进库表的零件信息
    public  static void printXlsxTwo(Pathname pn) throws IOException {
        FileInputStream fis = new FileInputStream(pn.getPathnameTwo());
        Workbook sheets=new XSSFWorkbook(fis);
        Sheet sheet2 = sheets.getSheetAt(0);
        for (int i = 1; i < sheet2.getLastRowNum() + 1; i++) {
            //获取行数
            Row row2 = sheet2.getRow(i);
            //获取单元值并且取值
            if(i==1){
                System.out.println(row2.getCell(0).getStringCellValue()
                        + "   " + row2.getCell(1).getStringCellValue());
            }else{
                double a=row2.getCell(1).getNumericCellValue();
                System.out.println(row2.getCell(0).getStringCellValue()
                        + "   " + (int)a);
            }
        }
        System.out.println("-----------------");
        fis.close();
        sheets.close();
    }

}
