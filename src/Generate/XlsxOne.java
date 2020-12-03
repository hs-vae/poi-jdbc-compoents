package Generate;
/*
    生成默认的零件表.xlsx
    包含零件编号,名称,规格
 */
import Entity.ComponentsOne;
import Entity.Pathname;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.Scanner;

/**
 * 零件表的生成类
 */
public class XlsxOne {

    public static void creat(ComponentsOne one, Pathname pn){
        Scanner sc=new Scanner(System.in);
        int mark=0;
        while (mark!=1){
            System.out.println("输入你想生成零件表的方式代号:");
            System.out.println("1.手动输入零件表信息 2.从外界导入零件表信息");
            int choose=sc.nextInt();
            if(choose==1){
                creatXlsx1(one,pn);
                mark=1;
            }else if(choose==2){
                creatXlsx2(one,pn);
                mark=1;
            }else{
                System.out.println("输入生成代号错误,请重新输入!");
            }
        }
    }
    //生成手动输入零件信息的零件表.xlsx
    public static void creatXlsx1(ComponentsOne one,Pathname pn) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("零件表");
        FileOutputStream fos = null;
        FileInputStream fis = null;
        Workbook sheets = null;
        try {
            Scanner sn = new Scanner(System.in);
            System.out.println("输入零件表.xlsx文件生成的位置:(需要带上文件名称)");
            String pathname = sn.nextLine();
            pn.setPathnameOne(pathname);
            fos = new FileOutputStream(pn.getPathnameOne());
            System.out.println("输入零件表的行数:");
            one.setNumOne(sn.nextInt());
            XSSFRow row = null;
            XSSFCell cell = null;
            Scanner sc = new Scanner(System.in);
            one.number = new String[one.getNumOne() + 1];
            one.name = new String[one.getNumOne() + 1];
            one.size = new String[one.getNumOne() + 1];
            //用String类型数组存放输入的零件编号,名字,规格
            for (int i = 0; i < one.getNumOne() + 1; i++) {
                if (i == 0) {
                    one.number[0] = "零件编号";
                    one.name[0] = "零件名称";
                    one.size[0] = "零件规格";
                } else {
                    System.out.println("请输入" + i + "行零件编号:");
                    one.number[i] = sc.nextLine();
                    System.out.println("请输入" + i + "行零件名称:");
                    one.name[i] = sc.nextLine();
                    System.out.println("请输入" + i + "行零件规格:");
                    one.size[i] = sc.nextLine();
                }
            }
            one.setNumber(one.number);
            one.setName(one.name);
            one.setSize(one.size);
            for (int i = 0; i < one.getNumOne() + 2; i++) {
                if (i == 0) {
                    row = sheet.createRow(i);
                    cell = row.createCell(i);
                    //设置单元格范围地址,参数分别表示为(起始行号,终止行号,起始列号,终止列号)
                    CellRangeAddress region = new CellRangeAddress(0, 0, 0, 2);
                    row.createCell(0).setCellValue("零件表信息");
                    //设置水平居中
                    CellStyle style = workbook.createCellStyle();
                    style.setAlignment(HorizontalAlignment.CENTER);
                    cell.setCellStyle(style);
                    sheet.addMergedRegion(region);
                } else {
                    row = sheet.createRow(i);
                    cell = row.createCell(i);
                    row.createCell(0).setCellValue(one.getNumber()[i - 1]);
                    row.createCell(1).setCellValue(one.getName()[i - 1]);
                    row.createCell(2).setCellValue(one.getSize()[i - 1]);
                }
            }
            System.out.println("生成成功,已存入" + pathname + "目录中!");
            workbook.write(fos);
            workbook.close();
            fos.close();
            //打印上面输入的零件表信息
            fis = new FileInputStream(pn.getPathnameOne());
            sheets = new XSSFWorkbook(fis);
            Sheet sheet1 = sheets.getSheetAt(0);
            System.out.println("------零件表------");
            for (int i = 1; i < sheet1.getLastRowNum() + 1; i++) {
                //获取行数
                Row row1 = sheet1.getRow(i);
                //获取单元值并且取值
                System.out.println(row1.getCell(0).getStringCellValue()
                        + "   " + row1.getCell(1).getStringCellValue()
                        + "   " + row1.getCell(2).getStringCellValue());
            }
            System.out.println("-----------------");

        } catch (FileNotFoundException e) {
            System.out.println("文件路径错误");
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (sheets != null) {
                try {
                    sheets.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    //从外面导入已经准备好的零件表.xlsx
     public  static void  creatXlsx2(ComponentsOne one,Pathname pn)  {
          FileInputStream fis = null;
          Workbook sheets = null;
              try {
                  System.out.println("输入你要导入的零件表文件路径:");
                  Scanner sn = new Scanner(System.in);
                  String pathname=sn.nextLine();
                  pn.setPathnameOne(pathname);
                  fis = new FileInputStream(pn.getPathnameOne());
                  sheets = new XSSFWorkbook(fis);
                  Sheet sheet1 = sheets.getSheetAt(0);
                  System.out.println("导入零件表.xlsx成功!");
                  System.out.println("------零件表-------");
                  one.number = new String[sheet1.getLastRowNum()]; //4
                  one.name = new String[sheet1.getLastRowNum()];
                  one.size = new String[sheet1.getLastRowNum()];
                  for (int i = 1; i < sheet1.getLastRowNum() + 1; i++) { //1 2 3 4
                      //获取行数
                      Row row1 = sheet1.getRow(i);
                      //获取单元值并且取值
                      System.out.println(row1.getCell(0).getStringCellValue()
                              + "   " + row1.getCell(1).getStringCellValue()
                              + "   " + row1.getCell(2).getStringCellValue());
                      one.number[i-1]=row1.getCell(0).getStringCellValue();
                      one.name[i-1]=row1.getCell(1).getStringCellValue();
                      one.size[i-1]=row1.getCell(2).getStringCellValue();
                      one.setNumber(one.number);
                      one.setName(one.name);
                      one.setSize(one.size);
                  }
                  System.out.println("-----------------");
              } catch (FileNotFoundException e) {
                  System.out.println("文件导入的路径错误");
                  e.printStackTrace();
              } catch (IOException e) {
                  e.printStackTrace();
              } finally {
                  if (sheets != null) {
                      try {
                          sheets.close();
                      } catch (IOException e) {
                          e.printStackTrace();
                      }
                  }
                  if (fis != null) {
                      try {
                          fis.close();
                      } catch (IOException e) {
                          e.printStackTrace();
                      }
                  }
              }
    }
    //打印输出零件表信息
    public  static void printXlsxOne(Pathname pn) throws IOException {
        FileInputStream fis = new FileInputStream(pn.getPathnameOne());
        Workbook sheets=new XSSFWorkbook(fis);
        Sheet sheet1 = sheets.getSheetAt(0);
        for (int i = 1; i < sheet1.getLastRowNum() + 1; i++) {
                //获取行数
                Row row1 = sheet1.getRow(i);
                //获取单元值并且取值
                System.out.println(row1.getCell(0).getStringCellValue()
                        + "   " + row1.getCell(1).getStringCellValue()
                        + "   " + row1.getCell(2).getStringCellValue());
            }
            System.out.println("-----------------");
        fis.close();
        sheets.close();
        }
}
