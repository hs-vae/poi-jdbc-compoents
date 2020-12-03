package Entity;

/**
 * 零件表,零件进库表,零件汇总表文件路径类
 */
public class Pathname {
    /**
     * 零件表文件路径
     */
    private static String pathnameOne;
    /**
     * 零件进库表文件路径
     */
    private static String pathnameTwo;
    /**
     * 零件汇总表文件路径
     */
    private static String pathnameSum;

    public  String getPathnameSum() {
        return pathnameSum;
    }

    public  void setPathnameSum(String pathnameSum) {
        Pathname.pathnameSum = pathnameSum;
    }

    public String getPathnameOne() {
        return pathnameOne;
    }

    public  void setPathnameOne(String pathnameOne) {
        Pathname.pathnameOne = pathnameOne;
    }

    public  String getPathnameTwo() {
        return pathnameTwo;
    }

    public  void setPathnameTwo(String pathnameTwo) {
        Pathname.pathnameTwo = pathnameTwo;
    }
}
