package Entity;

/**
 * 零件进库表的零件信息类
 */
public class ComponentsTwo {
    /**
     * 零件进库表中的零件编号
     */
    public  String[] number;
    /**
     * 零件进库表中的零件数量
     */
    public  int[] amount;
    /**
     * 零件进库表中的零件信息行数
     */
    public  int numTwo;

    public String[] getNumber() {
        return number;
    }

    public void setNumber(String[] number) {
        this.number = number;
    }

    public int[] getAmount() {
        return amount;
    }

    public void setAmount(int[] amount) {
        this.amount = amount;
    }

    public int getNumTwo() {
        return numTwo;
    }

    public void setNumTwo(int numTwo) {
        this.numTwo = numTwo;
    }

}
