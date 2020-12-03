package Entity;
/**
 * 零件汇总表的零件信息类
 */
public class ComponentsSum {
    /**
     * 零件汇总表中的零件编号
     */
    public  String[] number;
    /**
     * 零件进库表中的零件名称
     */
    public  String[] name;
    /**
     * 零件进库表中的零件规格
     */
    public  String[] size;
    /**
     * 零件进库表中的零件数量
     */
    public  int[] amount;

    public String[] getNumber() {
        return number;
    }

    public void setNumber(String[] number) {
        this.number = number;
    }

    public String[] getName() {
        return name;
    }

    public void setName(String[] name) {
        this.name = name;
    }

    public String[] getSize() {
        return size;
    }

    public void setSize(String[] size) {
        this.size = size;
    }

    public int[] getAmount() {
        return amount;
    }

    public void setAmount(int[] amount) {
        this.amount = amount;
    }
}
