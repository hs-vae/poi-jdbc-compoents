package Entity;
/**
 * 零件表的零件信息类
 */
public class ComponentsOne {
    /**
     * 零件表中的零件编号
     *
     */
    public  String[] number;
    /**
     * 零件表中的零件名称
     */
    public  String[] name;
    /**
     * 零件表中的零件规格
     */
    public  String[] size;
    /**
     * 零件表中的零件信息行数
     */
    public int numOne;

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

    public int getNumOne() {
        return numOne;
    }

    public void setNumOne(int numOne) {
        this.numOne = numOne;
    }



}
