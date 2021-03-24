package bean;

/**
 * @author liujingeng
 * @description 下拉框实体类
 * @create 2020/02/27
 */
public class ComboBoxItem {

    private String name;

    private String value;

    public ComboBoxItem(String name, String value) {
        this.name = name;
        this.value = value;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }
}