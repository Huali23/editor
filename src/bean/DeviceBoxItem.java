package bean;

/**
 * @author liujingeng
 * @description 设备下拉框实体类
 * @create 2020/02/27
 */
public class DeviceBoxItem extends ComboBoxItem{

    private String parentName;

    public String getParentName() {
        return parentName;
    }

    public void setParentName(String parentName) {
        this.parentName = parentName;
    }

    public DeviceBoxItem(String parentValue, String name, String value) {
        super(name, value);
        this.parentName = parentValue;
    }
}