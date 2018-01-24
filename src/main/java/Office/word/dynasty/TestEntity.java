package Office.word.dynasty;

/**
 * Created with IntelliJ IDEA.
 *
 * @Author medo_zy
 * @Desciption:
 * @Date 2017-12-18 8:57
 */
public class TestEntity {
    
    private Integer id;
    private String name;
    private String type;
    private String measureName;
    private String cmName;
    
    private String X;
    private String Y;
    private String Z;
    
    public TestEntity(){
    };

    public TestEntity(Integer id, String name, String type, String measureName, String cmName, String x, String y, String z) {
        this.id = id;
        this.name = name;
        this.type = type;
        this.measureName = measureName;
        this.cmName = cmName;
        X = x;
        Y = y;
        Z = z;
    }

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getMeasureName() {
        return measureName;
    }

    public void setMeasureName(String measureName) {
        this.measureName = measureName;
    }

    public String getCmName() {
        return cmName;
    }

    public void setCmName(String cmName) {
        this.cmName = cmName;
    }

    public String getX() {
        return X;
    }

    public void setX(String x) {
        X = x;
    }

    public String getY() {
        return Y;
    }

    public void setY(String y) {
        Y = y;
    }

    public String getZ() {
        return Z;
    }

    public void setZ(String z) {
        Z = z;
    }
}
