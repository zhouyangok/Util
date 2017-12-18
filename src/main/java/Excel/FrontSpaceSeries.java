package Excel;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * : front_space_series
 * entity 层
 * @author
 */

public class FrontSpaceSeries {
    //front_id : 
    private Integer frontId;

    //series_name : typ
    private String seriesName;

    //type_name : 
    private String typeName;

    //measure_status : 测量阶段
    private String measureStatus;

    //benchmark_cm_name : 基准立方镜
    private String benchmarkCmName;

    //measure_cm_name : 测量立方镜
    private String measureCmName;

    //measure_model : 测量模式
    private String measureModel;

    //measure_process_name : 测量过程
    private String measureProcessName;

    //measure_order : 测量次序
    private String measureOrder;

    //xx : xx
    private String xx;

    //xy : 
    private String xy;

    //xz : 
    private String xz;

    //yx : 
    private String yx;

    //yy : 
    private String yy;

    //yz : 
    private String yz;

    //zx : 
    private String zx;

    //zy : 
    private String zy;

    //zz : 
    private String zz;

    //x0 : 
    private String x0;

    //y0 : 
    private String y0;

    //z0 : 
    private String z0;
    
    private String rotx;
    
    private String roty;
    
    private String rotz;

    //measure_time : 测量时间
    private String measureTime;

    //measure_staff : 测量人员
    private String measureStaff;

    //measure_location : 测量地点
    private String measureLocation;

    //measure_equipment : 测量设备
    private String measureEquipment;
    
    //data_path : 文件路径
    private String dataPath;
    
    //create_time : 数据添加时间
    private Date createTime;
    
    //新增字段用于数据分析使用
    
    private String typeNames;
    
    private List<String> typeList  = new ArrayList();

    public List<String> getTypeList() {
        return typeList;
    }

    public void setTypeList(List<String> typeList) {
        this.typeList = typeList;
    }
    private String measurestatuses;
    
    private List<String> measureStatusList = new ArrayList();

    public List<String> getMeasureStatusList() {
        return measureStatusList;
    }

    public void setMeasureStatusList(List<String> measureStatusList) {
        this.measureStatusList = measureStatusList;
    }

    public String getTypeNames() {
        return typeNames;
    }

    public void setTypeNames(String typeNames) {
        this.typeNames = typeNames;
    }

  

    public void setMeasurestatuses(String measurestatuses) {
        this.measurestatuses = measurestatuses;
    }
    
    public String getMeasurestatuses() {
        return measurestatuses;
    }

    public String getRotx() {
        return rotx;
    }

    public void setRotx(String rotx) {
        this.rotx = rotx;
    }

    public String getRoty() {
        return roty;
    }

    public void setRoty(String roty) {
        this.roty = roty;
    }

    public String getRotz() {
        return rotz;
    }

    public void setRotz(String rotz) {
        this.rotz = rotz;
    }

    /**
     * front_id
     */
    public Integer getFrontId () {
        return frontId;
    }

    /**
     * front_id
     */
    public void setFrontId (Integer frontId) {
        this.frontId = frontId;
    }
    /**
     * series_name
     */
    public String getSeriesName () {
        return seriesName;
    }

    /**
     * series_name
     */
    public void setSeriesName (String seriesName) {
        this.seriesName = seriesName;
    }
    /**
     * type_name
     */
    public String getTypeName () {
        return typeName;
    }

    /**
     * type_name
     */
    public void setTypeName (String typeName) {
        this.typeName = typeName;
    }
    /**
     * measure_status
     */
    public String getMeasureStatus () {
        return measureStatus;
    }

    /**
     * measure_status
     */
    public void setMeasureStatus (String measureStatus) {
        this.measureStatus = measureStatus;
    }
    /**
     * benchmark_cm_name
     */
    public String getBenchmarkCmName () {
        return benchmarkCmName;
    }

    /**
     * benchmark_cm_name
     */
    public void setBenchmarkCmName (String benchmarkCmName) {
        this.benchmarkCmName = benchmarkCmName;
    }
    /**
     * measure_cm_name
     */
    public String getMeasureCmName () {
        return measureCmName;
    }

    /**
     * measure_cm_name
     */
    public void setMeasureCmName (String measureCmName) {
        this.measureCmName = measureCmName;
    }
    /**
     * measure_model
     */
    public String getMeasureModel () {
        return measureModel;
    }

    /**
     * measure_model
     */
    public void setMeasureModel (String measureModel) {
        this.measureModel = measureModel;
    }
    /**
     * measure_process_name
     */
    public String getMeasureProcessName () {
        return measureProcessName;
    }

    /**
     * measure_process_name
     */
    public void setMeasureProcessName (String measureProcessName) {
        this.measureProcessName = measureProcessName;
    }
    /**
     * measure_order
     */
    public String getMeasureOrder () {
        return measureOrder;
    }

    /**
     * measure_order
     */
    public void setMeasureOrder (String measureOrder) {
        this.measureOrder = measureOrder;
    }
    /**
     * xx
     */
    public String getXx () {
        return xx;
    }

    /**
     * xx
     */
    public void setXx (String xx) {
        this.xx = xx;
    }
    /**
     * xy
     */
    public String getXy () {
        return xy;
    }

    /**
     * xy
     */
    public void setXy (String xy) {
        this.xy = xy;
    }
    /**
     * xz
     */
    public String getXz () {
        return xz;
    }

    /**
     * xz
     */
    public void setXz (String xz) {
        this.xz = xz;
    }
    /**
     * yx
     */
    public String getYx () {
        return yx;
    }

    /**
     * yx
     */
    public void setYx (String yx) {
        this.yx = yx;
    }
    /**
     * yy
     */
    public String getYy () {
        return yy;
    }

    /**
     * yy
     */
    public void setYy (String yy) {
        this.yy = yy;
    }
    /**
     * yz
     */
    public String getYz () {
        return yz;
    }

    /**
     * yz
     */
    public void setYz (String yz) {
        this.yz = yz;
    }
    /**
     * zx
     */
    public String getZx () {
        return zx;
    }

    /**
     * zx
     */
    public void setZx (String zx) {
        this.zx = zx;
    }
    /**
     * zy
     */
    public String getZy () {
        return zy;
    }

    /**
     * zy
     */
    public void setZy (String zy) {
        this.zy = zy;
    }
    /**
     * zz
     */
    public String getZz () {
        return zz;
    }

    /**
     * zz
     */
    public void setZz (String zz) {
        this.zz = zz;
    }
    /**
     * x0
     */
    public String getX0 () {
        return x0;
    }

    /**
     * x0
     */
    public void setX0 (String x0) {
        this.x0 = x0;
    }
    /**
     * y0
     */
    public String getY0 () {
        return y0;
    }

    /**
     * y0
     */
    public void setY0 (String y0) {
        this.y0 = y0;
    }
    /**
     * z0
     */
    public String getZ0 () {
        return z0;
    }

    /**
     * z0
     */
    public void setZ0 (String z0) {
        this.z0 = z0;
    }
    /**
     * measure_time
     */
    public String getMeasureTime () {
        return measureTime;
    }

    /**
     * measure_time
     */
    public void setMeasureTime (String measureTime) {
        this.measureTime = measureTime;
    }
    /**
     * measure_staff
     */
    public String getMeasureStaff () {
        return measureStaff;
    }

    /**
     * measure_staff
     */
    public void setMeasureStaff (String measureStaff) {
        this.measureStaff = measureStaff;
    }
    /**
     * measure_location
     */
    public String getMeasureLocation () {
        return measureLocation;
    }

    /**
     * measure_location
     */
    public void setMeasureLocation (String measureLocation) {
        this.measureLocation = measureLocation;
    }
    /**
     * measure_equipment
     */
    public String getMeasureEquipment () {
        return measureEquipment;
    }

    /**
     * measure_equipment
     */
    public void setMeasureEquipment (String measureEquipment) {
        this.measureEquipment = measureEquipment;
    }

    public String getDataPath() {
        return dataPath;
    }

    public void setDataPath(String dataPath) {
        this.dataPath = dataPath;
    }

    public Date getCreateTime() {
        return createTime;
    }

    public void setCreateTime(Date createTime) {
        this.createTime = createTime;
    }
}
