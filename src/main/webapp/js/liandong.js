/**
 * Created by medo_zy on 2017-9-15.
 */

/********** 航天器系列数据 **********/
var GP =['神舟系列','天宫系列','尖兵系列','其他系列'];
/********** 航天器型号数据 **********/
var GT = [
    ['神舟一号','神舟二号','神舟三号','神舟四号','神舟五号','神舟六号','神舟七号','神舟八号','神舟九号','神舟十号'],
    ['天宫一号','天宫二号','天宫三号','天宫四号','天宫五号','天宫六号','天宫七号','天宫八号','天宫九号','天宫十号'],
    ['尖兵一号','尖兵二号','尖兵三号','尖兵四号','尖兵五号','尖兵六号','尖兵七号','尖兵八号','尖兵九号'],
    ['其他系列'],
];

$(function(){
    $("#test").ProvinceCity();
});
$.fn.ProvinceCity = function(){
    var _self = this;
    _self.data("seriesName",["请选择", ""]);
    _self.data("typeName",["请选择", ""]);
    var $sel1 =$("#seriesName");
    var $sel2 = $("#typeName");
    //默认省级下拉
    if(_self.data("typeName")){
        $sel1.append("<option value='"+_self.data("seriesName")[1]+"'>"+_self.data("seriesName")[0]+"</option>");
    }
    $.each( GP , function(index,data){
        $sel1.append("<option value='"+data+"'>"+data+"</option>");
    });
    //默认的1级城市下拉
    if(_self.data("typeName")){
        $sel2.append("<option value='"+_self.data("typeName")[1]+"'>"+_self.data("typeName")[0]+"</option>");
    }

    //省级联动 控制
    var index1 = "" ;
    $sel1.change(function(){
        //清空其它2个下拉框
        $sel2[0].options.length=0;

        index1 = this.selectedIndex;
        if(index1==0){	//当选择的为 “请选择” 时
            if(_self.data("typeName")){
                $sel2.append("<option value='"+_self.data("typeName")[1]+"'>"+_self.data("typeName")[0]+"</option>");
            }

        }else{
            $.each( GT[index1-1] , function(index,data){
                $sel2.append("<option value='"+data+"'>"+data+"</option>");
            });

        }
    }).change();

    return _self;
};