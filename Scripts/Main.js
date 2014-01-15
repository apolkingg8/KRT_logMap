/**
 * Created by EIJI on 2013/12/25.
 * Copyright (c) 2014 apolkingg8
 * Licence under MIT Licences.
 */

$(document).ready(function(){
    if(Modernizr.svg){
        initSVG();
    }else{
        alert("Sorry, your browser does not support SVG. PLEASE change one.");
        location.href = 'http://www.opera.com/zh-tw';
    }
});

var s = Snap('#svg'),
    chartR = Snap('#redChart'),
    chartO = Snap('#orangeChart');

var getDataByDate = function(year, month, stationId){
    var temp = {In:'', Out:''};

    stationId = stationId === 'R10'
        ? 'O5'
        : stationId;

    for(var i = 0; i < _DATA.length; i += 1){
        if(_DATA[i].Year === year && _DATA[i].Month === month){
            for(var j = 1; j < _DATA[i].Stations.length; j += 1){
                if(_DATA[i].Stations[j].ID === stationId){
                    temp.In = _DATA[i].Stations[j].In;
                    temp.Out = _DATA[i].Stations[j].Out;
                }
            }
        }
    }

    return temp;
};

var getData = function(percentage, stationId){
    var index = Math.round(_DATA.length * (percentage / 100));
    var temp = {};

    stationId = stationId === 'R10'
        ? 'O5'
        : stationId;
    if(_DATA[index]){
        temp.dateText = _DATA[index].Year + 1911 + '/' + _DATA[index].Month;

        for(var i = 0; i < _DATA[index].Stations.length; i += 1){
            if(_DATA[index].Stations[i].ID === stationId){
                temp.station = _DATA[index].Stations[i];
            }
        }
    }
    return temp;
};

var initChart = function(){
    var $chart = $('#redChart'),
        width = $chart.width(),
        height = $chart.height();

    drawChart(chartR, 1, 4500000, _DATA, width, height, 0, function(){
        drawChart(chartO, 2, 2000000, _DATA, width, height, 0)
    });
};

var drawChart = function(chart, stationIndex, base, data, width, height, i, callback){
    var x = width / data.length,
        y = height / base;
    var px = i * x,
        py = height - (data[i].Stations[stationIndex].In * y);

    if(i < data.length - 1){
        var npx = (i + 1) * x,
            npy = height - (data[i + 1].Stations[stationIndex].In * y),
            time = 5;

        chart.line(px, py, px, py).animate({
            x1: npx,
            y1: npy
        }, time, function(){
            drawChart(chart, stationIndex, base, data, width, height, i + 1, callback);
        }).attr({
                class: i
            });

    }else{
        if(typeof callback === "function"){
            callback();
        }
    }
};

var loadData = function(callback){

};

var initSlider = function(){
    var $circle = $('circle');

    $('#slider').slider({
        orientation: "vertical",
        slide: function(event, ui){

            var totalRed = getData(ui.value, '紅線小計'),
                totalOrange = getData(ui.value, '橘線小計');
            var $divTotal = $('#infoBox').find('.total');
            var $charts = $('.chart'),
                $redLine = $('#redChart').find('line').filter('.' + Math.round(ui.value / 100 * _DATA.length)),
                redX = $redLine.attr('x2'),
                redY = $redLine.attr('y2'),
                $orangeLine = $('#orangeChart').find('line').filter('.' + Math.round(ui.value / 100 * _DATA.length)),
                orangeX = $orangeLine.attr('x2'),
                orangeY = $orangeLine.attr('y2');

            // change text
            if(!totalRed.station){
                totalRed = getData(ui.value, '總計');
            }
            $divTotal.filter('.red').children('.value').text(totalRed.station.In);

            if(totalOrange.station){
                $divTotal.filter('.orange').children('.value').text(totalOrange.station.In);
            }else{
                $divTotal.filter('.orange').children('.value').text('');
            }

            // change chart
            $charts.find('circle').remove();
            chartR.circle(redX, redY, 5);
            chartO.circle(orangeX, orangeY, 5);

            // change circle
            $circle.each(function(){
                var data = getData(ui.value, $(this).attr('class').split(' ')[1]);

                $('#dateBox').text(data.dateText);
                if(data.station){
                    $(this).attr('r', parseInt(data.station.In) / 20000 + 3);
                }else{
                    $(this).attr('r', 0);
                }
            });
        },
        value: 100,
        max: 99
    });
};

var initSVG = function(){
    s = Snap('#svg');
    chartR = Snap('#redChart');
    chartO = Snap('#orangeChart');

    var docHeight = document.getElementById('mainContent').clientHeight,
        docWidth = Math.floor(docHeight / 1.26);

    var x = docWidth / 16,
        y = docHeight / 23;

    var clock = 150;

    var lines = [],
        texts = [],
        circles = [];

    var orangeLine = [
        [1, 18, 'O1', '西子灣', [-20, 25]], [3, 18, 'O2', '鹽埕埔', [-20, 25]],
        [4, 17, 'O4', '市議會', [-70, -20]], [5, 16, 'O5', '美麗島', [0, -25]],
        [6, 16, 'O6', '信義國小', [0, -25]], [7, 16, 'O7', '文化中心', [0, -25]],
        [8, 16, 'O8', '五塊厝', [0, -25]], [9, 16, 'O9', '技擊館', [0, -25]],
        [10, 16, 'O10', '衛武營', [0, -25]], [11, 16, 'O11', '鳳山西站', [0, -25]],
        [12, 16, 'O12', '鳳山', [0, -25]], [13, 16, 'O13', '大東', [0, -25]],
        [14, 16, 'O14', '鳳山國中', [0, -25]], [15, 17, 'OT1', '大寮', [15, -15]]
    ];

    var redLine = [
        [13, 22, 'R3', '小港', [15, 0]], [11, 22, 'R4', '高雄國際機場', [-20, 25]],
        [9, 22, 'R4A', '草衙', [15, -15]], [8, 21, 'R5', '前鎮高中', [15, -15]],
        [7, 20, 'R6', '凱旋', [15, -15]], [6, 19, 'R7', '獅甲', [15, -15]],
        [5, 18, 'R8', '三多商圈', [15, -10]], [5, 17, 'R9', '中央公園', [15, -10]],
        [5, 16, 'R10', '', [-70, -17]], [5, 15, 'R11', '高雄車站', [15, -12]],
        [5, 14, 'R12', '後驛', [15, -12]], [5, 13, 'R13', '凹子底', [15, -12]],
        [5, 12, 'R14', '巨蛋', [15, -12]], [5, 11, 'R15', '生態園區', [15, -12]],
        [5, 10, 'R16', '左營', [15, -12]], [5, 9, 'R17', '世運', [15, -12]],
        [5, 8, 'R18', '油廠國小', [15, -12]], [6, 7, 'R19', '楠梓加工區', [15, -12]],
        [7, 6, 'R20', '後勁', [15, -12]], [7, 5, 'R21', '都會公園', [15, -12]],
        [7, 4, 'R22', '青埔', [15, -12]], [7, 3, 'R22A', '橋頭糖廠', [15, -12]],
        [7, 2, 'R23', '橋頭火車站', [15, -12]], [7, 1, 'R24', '南岡山', [15, -12]]
    ];

    var drawLine = function(points, i, attrClass, callback){
        var stationId = points[i][2],
            stationName = points[i][3],
            textShiftX = points[i][4][0],
            textShiftY = points[i][4][1],
            px = points[i][0] * x,
            py = points[i][1] * y;

        var stationData = getDataByDate(102, 11, stationId);

        if(i < points.length - 1){
            var npx = points[i + 1][0] * x,
                npy = points[i + 1][1] * y,
                //time = Math.sqrt(Math.pow((px - npx), 2) + Math.pow((py - npy), 2)) * clock;
                time = 100;

            var l = s.line(px, py, px, py).animate({
                x1: npx,
                y1: npy
            }, time, function(){
                drawLine(points, i + 1, attrClass, callback);

                var c = s.circle(px, py, 0).animate({
                    r: (stationData.In / 20000) + 3
                    //r: 0
                }, clock).attr({
                    class: attrClass + ' ' + stationId
                });
                //circles.push(c);

                var t = s.text(px + textShiftX, py + textShiftY, stationName).attr({
                    class: stationId
                });
                if(['O5', 'O6', 'O7', 'O8', 'O9', 'O10', 'O11', 'O12', 'O13', 'O14'].indexOf(stationId) > -1){
                    t.attr({
                        transform: 'rotate(-30deg)'
                    });
                }
                //texts.push(t);
            }).attr({
                class: attrClass
            });
            //lines.push(l);
        }else{
            s.circle(px, py, 0).animate({
                r: 10
            }, clock, function(){
                if(typeof callback === "function"){
                    callback();
                }
            }).attr({
                class: attrClass + ' ' + stationId
            });
            s.text(px + textShiftX, py + textShiftY, stationName).attr({
                class: stationId
            });
        }
    };

    $('#mainContent').css('width', docWidth);

    drawLine(redLine, 0, 'redLine', function(){
        initSlider();
        $('#infoBox').show(function(){
            initChart()
        });

    });
    setTimeout(function(){
        drawLine(orangeLine, 0, 'orangeLine');
    }, 600);

};


