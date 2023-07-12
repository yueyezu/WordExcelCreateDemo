<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <style>
        #bg {
            display: none;
            width: 100%;
            height: 100%;
            background-color: black;
            z-index: 1001;
            position: absolute;
            top: 0%;
            left: 0%;
            -moz-opacity: 0.2;
            opacity: .20;
            filter: alpha(opacity=20);
        }
    </style>
    <script src="jsLib/jquery-1.10.2.min.js" type="text/javascript"></script>
    <script src="jsLib/waitingTip.js" type="text/javascript"></script>
    <script>
        $(function () {
            var w2 = new WaitingTip({ innerHTML: '<img src="images/waiting.gif" />word生成中...' });
            $('#ajaxword').click(function (e) {
                console.log(e);
                var wordName = $('#wordName').val();
                w2.show($('#bg'), "center");
                $('#bg').show();
                $.ajax({
                    url: "Handler.ashx?action=createword",
                    dataType: 'json',
                    data: {
                        wordName: wordName
                    },
                    success: function (e) {
                        if (e.success) {
                            window.open("Handler.ashx?action=getword&wordName=" + wordName, "_self");
                        } else {
                            alert('未能成功生成');
                        }
                        w2.hide();
                        $('#bg').hide();
                    },
                    error: function (x) {
                        alert('请求错误！');
                        w2.hide();
                        $('#bg').hide();
                    }
                });
            });


            $('#ajaxexcel').click(function (e) {
                var excelName = $('#excelName').val();
                w2.show($('#bg'), "center");
                $('#bg').show();
                $.ajax({
                    url: "Handler.ashx?action=createExcel",
                    dataType: 'json',
                    data: {
                        excelName: excelName
                    },
                    success: function (e) {
                        if (e.success) {
                            window.open("Handler.ashx?action=getexcel&excelName=" + excelName, "_self");
                        } else {
                            alert('未能成功生成');
                        }
                        w2.hide();
                        $('#bg').hide();
                    },
                    error: function (x) {
                        alert('请求错误！');
                        w2.hide();
                        $('#bg').hide();
                    }
                });
            });

        });
    </script>
</head>
<body style="min-height: 500px;">
    <div id="bg"></div>
    <div style="margin: 100px 400px;">
        <span>word的名称：</span><input type="text" id="wordName" />
        <input type="button" id="ajaxword" value="导出word" />
    </div>
    <div style="margin: 20px 400px;">
        <span>excel的名称：</span><input type="text" id="excelName" />
        <input type="button" id="ajaxexcel" value="导出Excel" />
    </div>
</body>
</html>
