<%@ Page Language="C#" AutoEventWireup="true"  CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Process XML using jQuery</title>
    <script src="Scripts/jquery-1.7.2.min.js" type="text/javascript"></script>
    
    <script type="text/javascript" language="javascript">
    $(document).ready(function(){
        $("#dvContent").append("<ul></ul>");
        $.ajax({
            type: "GET",
            url: "BookList.xml",
            dataType: "xml",
            success: function(xml){
                $(xml).find('Book').each(function(){
                var sTitle = $(this).find('Title').text();
                var sPublisher = $(this).find('Publisher').text();
                $("<li></li>").html(sTitle + ", " + sPublisher).appendTo("#dvContent ul");
            });
            },
            error: function() {
            alert("An error occurred while processing XML file.");
            }
        });
    });    
</script>
<style type="text/css">
body
{
  font-family  : Arial;
  font-size  : 10pt;
}
</style>
</head>
<body>
    <form id="form1" runat="server">
    <div id="dvContent">
    
    </div>
    </form>
</body>
</html>
