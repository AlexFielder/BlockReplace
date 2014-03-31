$(document).ready(function () { // load xml file using jquery ajax
    $.ajax({
        type: "GET",
        url: "Drawings.xml",
        dataType: "xml",
        success: function (xml) {
            var output = '<ul>';
            $(xml).find('cat').each(function () {
                var name = $(this).find('Drawing').text();
                var id = $(this).attr('DrawingID');
                output += '<li>' + name + ' (' + id + ')</li>';
            });
            output += '<ul>';
            $('#update').html(output);
        }
    });
});