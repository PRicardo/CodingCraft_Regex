$(document).ready(function () {
    $('#excelRead').change(function () {
        $('#btn-analise').prop('disabled', false);
    });
});

function upload() {
    var formData = new FormData();
    var totalFiles = document.getElementById("excelRead").files.length;

    for (var i = 0; i < totalFiles; i++) {
        var file = document.getElementById("excelRead").files[i];
        
        formData.append("FileUpload", file);
    }

    $.ajax({
        type: 'POST',
        url: '/Home/AnaliseExcel',
        data: formData,
        dataType: 'html',
        contentType: false,
        processData: false,
        beforeSend: function () {
            $('#modelLoader').modal('show')
        },
        success: function (response) {
            $('#modelLoader').modal('hide');
            $('#content-analise').html(response);
        },
        error: function (error) {
            $('#modelLoader').modal('hide');
            alert("error");
        }
    });
}