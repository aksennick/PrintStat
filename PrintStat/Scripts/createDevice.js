

$("#Manufacturer").attr("placeholder", "Начните вводить производителя");

$(document).ready(function() {
    $('#Manufacturer').autocomplete({
        source: '/Manufacturer/AutocompleteManufacturer'

    });
});


$(document).ready(function () {

    $("#Manufacturer").keyup(function () {
        $.get(('/Model/GetModelForManufacturer/' + $(this).val()), function (data) {
            var models = $.parseJSON(data);
            var ddlSelectedModels = $("#ModelID");
            $("#ModelID > option").remove();
                

            if (data != "[]") {
                for (i = 0; i < models.length; i++) {
                    ddlSelectedModels.append($("<option />").val(models[i].ID).text(models[i].Name));
                }
            }
            else
            {
                $("#ModelID > option").remove();
                //ddlSelectedModels.append($("<option />").placeholder("Модели для выбора отсутствуют"));
            }
        });
    });
});


