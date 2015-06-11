
var myTag = new Array();
var myPaperType = new Array();
var mySizePaper = new Array();

function showAddModel() {

    var str = $('#Manufacturer').val();
    $('#dialogDivModel').load('/Model/PartialCreateModel/').dialog(
    {
        autoOpen: true,
        title: 'Добавление модели устройства',
        modal: true,
        resizable: true,
        minWidth: '550',
        minHeight: '60',
        width: '550',
        height: 'auto',

        buttons:
        {
            "Добавить": function() {
                var manufacturer = $('#ManufacturerID').val();
                if (manufacturer == "") {
                    alert('Поле "Производитель" обязательно для заполнения');
                    return;
                }
                var manuf = { Name: manufacturer };
                var idManuf;
                $.post("/Manufacturer/AddManufacturer", manuf, null, "json")
                    .success(function(data) {
                        myTag = [];
                        myPaperType = [];
                        mySizePaper = [];
                        $('#divTag input:checked').each(function() {
                            myTag.push($(this).val());
                        });

                        $('#divSizePaper input:checked').each(function() {
                            mySizePaper.push($(this).val());
                        });

                        $('#divPaperType input:checked').each(function() {
                            myPaperType.push($(this).val());
                        });

                        var dataObject = JSON.stringify({
                            'ManufacturerID': data,
                            'Name': $('#NameModel').val(),
                            'DeviceTypeID': $('#DeviceTypeID').val(),
                            'PrintKindID': $('#PrintKindID').val(),
                            'ChosenTagIds': myTag,
                            'ChosenPaperTypeIds': myPaperType,
                            'ChosenSizePaperIds': mySizePaper
                        });


                        $.ajax({
                            type: 'POST',
                            url: '/Model/AddModel',
                            data: dataObject,
                            contentType: 'application/json',
                            success: function(data) {
                                if (data) {
                                    $('#dialogDivModel').dialog("close");
                                    // alert("Типо не добавил");
                                }
                            },
                            error: function() { alert('Не удалось добавить модель'); }
                        });

                    })
                    .error(function() {
                        alert("Не удалось добавить производителя");
                    });
            },
            "Отмена": function() {
                $(this).dialog("close");
            }
        }
    });



    setTimeout(function() {
        $("#dialogDivModel").find('input').each(function() {
            if (this.name == "ManufacturerID") {
                
                // $("#ManufacturerID").attr("value", str);
                $("#ManufacturerID").val(str);
            }
        })}, 200);


}
