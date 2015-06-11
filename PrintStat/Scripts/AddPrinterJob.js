
     $(document).ready(function () {
         $("#DeviceID").change(function () {
             var dID= $(this).val();             
             $.getJSON("../Home/LoadSizePaperForDevice", { idDev: dID },
                    function (data) {
                        var select = $("#SizePaperID");
                        select.empty();
                        var selectData;
                        $.each(data, function (index, itemData) {
                            selectData+= "<option value=" + itemData.Value + ">" + itemData.Text + "</option>";
                       
                        });
                        $("#SizePaperID").html(selectData).show();
                    });



              $.getJSON("../Home/LoadPaperTypeForDevice", { idDev: dID },
                    function (data) {
                        var select = $("#PaperTypeID");
                        select.empty();
                        var selectData;
                        $.each(data, function (index, itemData) {
                            selectData+= "<option value=" + itemData.Value + ">" + itemData.Text + "</option>";
                       
                        });
                        $("#PaperTypeID").html(selectData).show();
                    });
         });
});

                 