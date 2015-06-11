
    $(document).ready(function ($) {
        var e = document.getElementById('color');
        e.style.display = 'none'; 
        $('#TypeConsumableID').change(function() {
            var obj = $("#TypeConsumableID :selected");
            var url = "/Consumable/CheckCartridge/";
            
            $.get(url, { TypeName: "картридж" }, function (data) {
                var id =(data);
                if (obj.val() == id) {
                    e.style.display = 'block';
                } else {
                    e.style.display = 'none';
                }
            });
        });
    });
