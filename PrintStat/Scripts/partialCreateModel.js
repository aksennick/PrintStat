
$("#ManufacturerID").attr("placeholder", "Начните вводить производителя");

$("#dialogDivModel").ready(function() {
    $('#ManufacturerID').autocomplete({
        source: '/Manufacturer/AutocompleteManufacturer',
        minLength: 0
    }).focus(function() {
        $(this).autocomplete("search");
    });
});






