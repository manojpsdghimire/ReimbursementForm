
function sum_expenses() {

    var airfare_expense = parseFloat(document.getElementById('airfare_expense').value);
    var personal_auto_expense = parseFloat(document.getElementById('personal_expense').value);
    var lodging_expense = parseFloat(document.getElementById('lodging_expense').value);
    var meal_expense = parseFloat(document.getElementById('meal_expense').value);
    var redgestration_expense = parseFloat(document.getElementById('registration_expense').value);
    var misc_expense = parseFloat(document.getElementById('misci_expense').value);
    var total_expenses = parseFloat(Number(airfare_expense + personal_auto_expense + lodging_expense + meal_expense + redgestration_expense + misc_expense).toFixed(2));

    //document.getElementById('total_expense').value = parseFloat(Number(total_expenses).toFixed(2));
    document.getElementById('total_expense').value = total_expenses.toFixed(2);

}



function setTwoNumberDecimal(el) {

    if (el.value == "") {
        el.value = "0.00"
    }
    else {
        el.value = parseFloat(el.value).toFixed(2);
    }
}
