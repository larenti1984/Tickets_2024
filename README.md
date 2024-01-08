### scrpit para excel automate

function main(workbook: ExcelScript.Workbook) {
    // Your code here
    // Seleccionar el rango de celdas que deseas borrar
    let range = workbook.getWorksheet("TICKETS").getRange("A2:I100");

    // Borrar el contenido del rango seleccionado
    range.clear();
}

