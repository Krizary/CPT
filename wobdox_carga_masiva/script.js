function copiarYCrearArchivo() {
  Excel.run(async (context) => {
    const hoja = context.workbook.worksheets.getActiveWorksheet();
    const rango = hoja.getRange("A1:C10");
    rango.load("values");
    await context.sync();

    const nuevoLibro = context.application.createWorkbook();
    const nuevaHoja = nuevoLibro.worksheets.getItem("Sheet1");
    const rangoNuevo = nuevaHoja.getRange("A1:C10");
    rangoNuevo.values = rango.values;
    await context.sync();

    console.log("Rango copiado y nuevo archivo creado.");
  }).catch(function (error) {
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}