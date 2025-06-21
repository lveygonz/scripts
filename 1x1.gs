/** @OnlyCurrentDoc */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ Setup') // Nombre del menú
    .addItem('☕ Crea las pestañas para cada estudiante', 'createStudentTabs') // Elemento del menú
    .addItem('🔄 Actualiza los cambios en las pestañas de los estudiantes', 'updateStudentTabs') // Elemento del menú
    .addItem('➕ Crea una nueva actividad', 'createActivitySheet') // Elemento del menú
    .addToUi();
  ui.createMenu('📥 Grading') // Nombre del menú
    .addItem('🚀 Envía las calificaciones a las pestañas de los estudiantes', 'mainSendGrades') // Elemento del menú
    .addItem('🧼 Elimina las calificaciones de una actividad', 'clearActivityGrades') // Elemento del menú
    .addItem('🚚 Recoge las calificaciones globales en la hoja de grupo', 'fetchGrades') // Elemento del menú
    .addToUi();
}

function createActivitySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName("Act Tmpl");

  if (!templateSheet) {
    SpreadsheetApp.getUi().alert("❌ No se encontró la pestaña 'Act Tmpl'.");
    return;
  }

  const ui = SpreadsheetApp.getUi();

  // Preguntar número de actividad
  const response = ui.prompt("Input activity number", "Escribe el número de actividad (ej. 3):", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
    Logger.log("⏹️ Acción cancelada por el usuario.");
    return;
  }

  const activityNumber = response.getResponseText().trim();
  if (!activityNumber) {
    ui.alert("⚠️ No se ingresó un número de actividad válido.");
    return;
  }
  if (!/^\d+$/.test(activityNumber)) {
    ui.alert("⚠️ El número de actividad debe ser un número entero positivo.");
    return;
  }

  const newSheetName = `ACT ${activityNumber}`;
  if (ss.getSheetByName(newSheetName)) {
    ui.alert(`❌ Ya existe una pestaña con el nombre "${newSheetName}".`);
    return;
  }

  // Preguntar si quiere evaluar más de 3 criterios
  const moreCriteriaResponse = ui.alert(
    "¿Quieres evaluar más de 3 criterios?",
    ui.ButtonSet.YES_NO
  );

  let numCriteriaExtra = 0;
  const defaultCriteriaCount = 3;

  if (moreCriteriaResponse === ui.Button.YES) {
    const criteriaNumberResponse = ui.prompt(
      "Número de criterios",
      "¿Cuántos criterios quieres evaluar en total?",
      ui.ButtonSet.OK_CANCEL
    );
    if (criteriaNumberResponse.getSelectedButton() !== ui.Button.OK) {
      Logger.log("⏹️ Acción cancelada al preguntar número de criterios.");
      return;
    }
    const totalCriteriaStr = criteriaNumberResponse.getResponseText().trim();
    if (!/^\d+$/.test(totalCriteriaStr)) {
      ui.alert("⚠️ El número de criterios debe ser un entero positivo.");
      return;
    }
    const totalCriteria = parseInt(totalCriteriaStr, 10);
    if (totalCriteria <= defaultCriteriaCount) {
      ui.alert(`⚠️ El número debe ser mayor que ${defaultCriteriaCount}.`);
      return;
    }
    numCriteriaExtra = totalCriteria - defaultCriteriaCount;
  }

  // Copiar plantilla y renombrar
  const newSheet = templateSheet.copyTo(ss).setName(newSheetName);
  newSheet.getRange("A2").setValue(newSheetName);

  // Insertar columnas extras si las hay
  if (numCriteriaExtra > 0) {
    const lastDefaultCol = 5; // Columna E (5), última de criterios por defecto
    const lastRow = newSheet.getMaxRows();

    for (let i = 0; i < numCriteriaExtra; i++) {
      newSheet.insertColumnAfter(lastDefaultCol + i);

      const sourceColIndex = lastDefaultCol;
      const targetColIndex = lastDefaultCol + i + 1;

      const sourceRange = newSheet.getRange(1, sourceColIndex, lastRow);
      const targetRange = newSheet.getRange(1, targetColIndex, lastRow);

      sourceRange.copyTo(targetRange, {contentsOnly: false});
    }

    // Ahora hacemos la fusión horizontal de la fila 1 y fila 2, desde columna D (4) hasta la última
    const firstCol = 4; // D
    const lastCol = lastDefaultCol + numCriteriaExtra; // columna final con las nuevas insertadas

    // Fusión fila 1
    newSheet.getRange(1, firstCol, 1, lastCol - firstCol + 1).mergeAcross();

    // Fusión fila 2
    newSheet.getRange(2, firstCol, 1, lastCol - firstCol + 1).mergeAcross();
  }

  ss.setActiveSheet(newSheet);
  Logger.log(`✅ Pestaña duplicada: ${newSheetName} con ${numCriteriaExtra} columnas adicionales.`);
}


function createStudentTabs() {
  const ss = SpreadsheetApp.getActive();
  const sh1 = ss.getSheetByName('GROUP');
  const templateSheet = ss.getSheetByName('Std Tmpl');

  if (!sh1 || !templateSheet) return;

  const data = sh1.getRange(2, 1, sh1.getLastRow() - 1, 2).getValues(); // Columnas A y B
  const students = data.map(row => ({
    fullName: row[0],      // Columna A
    shortName: row[1]      // Columna B
  }));

  const existingSheets = ss.getSheets().map(sheet => sheet.getName());

  students.forEach((student, index) => {
    const { fullName, shortName } = student;

    if (isValidSheetName(shortName) && !existingSheets.includes(shortName)) {
      try {
        // Copiar plantilla
        const newSheet = templateSheet.copyTo(ss);
        newSheet.setName(shortName);
        const sheetId = newSheet.getSheetId();

        // Añadir enlace clicable en columna B de GROUP
        const cell = sh1.getRange(index + 2, 2);
        const richTextLink = SpreadsheetApp.newRichTextValue()
          .setText(shortName)
          .setLinkUrl(`#gid=${sheetId}`)
          .build();
        cell.setRichTextValue(richTextLink);

        // Copiar texto enriquecido (RichTextValue) desde GROUP!A a nueva hoja A1
        const sourceCell = sh1.getRange(index + 2, 1);
        const targetCell = newSheet.getRange('A1');
        const richTextValue = sourceCell.getRichTextValue();

        if (richTextValue) {
          targetCell.setRichTextValue(richTextValue);
        } else {
          targetCell.setValue(fullName);
        }

      } catch (e) {
        Logger.log(`Error al crear hoja "${shortName}": ${e.message}`);
      }
    } else {
      Logger.log(`Nombre inválido o ya existe: ${shortName}`);
    }
  });
}

function updateStudentTabs() {
  const ss = SpreadsheetApp.getActive();
  const sh1 = ss.getSheetByName('GROUP');
  const templateSheet = ss.getSheetByName('Std Tmpl');

  if (!sh1 || !templateSheet) {
    Logger.log("La hoja 'GROUP' o 'Std Tmpl' no existe.");
    return;
  }

  // Crear diccionario { nombrePestaña: nombreCompleto }
  const data = sh1.getRange(2, 1, sh1.getLastRow() - 1, 2).getValues();
  const studentDict = {};
  data.forEach(row => {
    const fullName = row[0];
    const sheetName = row[1];
    if (sheetName && isValidSheetName(sheetName)) {
      studentDict[sheetName] = fullName;
    }
  });

  const templateHeaders = templateSheet.getRange(1, 1, 1, templateSheet.getLastColumn()).getValues()[0];
  const templateRowLabels = templateSheet.getRange(1, 1, templateSheet.getLastRow(), 1).getValues().flat();

  Object.entries(studentDict).forEach(([sheetName, fullName]) => {
    const studentSheet = ss.getSheetByName(sheetName);
    if (!studentSheet) {
      Logger.log(`No existe la pestaña '${sheetName}'`);
      return;
    }

    const studentHeaders = studentSheet.getRange(1, 1, 1, studentSheet.getLastColumn()).getValues()[0];
    const studentRowLabels = studentSheet.getRange(1, 1, studentSheet.getLastRow(), 1).getValues().flat();

    // --- Añadir columnas que falten ---
    templateHeaders.forEach((header, index) => {
      if (!studentHeaders.includes(header)) {
        const colIndex = studentSheet.getLastColumn() + 1;

        // Copiar valores
        const values = templateSheet.getRange(1, index + 1, templateSheet.getLastRow(), 1).getValues();
        studentSheet.insertColumnAfter(studentSheet.getLastColumn());
        const newColRange = studentSheet.getRange(1, colIndex, templateSheet.getLastRow(), 1);
        newColRange.setValues(values);

        // Copiar fusiones
        const templateMergedRanges = templateSheet.getRange(1, index + 1, templateSheet.getLastRow(), 1).getMergedRanges();
        templateMergedRanges.forEach(r => {
          const rowOffset = r.getRow() - 1;
          const numRows = r.getNumRows();
          studentSheet.getRange(rowOffset + 1, colIndex, numRows, 1).merge();
        });

        // Copiar formato
        templateSheet.getRange(1, index + 1, templateSheet.getLastRow(), 1)
          .copyFormatToRange(studentSheet, colIndex, colIndex, 1, templateSheet.getLastRow());

        Logger.log(`Columna '${header}' añadida en '${sheetName}'`);
      }
    });

    // --- Añadir filas que falten ---
    templateRowLabels.forEach((label, index) => {
      if (!studentRowLabels.includes(label)) {
        const rowIndex = studentSheet.getLastRow() + 1;

        // Copiar valores
        const values = templateSheet.getRange(index + 1, 1, 1, templateSheet.getLastColumn()).getValues();
        studentSheet.insertRowAfter(studentSheet.getLastRow());
        const newRowRange = studentSheet.getRange(rowIndex, 1, 1, templateSheet.getLastColumn());
        newRowRange.setValues(values);

        // Copiar fusiones
        const templateMergedRanges = templateSheet.getRange(index + 1, 1, 1, templateSheet.getLastColumn()).getMergedRanges();
        templateMergedRanges.forEach(r => {
          const colOffset = r.getColumn() - 1;
          const numCols = r.getNumColumns();
          studentSheet.getRange(rowIndex, colOffset + 1, 1, numCols).merge();
        });

        // Copiar formato
        templateSheet.getRange(index + 1, 1, 1, templateSheet.getLastColumn())
          .copyFormatToRange(studentSheet, 1, templateSheet.getLastColumn(), rowIndex, rowIndex);

        Logger.log(`Fila '${label}' añadida en '${sheetName}'`);
      }
    });
  });

  Logger.log("Todas las pestañas fueron sincronizadas correctamente.");
}


function isValidSheetName(name) {
  const invalidChars = /[\\\/\?\*\[\]\:\;]/;
  return !!name && !invalidChars.test(name) && name.length <= 100;
}

function mainSendGrades() {
  sendActivityGradesToMatrix();
}

function sendActivityGradesToMatrix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activitySheet = ss.getActiveSheet();

  const activityId = activitySheet.getRange('A2').getValue();
  if (!activityId) {
    Logger.log('❌ No hay ID de actividad en A2.');
    return;
  }
  
  const lastRow = activitySheet.getLastRow();
  const lastColumn = activitySheet.getLastColumn();

  if (lastRow < 4 || lastColumn < 3) {
    Logger.log('❌ No hay suficientes filas o columnas.');
    return;
  }

  const headers = activitySheet.getRange(3, 3, 1, lastColumn - 2).getValues()[0];

  for (let row = 4; row <= lastRow; row++) {
    const studentId = activitySheet.getRange(row, 2).getValue();
    if (!studentId) continue;

    const grades = activitySheet.getRange(row, 3, 1, headers.length).getValues()[0];
    const studentData = {};

    for (let i = 0; i < headers.length; i++) {
      const criterion = headers[i];
      const grade = grades[i];
      if (criterion && grade !== '') {
        studentData[criterion] = grade;
      }
    }

    if (Object.keys(studentData).length > 0) {
      writeStudentGrades(ss, studentId, studentData, activityId);
    }
  }

  Logger.log('✅ Calificaciones enviadas con éxito.');
}

function writeStudentGrades(ss, studentId, studentData, activityId) {
  const studentSheet = ss.getSheetByName(studentId);
  if (!studentSheet) {
    Logger.log(`❌ No se encontró hoja para el estudiante: ${studentId}`);
    return;
  }

  // Busca el activityId en la fila 1 (donde están las celdas fusionadas verticalmente)
  let activityCol = findActivityColumnVerticalMerge(studentSheet, activityId, 5); // empieza en columna E (5)

  if (activityCol === 0) {
    // No existe la columna, crearla al final
    const lastCol = studentSheet.getLastColumn();
    activityCol = lastCol + 1;
    studentSheet.getRange(1, activityCol).setValue(activityId); // Escribir en fila 1
    studentSheet.getRange(2, activityCol).clearContent();      // Limpiar fila 2
    Logger.log(`🆕 Columna creada para actividad "${activityId}" en hoja de ${studentId}`);
  }

  const lastRow = studentSheet.getLastRow();
  const criteriaRange = studentSheet.getRange('B3:B' + lastRow);
  const criteriaList = criteriaRange.getValues();

  for (const criterion in studentData) {
    const grade = studentData[criterion];
    let rowIndex = -1;

    for (let i = 0; i < criteriaList.length; i++) {
      const sheetCriterion = criteriaList[i][0];
      if (normalize(sheetCriterion) === normalize(criterion)) {
        rowIndex = i + 3;
        break;
      }
    }

    if (rowIndex === -1) {
      Logger.log(`❌ Criterio "${criterion}" no encontrado en hoja de ${studentId}`);
      continue;
    }

    const cell = studentSheet.getRange(rowIndex, activityCol);
    // Sobrescribir sin importar si hay valor anterior
    cell.setValue(grade);
    Logger.log(`✅ ${studentId} → ${criterion}: ${grade} [${rowIndex},${activityCol}]`);
  }
}

function findActivityColumnVerticalMerge(sheet, activityId, startCol) {
  const lastColumn = sheet.getLastColumn();
  const row = 1; // fila donde está el valor en la celda fusionada vertical (ej: E1 y E2 fusionadas)
  const values = sheet.getRange(row, startCol, 1, lastColumn - startCol + 1).getValues()[0];

  for (let i = 0; i < values.length; i++) {
    if (values[i] === activityId) {
      return startCol + i;
    }
  }
  return 0; // no encontrado
}

function normalize(str) {
  return String(str).toLowerCase().trim();
}

function clearActivityGrades() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Paso 1: Solicitar número de actividad
  const response = ui.prompt('Eliminar calificaciones', 'Introduce el número de la actividad (ej. 1 para ACT 1):', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    Logger.log("Operación cancelada por el usuario.");
    return;
  }

  const number = response.getResponseText().trim();
  if (!/^\d+$/.test(number)) {
    ui.alert('Por favor, introduce un número válido.');
    return;
  }

  const activityId = `ACT ${number}`;
  Logger.log(`📌 Eliminando calificaciones para: ${activityId}`);

  // Paso 2: Obtener los nombres de pestañas desde "GROUP"
  const GROUPSheet = ss.getSheetByName('GROUP');
  if (!GROUPSheet) {
    Logger.log('❌ No se encuentra la hoja "GROUP".');
    return;
  }

  const studentNames = GROUPSheet.getRange(2, 2, GROUPSheet.getLastRow() - 1, 1).getValues().flat();
  Logger.log(`📌 Estudiantes encontrados: ${studentNames.join(', ')}`);

  // Paso 3: Recorrer los estudiantes
  studentNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      Logger.log(`⚠️ No se encontró la hoja del estudiante: ${name}`);
      return;
    }

    const headerRow = sheet.getRange(2, 5, 1, sheet.getLastColumn() - 4).getValues()[0]; // Fila 2, desde columna E
    const mergedRanges = sheet.getRange(2, 5, 1, sheet.getLastColumn() - 4).getMergedRanges();

    let targetCol = null;

    // Buscar el ID de actividad en celdas fusionadas primero
    for (let range of mergedRanges) {
      if (range.getValue() === activityId) {
        targetCol = range.getColumn();
        break;
      }
    }

    // Si no está fusionada, buscar en el array de valores
    if (!targetCol) {
      headerRow.forEach((val, idx) => {
        if (val === activityId) {
          targetCol = idx + 5; // +5 porque empezamos en la columna E (índice base 1)
        }
      });
    }

    if (!targetCol) {
      Logger.log(`🔍 Actividad "${activityId}" no encontrada en hoja de ${name}.`);
      return;
    }

    // Borrar calificaciones (desde fila 3 hasta final)
    const lastRow = sheet.getLastRow();
    const targetRange = sheet.getRange(3, targetCol, lastRow - 2, 1);
    targetRange.clearContent();

    Logger.log(`✅ Calificaciones borradas para ${name} en columna ${targetCol}`);
  });

  ui.alert(`✔️ Calificaciones eliminadas para "${activityId}" en todas las hojas.`);
}

function fetchGrades() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaProg = ss.getSheetByName("SYLLABUS");
  const hojaGROUP = ss.getSheetByName("GROUP");

  const lastRowProg = hojaProg.getLastRow();
  if (lastRowProg < 2) {
    Logger.log("La hoja SYLLABUS no tiene datos suficientes.");
    return;
  }

  const rangoColA = hojaProg.getRange("A2:A" + lastRowProg);
  const datosProg = rangoColA.getRichTextValues();
  const fusiones = rangoColA.getMergedRanges();

  let dictOA = {};

  // Buscar celdas fusionadas con texto "OA X"
  fusiones.forEach(rango => {
    const valor = hojaProg.getRange(rango.getRow(), 1).getValue();
    if (typeof valor === "string" && valor.trim().startsWith("OA")) {
      dictOA[valor.trim()] = rango.getRow();
    }
  });

  // Buscar celdas individuales no fusionadas
  for (let i = 0; i < datosProg.length; i++) {
    const valor = datosProg[i][0].getText();
    const fila = i + 2;
    if (valor && valor.startsWith("OA") && !(valor in dictOA)) {
      dictOA[valor] = fila;
    }
  }

  // Leer primera fila de hoja GROUP
  const lastColGROUP = hojaGROUP.getLastColumn();
  const lastRowGROUP = hojaGROUP.getLastRow();
  if (lastColGROUP < 1 || lastRowGROUP < 2) {
    Logger.log("La hoja GROUP no tiene datos suficientes.");
    return;
  }

  const headerGROUP = hojaGROUP.getRange(1, 1, 1, lastColGROUP).getValues()[0];

  for (let col = 0; col < headerGROUP.length; col++) {
    const valor = headerGROUP[col];
    if (typeof valor === "string" && valor.trim().startsWith("OA") && dictOA[valor.trim()]) {
      const filaOA = dictOA[valor.trim()];
      const numFilas = lastRowGROUP - 1;
      const formulas = [];

      for (let i = 2; i <= lastRowGROUP; i++) {
        formulas.push([`=INDIRECT(B${i}&"!C${filaOA}")`]);
      }

      hojaGROUP.getRange(2, col + 1, numFilas, 1).setFormulas(formulas);
    }
  }

  SpreadsheetApp.flush();
  Logger.log("fetchGrades: Proceso completado.");
}


