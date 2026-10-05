const SPREADSHEET_ID = '1pwjPU06bmGqFlxl5TWcP2fuQNdB1YmoSFpAQhflmnOg';

const HOJAS = Object.freeze({
  PARTICIPANTES: 'Participantes',
  SESIONES: 'Sesiones',
  ASISTENCIAS: 'Asistencias'
});

function obtenerLibro_() {
  if (!SPREADSHEET_ID) {
    throw new Error(
      'Configura SPREADSHEET_ID en Config.gs antes de ejecutar la aplicación.'
    );
  }

  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

