function obtenerSesiones() {

  const ss = obtenerLibro_();

  const hoja =
    ss.getSheetByName(
      HOJAS.SESIONES
    );

  if (!hoja || hoja.getLastRow() < 2) {
    return [];
  }

  const datos =
    hoja
      .getRange(
        2,
        1,
        hoja.getLastRow() - 1,
        5
      )
      .getValues();

  const zona =
    Session.getScriptTimeZone();

  return datos.map(fila => ({

    id: fila[0],

    numero: fila[1],

    fecha:
      fila[2] instanceof Date
        ? Utilities.formatDate(
            fila[2],
            zona,
            'yyyy-MM-dd'
          )
        : fila[2],

    tema: fila[3],

    estado: fila[4]

  }));
}


// ======================================
// OBTENER PARTICIPANTES
// ======================================

function obtenerParticipantes(idSesion) {

  if (!idSesion) {
    throw new Error(
      'Debes seleccionar una sesión.'
    );
  }

  const ss = obtenerLibro_();

  const hojaParticipantes =
    ss.getSheetByName(
      HOJAS.PARTICIPANTES
    );

  const hojaAsistencias =
    ss.getSheetByName(
      HOJAS.ASISTENCIAS
    );

  if (!hojaParticipantes) {
    throw new Error(
      'No existe la hoja Participantes.'
    );
  }

  if (!hojaAsistencias) {
    throw new Error(
      'No existe la hoja Asistencias.'
    );
  }

  const participantes =
    hojaParticipantes.getLastRow() < 2

      ? []

      : hojaParticipantes
          .getRange(
            2,
            1,
            hojaParticipantes.getLastRow() - 1,
            4
          )
          .getValues();


  const asistencias =
    hojaAsistencias.getLastRow() < 2

      ? []

      : hojaAsistencias
          .getRange(
            2,
            1,
            hojaAsistencias.getLastRow() - 1,
            4
          )
          .getValues();


  const estadoPorParticipante =
    new Map();


  asistencias.forEach(fila => {

    if (fila[0] === idSesion) {

      estadoPorParticipante.set(
        fila[1],
        fila[2]
      );
    }

  });


  return participantes

    .filter(
      fila => fila[3] === true
    )

    .map(fila => ({

      id: fila[0],

      nombre: fila[1],

      correo: fila[2],

      presente:
        estadoPorParticipante.get(
          fila[0]
        ) === 'PRESENTE'

    }));
}


// ======================================
// GUARDAR ASISTENCIA
// ======================================

function guardarAsistencia(
  idSesion,
  participantes
) {

  if (!idSesion) {

    throw new Error(
      'Debes seleccionar una sesión.'
    );

  }


  if (
    !Array.isArray(participantes) ||
    participantes.length === 0
  ) {

    throw new Error(
      'No se recibieron participantes.'
    );

  }


  const ss = obtenerLibro_();


  const hoja =
    ss.getSheetByName(
      HOJAS.ASISTENCIAS
    );


  if (!hoja) {

    throw new Error(
      'No existe la hoja Asistencias.'
    );

  }


  const existentes =
    hoja.getLastRow() < 2

      ? []

      : hoja
          .getRange(
            2,
            1,
            hoja.getLastRow() - 1,
            4
          )
          .getValues();


  /*
   * Mapa para evitar duplicados.
   *
   * Ejemplo:
   *
   * S01|P01
   * S01|P02
   * S02|P01
   */

  const indice =
    new Map();


  existentes.forEach(
    (fila, i) => {

      const clave =
        `${fila[0]}|${fila[1]}`;

      indice.set(
        clave,
        i
      );

    }
  );


  const ahora =
    new Date();


  participantes.forEach(p => {

    const clave =
      `${idSesion}|${p.id}`;


    const registro = [

      idSesion,

      p.id,

      p.presente
        ? 'PRESENTE'
        : 'AUSENTE',

      ahora

    ];


    if (indice.has(clave)) {

      // Actualizar

      existentes[
        indice.get(clave)
      ] = registro;

    }

    else {

      // Insertar

      indice.set(
        clave,
        existentes.length
      );

      existentes.push(
        registro
      );

    }

  });


  // Limpiar registros anteriores

  if (hoja.getLastRow() > 1) {

    hoja
      .getRange(
        2,
        1,
        hoja.getLastRow() - 1,
        4
      )
      .clearContent();

  }


  // Volver a escribir registros

  if (existentes.length > 0) {

    hoja
      .getRange(
        2,
        1,
        existentes.length,
        4
      )
      .setValues(existentes);


    hoja
      .getRange(
        2,
        4,
        existentes.length,
        1
      )
      .setNumberFormat(
        'yyyy-mm-dd hh:mm:ss'
      );

  }


  marcarSesionRegistrada_(
    idSesion
  );


  const presentes =
    participantes
      .filter(
        p => p.presente
      )
      .length;


  return {

    ok: true,

    idSesion: idSesion,

    presentes: presentes,

    ausentes:
      participantes.length -
      presentes,

    total:
      participantes.length

  };

}


// ======================================
// CAMBIAR ESTADO DE SESIÓN
// ======================================

function marcarSesionRegistrada_(
  idSesion
) {

  const ss =
    obtenerLibro_();


  const hoja =
    ss.getSheetByName(
      HOJAS.SESIONES
    );


  if (
    !hoja ||
    hoja.getLastRow() < 2
  ) {

    return;

  }


  const datos =
    hoja
      .getRange(
        2,
        1,
        hoja.getLastRow() - 1,
        5
      )
      .getValues();


  const posicion =
    datos.findIndex(
      fila =>
        fila[0] === idSesion
    );


  if (posicion >= 0) {

    hoja
      .getRange(
        posicion + 2,
        5
      )
      .setValue(
        'REGISTRADA'
      );

  }

}