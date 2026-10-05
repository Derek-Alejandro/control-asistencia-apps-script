function obtenerSesiones() {
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


// =========================================
// OBTENER PARTICIPANTES
// =========================================

function obtenerParticipantes(
  idSesion
) {
  if (!idSesion) {
    throw new Error(
      'Debes seleccionar una sesión.'
    );
  }

  const ss =
    obtenerLibro_();

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
            7
          )
          .getValues();

  const estadoPorParticipante =
    new Map();

  asistencias.forEach(fila => {
    if (fila[0] === idSesion) {
      estadoPorParticipante.set(
        fila[1],
        {
          estado:
            fila[2] || 'AUSENTE',

          minutosRetardo:
            Number(fila[4]) || 0,

          justificado:
            fila[5] === true,

          observacion:
            fila[6] || ''
        }
      );
    }
  });

  return participantes
    .filter(
      fila => fila[3] === true
    )
    .map(fila => {
      const registro =
        estadoPorParticipante.get(
          fila[0]
        );

      return {
        id: fila[0],

        nombre: fila[1],

        correo: fila[2],

        estado:
          registro
            ? registro.estado
            : 'AUSENTE',

        minutosRetardo:
          registro
            ? registro.minutosRetardo
            : 0,

        justificado:
          registro
            ? registro.justificado
            : false,

        observacion:
          registro
            ? registro.observacion
            : ''
      };
    });
}


// =========================================
// GUARDAR ASISTENCIA
// =========================================

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

  const estadosPermitidos =
    [
      'PRESENTE',
      'AUSENTE',
      'RETARDO'
    ];

  participantes.forEach(p => {
    if (
      !p.id ||
      !estadosPermitidos.includes(
        p.estado
      )
    ) {
      throw new Error(
        `Datos inválidos para el participante ${p.id || 'desconocido'}.`
      );
    }
  });

  // Evita que dos usuarios escriban
  // al mismo tiempo.
  const lock =
    LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const ss =
      obtenerLibro_();

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
              7
            )
            .getValues();

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

      let minutosRetardo =
        Number(
          p.minutosRetardo
        ) || 0;

      if (
        p.estado !==
        'RETARDO'
      ) {
        minutosRetardo = 0;
      }

      if (
        minutosRetardo < 0
      ) {
        minutosRetardo = 0;
      }

      const registro = [
        idSesion,
        p.id,
        p.estado,
        ahora,
        minutosRetardo,
        Boolean(
          p.justificado
        ),
        String(
          p.observacion || ''
        ).trim()
      ];

      if (indice.has(clave)) {
        existentes[
          indice.get(clave)
        ] = registro;
      } else {
        indice.set(
          clave,
          existentes.length
        );

        existentes.push(
          registro
        );
      }
    });

    if (
      hoja.getLastRow() > 1
    ) {
      hoja
        .getRange(
          2,
          1,
          hoja.getLastRow() - 1,
          7
        )
        .clearContent();
    }

    if (
      existentes.length > 0
    ) {
      hoja
        .getRange(
          2,
          1,
          existentes.length,
          7
        )
        .setValues(
          existentes
        );

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
      participantes.filter(
        p =>
          p.estado ===
          'PRESENTE'
      ).length;

    const retardos =
      participantes.filter(
        p =>
          p.estado ===
          'RETARDO'
      ).length;

    const ausentes =
      participantes.filter(
        p =>
          p.estado ===
          'AUSENTE'
      ).length;

    const justificados =
      participantes.filter(
        p =>
          p.justificado
      ).length;

    return {
      ok: true,

      idSesion,

      total:
        participantes.length,

      presentes,

      retardos,

      ausentes,

      justificados
    };

  } finally {
    lock.releaseLock();
  }
}


// =========================================
// MARCAR SESIÓN COMO REGISTRADA
// =========================================

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