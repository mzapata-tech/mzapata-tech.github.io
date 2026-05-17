// ============================================================
// ⚠️ STAGING — NO ES PRODUCCIÓN
// ============================================================
const IS_STAGING = true; // ← cambiar a false en producción

const TWILIO_ACCOUNT_SID = "ACxxxxxxxxxxxxxxxxxxxx";
const TWILIO_AUTH_TOKEN = "xxxxxxxxxxxxxxxxxxxxxxxxx";
const TWILIO_WHATSAPP_FROM = "whatsapp:+14155238886"; // número sandbox Twilio

const OPENAI_API_KEY = "sk-xxxxxxxxxxxxxxxxxxxxxxx";

// ============================================================
// CLÍNICA INTEGRAL MEDIVEINS - Sistema de Citas
// Apps Script para Google Sheets
// Archivo: Code.gs
// ============================================================

const CONFIG = {
  SHEET_CITAS: "CITAS",
  SHEET_DOCTORES: "DOCTORES",
  SHEET_SERVICIOS: "SERVICIOS",
  SHEET_PACIENTES: "PACIENTES",
  NOMBRE_CLINICA: "Clínica Integral Mediveins",
  EMAIL_ADMIN: "clinicaintegralmediveins@gmail.com", // Completar con email de la secretaria/admin
  LINK_UBICACION: "https://waze.com/ul/hd1u0x01yd",
};

// ============================================================
// MENÚ Y APERTURA
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🏥 Clínica Mediveins")
    .addItem("📋 Abrir Panel de Citas", "mostrarPanel")
    .addSeparator()
    .addItem("⚙️ Configurar sistema (primera vez)", "setupSheets")
    .addItem("⏰ Activar recordatorios diarios", "configurarTriggerDiario")
    .addSeparator()
    .addItem("📊 Ver resumen del día", "mostrarResumenHoy")
    .addItem("🔄 Activar sync Calendar → Sheets", "configurarTriggerCalendar")
    .addSeparator()
    .addItem(
      "📢 Exportar pacientes para campaña",
      "exportarPacientesParaCampana",
    )
    .addItem("🧹 Limpiar y exportar campaña", "limpiarYExportarCampana")
    .addItem("🔒 Proteger formato de Pacientes", "protegerHojaPacientes")
    .addToUi();
}

function mostrarPanel() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("Panel de Citas · Mediveins")
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ============================================================
// CONFIGURACIÓN INICIAL DE HOJAS
// ============================================================

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const definicion = {
    CITAS: {
      headers: [
        "ID",
        "Fecha",
        "Hora",
        "Doctor",
        "Especialidad",
        "Paciente",
        "Teléfono",
        "Servicio",
        "Precio (₡)",
        "Estado",
        "Notas",
        "En Calendar",
        "Event ID",
        "Fecha Creación",
      ],
      anchos: [
        80, 100, 70, 160, 140, 180, 110, 180, 100, 110, 200, 90, 120, 120,
      ],
    },
    DOCTORES: {
      headers: [
        "Nombre",
        "Especialidad",
        "Teléfono",
        "Calendar ID",
        "Usa Calendar (Sí/No)",
        "Color",
      ],
      anchos: [200, 160, 120, 260, 150, 100],
    },
    SERVICIOS: {
      headers: ["Doctor", "Servicio", "Precio (₡)", "Duración (min)"],
      anchos: [200, 220, 120, 130],
    },
    PACIENTES: {
      headers: ["Nombre", "Teléfono", "Email", "Notas", "Última visita"],
      anchos: [200, 120, 200, 240, 120],
    },
  };

  for (const [nombre, def] of Object.entries(definicion)) {
    let sheet = ss.getSheetByName(nombre);
    if (!sheet) sheet = ss.insertSheet(nombre);

    const hRange = sheet.getRange(1, 1, 1, def.headers.length);
    hRange.setValues([def.headers]);
    hRange.setBackground("#1a73e8");
    hRange.setFontColor("#ffffff");
    hRange.setFontWeight("bold");
    hRange.setFontSize(11);
    sheet.setFrozenRows(1);

    def.anchos.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }

  // Datos de ejemplo para DOCTORES
  const shDoc = ss.getSheetByName("DOCTORES");
  if (shDoc.getLastRow() <= 1) {
    shDoc.getRange(2, 1, 3, 6).setValues([
      ["Dr. García Mora", "Cardiología", "8888-1111", "", "No", "#4285f4"],
      ["Dra. López Vargas", "Pediatría", "8888-2222", "", "No", "#34a853"],
      [
        "Dr. Ramírez Solís",
        "Medicina General",
        "8888-3333",
        "",
        "Sí",
        "#ea4335",
      ],
    ]);
  }

  // Datos de ejemplo para SERVICIOS
  const shSer = ss.getSheetByName("SERVICIOS");
  if (shSer.getLastRow() <= 1) {
    shSer.getRange(2, 1, 6, 4).setValues([
      ["Dr. García Mora", "Consulta Cardiología", 35000, 30],
      ["Dr. García Mora", "Electrocardiograma", 15000, 20],
      ["Dra. López Vargas", "Consulta Pediátrica", 25000, 30],
      ["Dra. López Vargas", "Control de crecimiento", 20000, 20],
      ["Dr. Ramírez Solís", "Consulta General", 20000, 20],
      ["Dr. Ramírez Solís", "Certificado médico", 10000, 15],
    ]);
  }

  ui.alert(
    "✅ Sistema configurado correctamente.\n\nAhora:\n1. Abra el menú 🏥 Clínica Mediveins\n2. Complete los doctores en la hoja DOCTORES\n3. Complete los servicios en la hoja SERVICIOS\n4. Abra el Panel de Citas para comenzar",
  );
}

// ============================================================
// FUNCIONES DE DATOS — LECTURAS
// ============================================================

function getDoctores() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.SHEET_DOCTORES,
  );
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 6)
    .getValues()
    .filter((r) => r[0])
    .map((r) => ({
      nombre: r[0],
      especialidad: r[1],
      telefono: r[2],
      calendarId: r[3],
      usaCalendar: String(r[4]).toLowerCase() === "sí",
      color: r[5] || "#1a73e8",
    }));
}

function getServiciosPorDoctor(nombreDoctor) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.SHEET_SERVICIOS,
  );
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 4)
    .getValues()
    .filter((r) => r[0] === nombreDoctor)
    .map((r) => ({ servicio: r[1], precio: r[2], duracion: r[3] }));
}

function getPacientesSugeridos(query) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.SHEET_PACIENTES,
  );
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const q = String(query).toLowerCase();
  return sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 2)
    .getValues()
    .filter(
      (r) =>
        r[0] &&
        (String(r[0]).toLowerCase().includes(q) || String(r[1]).includes(q)),
    )
    .map((r) => ({ nombre: r[0], telefono: r[1] }))
    .slice(0, 6);
}

function getCitas(filtros) {
  filtros = filtros || {};
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.SHEET_CITAS,
  );
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();

  let citas = data
    .map((row, i) => {
      const fechaRaw = row[1];
      let fechaStr = "";
      if (fechaRaw instanceof Date && !isNaN(fechaRaw)) {
        fechaStr = Utilities.formatDate(
          fechaRaw,
          "America/Costa_Rica",
          "yyyy-MM-dd",
        );
      } else if (typeof fechaRaw === "string" && fechaRaw) {
        fechaStr = fechaRaw.substring(0, 10);
      }

      return {
        rowIndex: i + 2,
        id: row[0],
        fecha: fechaStr,
        hora:
          row[2] instanceof Date
            ? Utilities.formatDate(row[2], "America/Costa_Rica", "HH:mm")
            : row[2]
              ? String(row[2]).substring(0, 5)
              : "",
        doctor: row[3],
        especialidad: row[4],
        paciente: row[5],
        telefono: row[6],
        servicio: row[7],
        precio: row[8] || 0,
        estado: row[9] || "Pendiente",
        notas: row[10] || "",
        enCalendar: row[11],
        eventId: row[12],
      };
    })
    .filter((c) => c.id !== "" && c.id !== undefined);

  if (filtros.fecha) citas = citas.filter((c) => c.fecha === filtros.fecha);
  if (filtros.doctor) citas = citas.filter((c) => c.doctor === filtros.doctor);
  if (filtros.estado) citas = citas.filter((c) => c.estado === filtros.estado);
  if (filtros.buscar) {
    const q = String(filtros.buscar).toLowerCase();
    citas = citas.filter(
      (c) =>
        String(c.paciente).toLowerCase().includes(q) ||
        String(c.telefono).includes(q),
    );
  }

  return citas.sort((a, b) => {
    if (a.fecha !== b.fecha) return a.fecha.localeCompare(b.fecha);
    return String(a.hora).localeCompare(String(b.hora));
  });
}

// ============================================================
// FUNCIONES DE DATOS — ESCRITURAS
// ============================================================

function agregarCita(cita) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_CITAS);

    // Obtener especialidad del doctor
    const doctores = getDoctores();
    const doctorObj = doctores.find((d) => d.nombre === cita.doctor);
    const especialidad = doctorObj ? doctorObj.especialidad : "";

    const id = "C" + new Date().getTime();
    const fecha = new Date(cita.fecha + "T12:00:00-06:00");
    const ahora = new Date();

    const row = [
      id,
      fecha,
      cita.hora,
      cita.doctor,
      especialidad,
      cita.paciente,
      cita.telefono,
      cita.servicio,
      cita.precio,
      "Pendiente",
      cita.notas || "",
      "No",
      "",
      ahora,
    ];

    sheet.appendRow(row);

    // Color inicial (pendiente = amarillo claro)
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, 14).setBackground("#fff8e1");

    // Guardar / actualizar paciente
    guardarPaciente(cita.paciente, cita.telefono);

    // Intentar sincronizar con Google Calendar
    if (doctorObj && doctorObj.usaCalendar && doctorObj.calendarId) {
      cita.calendarId = doctorObj.calendarId;
      const eventId = sincronizarConCalendar(cita, id, especialidad);
      if (eventId) {
        sheet.getRange(lastRow, 12).setValue("Sí");
        sheet.getRange(lastRow, 13).setValue(eventId);
      }
    }

    return { success: true, id: id };
  } catch (e) {
    Logger.log("Error agregarCita: " + e.message);
    return { success: false, error: e.message };
  }
}

function actualizarEstadoCita(rowIndex, nuevoEstado) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      CONFIG.SHEET_CITAS,
    );
    sheet.getRange(rowIndex, 10).setValue(nuevoEstado);

    const colores = {
      Confirmada: "#e6f4ea",
      Pendiente: "#fff8e1",
      Cancelada: "#fce8e6",
      Completada: "#e8f0fe",
    };
    sheet
      .getRange(rowIndex, 1, 1, 14)
      .setBackground(colores[nuevoEstado] || "#ffffff");
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function guardarPaciente(nombre, telefono) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      CONFIG.SHEET_PACIENTES,
    );
    const data =
      sheet.getLastRow() > 1
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues()
        : [];

    const existe = data.some(
      (r) => String(r[0]).toLowerCase() === String(nombre).toLowerCase(),
    );

    if (!existe && nombre) {
      sheet.appendRow([
        nombre,
        telefono,
        "",
        "",
        Utilities.formatDate(new Date(), "America/Costa_Rica", "yyyy-MM-dd"),
      ]);
    } else if (existe) {
      const idx = data.findIndex(
        (r) => String(r[0]).toLowerCase() === String(nombre).toLowerCase(),
      );
      if (idx !== -1) {
        if (telefono && !data[idx][1])
          sheet.getRange(idx + 2, 2).setValue(telefono);
        sheet
          .getRange(idx + 2, 5)
          .setValue(
            Utilities.formatDate(
              new Date(),
              "America/Costa_Rica",
              "yyyy-MM-dd",
            ),
          );
      }
    }

    // ── Sincronizar con Google Contacts ──────────────────────────
    sincronizarContactoWhatsApp(nombre, telefono);
  } catch (e) {
    Logger.log("Error guardarPaciente: " + e.message);
  }
}

// ============================================================
// GOOGLE CALENDAR (solo para doctores que lo usan)
// ============================================================

function sincronizarConCalendar(cita, citaId, especialidad) {
  try {
    const calendar = CalendarApp.getCalendarById(cita.calendarId || "");
    if (!calendar) {
      Logger.log("Calendar no encontrado: " + (cita.calendarId || "sin ID"));
      return null;
    }

    // ✅ Construcción correcta de fecha y hora
    const [y, mo, d] = cita.fecha.split("-").map(Number);
    const [hh, mm] = String(cita.hora).split(":").map(Number);
    const inicio = new Date(y, mo - 1, d, hh, mm, 0);
    const duracion = cita.duracion || 30; // minutos
    const fin = new Date(inicio.getTime() + duracion * 60000);

    const evento = calendar.createEvent(
      `${cita.paciente} · ${cita.servicio}`,
      inicio,
      fin,
      {
        description:
          `Paciente: ${cita.paciente}\n` +
          `Teléfono: ${cita.telefono}\n` +
          `Servicio: ${cita.servicio}\n` +
          `Precio: ₡${Number(cita.precio).toLocaleString()}\n` +
          `Especialidad: ${especialidad}\n` +
          `ID Sistema: ${citaId}`,
        location: CONFIG.NOMBRE_CLINICA,
      },
    );
    return evento.getId();
  } catch (e) {
    Logger.log("Error Calendar sync: " + e.message);
    return null;
  }
}

// ============================================================
// GENERADOR DE MENSAJES WHATSAPP
// ============================================================

// ============================================================
// UTILIDAD — Convierte "HH:mm" (24h) a "h:mm a.m. / p.m." (12h)
// ============================================================
function formatearHora12(hora24) {
  if (!hora24) return "";
  const [hh, mm] = String(hora24).split(":").map(Number);
  const periodo = hh < 12 ? "a.m." : "p.m.";
  const horas12 = hh % 12 || 12; // 0 → 12, 13 → 1, etc.
  const minutos = mm === 0 ? "" : `:${String(mm).padStart(2, "0")}`;
  return `${horas12}${minutos} ${periodo}`; // ej: "9 a.m." · "2:30 p.m."
}

function generarMensajeWhatsApp(cita, tipo) {
  tipo = tipo || "confirmacion";

  // ── Parseo robusto de fecha ──────────────────────────────────────────────
  // cita.fecha puede llegar como 'yyyy-MM-dd', objeto Date, string largo, etc.
  let fechaObj;
  const raw = cita.fecha || "";

  if (raw instanceof Date) {
    fechaObj = raw;
  } else if (typeof raw === "string" && /^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    // Desglosamos manualmente para evitar desfases de zona horaria (UTC vs CR)
    const [y, m, d] = raw.split("-").map(Number);
    fechaObj = new Date(y, m - 1, d); // fecha local, sin conversión UTC
  } else {
    fechaObj = new Date(raw); // último recurso
  }

  if (isNaN(fechaObj.getTime())) fechaObj = new Date(); // fallback seguro

  const diasSemana = [
    "domingo",
    "lunes",
    "martes",
    "miércoles",
    "jueves",
    "viernes",
    "sábado",
  ];
  const meses = [
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
  ];

  const dia = diasSemana[fechaObj.getDay()];
  const numDia = fechaObj.getDate();
  const mes = meses[fechaObj.getMonth()];
  const anio = fechaObj.getFullYear();
  const fechaLegible = `${dia} ${numDia} de ${mes} de ${anio}`;

  const nombre = String(cita.paciente || "").split(" ")[0];
  const hora12 = formatearHora12(cita.hora); // ← línea nueva
  const servicioFinal = String(cita.servicio || "").includes("/")
    ? cita.servicio
    : String(cita.especialidad || "")
        .replace("Médico general", "Med.General")
        .replace("Medicina general", "Med.General")
        .replace("Medicina General", "Med.General")
        .replace("Cardiología", "Cardiol.")
        .replace("Pediatría", "Pediatr.")
        .replace("Fisiatría", "Fisiatr.") +
      "/" +
      cita.servicio;

  const precioFormato = String(Number(cita.precio || 0)).replace(
    /\B(?=(\d{3})+(?!\d))/g,
    ".",
  );

  const plantillas = {
    confirmacion:
      `Hola ${nombre}, le confirmamos su cita en *${CONFIG.NOMBRE_CLINICA}*:\n\n` +
      `📅 *Fecha:* ${fechaLegible}\n` +
      `⏰ *Hora:* ${hora12}\n` +
      `*Doctor:* ${cita.doctor}\n` +
      `*Servicio:* ${servicioFinal} ¢${precioFormato}\n\n` +
      `📍 *¿Cómo llegar?* → ${CONFIG.LINK_UBICACION}\n\n` +
      `Agradecemos confirme su asistencia respondiendo a este mensaje\n\n` +
      `¡Le esperamos!`,

    recordatorio:
      `Buenos días ${nombre}, le recordamos su cita en *${CONFIG.NOMBRE_CLINICA}*:\n\n` +
      `📅 *Fecha:* ${fechaLegible}\n` +
      `⏰ *Hora:* ${hora12}\n` +
      `*Doctor:* ${cita.doctor}\n` +
      `*Servicio:* ${servicioFinal} ¢${precioFormato}\n\n` +
      `📍 *¿Cómo llegar?* → ${CONFIG.LINK_UBICACION}\n\n` +
      `Si necesita reprogramar, contáctenos.`,

    cancelacion:
      `Hola ${nombre}, lamentamos informarle que su cita del *${fechaLegible}* a las *${hora12}* con ${cita.doctor} ha tenido que ser cancelada.\n\n` +
      `Por favor contáctenos para reagendarla. Disculpe los inconvenientes 🙏\n\n` +
      `*${CONFIG.NOMBRE_CLINICA}*`,

    agendada:
      `Hola ${nombre}, su cita ha sido agendada de manera exitosa en *${CONFIG.NOMBRE_CLINICA}*:\n\n` +
      `📅 *Fecha:* ${fechaLegible}\n` +
      `⏰ *Hora:* ${hora12}\n` +
      `*Doctor:* ${cita.doctor}\n` +
      `*Servicio:* ${servicioFinal} ¢${precioFormato}\n\n` +
      `📍 *¿Cómo llegar?* → ${CONFIG.LINK_UBICACION}\n\n` +
      `Si tiene alguna consulta no dude en contactarnos. ¡Le esperamos!`,
  };

  return plantillas[tipo] || plantillas.confirmacion;
}

function generarMensajesLote(filtros, tipo) {
  const citas = getCitas(filtros);
  return citas.map((c) => ({
    paciente: c.paciente,
    telefono: c.telefono,
    doctor: c.doctor,
    hora: c.hora,
    fecha: c.fecha,
    estado: c.estado,
    mensaje: generarMensajeWhatsApp(c, tipo),
  }));
}

// ============================================================
// RECORDATORIOS AUTOMÁTICOS (trigger diario 8am)
// ============================================================

function verificarRecordatorios() {
  const tz = "America/Costa_Rica";
  const manana = new Date();
  manana.setDate(manana.getDate() + 1);
  const fechaManana = Utilities.formatDate(manana, tz, "yyyy-MM-dd");

  const citasManana = getCitas({ fecha: fechaManana });

  // Resumen en el Sheet (pestaña o celda especial)
  registrarLogDiario(citasManana, fechaManana);

  // Email resumen
  const email = CONFIG.EMAIL_ADMIN || Session.getActiveUser().getEmail();

  if (!email) return;

  const total = citasManana.length;
  if (total === 0) {
    MailApp.sendEmail({
      to: email,
      subject: `🏥 Mediveins · No hay citas para mañana ${fechaManana}`,
      body: `No hay citas registradas para mañana ${fechaManana}.\n\nSistema de Citas · ${CONFIG.NOMBRE_CLINICA}`,
    });
    return;
  }

  const lineas = citasManana
    .map(
      (c) =>
        `  • ${c.hora}  ${c.paciente.padEnd(25)}  ${c.doctor}  [${c.estado}]`,
    )
    .join("\n");

  MailApp.sendEmail({
    to: email,
    subject: `🏥 Mediveins · ${total} cita(s) para mañana ${fechaManana}`,
    htmlBody: `
      <h2 style="color:#1a73e8;">Citas para mañana ${fechaManana}</h2>
      <p>Hay <strong>${total}</strong> cita(s) programada(s):</p>
      <table style="border-collapse:collapse;width:100%;font-family:Arial,sans-serif;font-size:14px;">
        <tr style="background:#1a73e8;color:#fff;">
          <th style="padding:8px 12px;text-align:left;">Hora</th>
          <th style="padding:8px 12px;text-align:left;">Paciente</th>
          <th style="padding:8px 12px;text-align:left;">Doctor</th>
          <th style="padding:8px 12px;text-align:left;">Servicio</th>
          <th style="padding:8px 12px;text-align:left;">Estado</th>
        </tr>
        ${citasManana
          .map(
            (c, i) => `
        <tr style="background:${i % 2 === 0 ? "#f8f9fa" : "#fff"};">
          <td style="padding:8px 12px;">${c.hora}</td>
          <td style="padding:8px 12px;">${c.paciente}</td>
          <td style="padding:8px 12px;">${c.doctor}</td>
          <td style="padding:8px 12px;">${c.servicio}</td>
          <td style="padding:8px 12px;">${c.estado}</td>
        </tr>`,
          )
          .join("")}
      </table>
      <br>
      <p style="color:#5f6368;font-size:12px;">Sistema de Citas · ${CONFIG.NOMBRE_CLINICA}</p>
    `,
  });
}

function registrarLogDiario(citas, fecha) {
  // Opcional: crear hoja LOG para historial
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("LOG");
  if (!logSheet) {
    logSheet = ss.insertSheet("LOG");
    logSheet
      .getRange(1, 1, 1, 3)
      .setValues([["Fecha", "Hora registro", "Citas mañana"]]);
    logSheet
      .getRange(1, 1, 1, 3)
      .setBackground("#5f6368")
      .setFontColor("#fff")
      .setFontWeight("bold");
  }
  logSheet.appendRow([fecha, new Date(), citas.length]);
}

function mostrarResumenHoy() {
  const tz = "America/Costa_Rica";
  const hoy = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const citas = getCitas({ fecha: hoy });

  if (citas.length === 0) {
    SpreadsheetApp.getUi().alert(`📭 No hay citas registradas para hoy ${hoy}`);
    return;
  }

  const resumen = citas
    .map((c) => `${c.hora}  ${c.paciente}  →  ${c.doctor}  [${c.estado}]`)
    .join("\n");

  SpreadsheetApp.getUi().alert(
    `📋 Citas de hoy (${hoy}) — Total: ${citas.length}\n\n${resumen}`,
  );
}

function configurarTriggerDiario() {
  // Eliminar triggers previos de esta función para no duplicar
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === "verificarRecordatorios")
    .forEach((t) => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("verificarRecordatorios")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .inTimezone("America/Costa_Rica")
    .create();

  SpreadsheetApp.getUi().alert(
    "✅ Recordatorios activados.\n\n" +
      "Cada mañana a las 8:00am recibirá un correo con el resumen de citas del día siguiente.\n\n" +
      "Asegúrese de que el campo EMAIL_ADMIN en Code.gs tenga el correo correcto.",
  );
}

// ============================================================
// SINCRONIZACIÓN CALENDAR → SHEETS (trigger cada 15 min)
// ============================================================

function importarCitasDesdeCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_CITAS);
  const doctores = getDoctores().filter((d) => d.usaCalendar && d.calendarId);

  if (doctores.length === 0) return;

  // Rango de búsqueda: desde hoy hasta 60 días adelante
  const ahora = new Date();
  const limite = new Date();
  limite.setDate(limite.getDate() + 60);

  // IDs de citas ya registradas en Sheets (columna 1)
  const existentes =
    sheet.getLastRow() > 1
      ? sheet
          .getRange(2, 13, sheet.getLastRow() - 1, 1)
          .getValues()
          .flat()
          .filter(Boolean)
      : [];

  doctores.forEach((doctor) => {
    try {
      const calendar = CalendarApp.getCalendarById(doctor.calendarId);
      if (!calendar) return;

      const eventos = calendar.getEvents(ahora, limite);

      eventos.forEach((evento) => {
        const eventId = evento.getId();

        // Si ya existe en Sheets, no duplicar
        if (existentes.includes(eventId)) return;

        // Leer descripción para extraer datos
        const desc = evento.getDescription() || "";
        const tituloPartes = evento.getTitle().split(" · ");
        const paciente =
          extraerCampo(desc, "Paciente") ||
          tituloPartes[0] ||
          evento.getTitle();
        const telefono = extraerCampo(desc, "Teléfono") || "";
        const servicio =
          extraerCampo(desc, "Servicio") || tituloPartes[1] || "";
        const precio = extraerCampo(desc, "Precio") || "";
        const idSist = extraerCampo(desc, "ID Sistema");

        // Si ya tiene ID del sistema, fue creado desde Sheets → omitir
        if (idSist) return;

        const inicio = evento.getStartTime();
        const fecha = Utilities.formatDate(
          inicio,
          "America/Costa_Rica",
          "yyyy-MM-dd",
        );
        const hora = Utilities.formatDate(
          inicio,
          "America/Costa_Rica",
          "HH:mm",
        );
        const id =
          "CAL" + new Date().getTime() + Math.random().toString(36).slice(2, 5);

        const fila = [
          id,
          new Date(fecha + "T12:00:00-06:00"),
          hora,
          doctor.nombre,
          doctor.especialidad,
          paciente,
          telefono,
          servicio,
          precio.replace(/[₡,]/g, "") || 0,
          "Confirmada",
          "",
          "Sí",
          eventId,
          new Date(),
        ];

        sheet.appendRow(fila);
        sheet.getRange(sheet.getLastRow(), 1, 1, 14).setBackground("#e6f4ea");
        guardarPaciente(paciente, telefono);
      });
    } catch (e) {
      Logger.log("Error importar calendar " + doctor.nombre + ": " + e.message);
    }
  });
}

// Helper para extraer campos de la descripción del evento
function extraerCampo(texto, campo) {
  const regex = new RegExp(campo + ":\\s*(.+)", "i");
  const match = texto.match(regex);
  return match ? match[1].trim() : "";
}

function configurarTriggerCalendar() {
  // Eliminar triggers previos para no duplicar
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === "importarCitasDesdeCalendar")
    .forEach((t) => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("importarCitasDesdeCalendar")
    .timeBased()
    .everyMinutes(15)
    .create();

  SpreadsheetApp.getUi().alert(
    "✅ Sincronización activada.\n\n" +
      "Cada 15 minutos el sistema revisará Google Calendar\n" +
      "y agregará automáticamente las citas nuevas al Sheets.",
  );
}

// ============================================================
// SINCRONIZACIÓN CON GOOGLE CONTACTS (para WhatsApp Web/Business)
// ============================================================

function sincronizarContactoWhatsApp(nombre, telefono) {
  if (!nombre || !telefono) return;

  try {
    const telefonoLimpio = String(telefono).replace(/[\s\-\(\)]/g, "");
    const telefonoIntl = telefonoLimpio.startsWith("+")
      ? telefonoLimpio
      : "+506" + telefonoLimpio;

    // Buscar si ya existe por teléfono
    const contactos = ContactsApp.getContactsByPhone(
      telefonoIntl,
      ContactsApp.Field.MOBILE_PHONE,
    );

    if (contactos && contactos.length > 0) {
      const nombreActual = contactos[0].getFullName();
      if (nombreActual.toLowerCase() !== nombre.toLowerCase()) {
        contactos[0].setFullName(nombre);
        Logger.log("✅ Contacto actualizado: " + nombre);
      } else {
        Logger.log("ℹ️ Contacto ya existe: " + nombre);
      }
      return;
    }

    // No existe — crear nuevo
    const nuevoContacto = ContactsApp.createContact(
      nombre.split(" ")[0],
      nombre.split(" ").slice(1).join(" "),
      "",
    );

    nuevoContacto.addPhone(ContactsApp.Field.MOBILE_PHONE, telefonoIntl);

    const grupo = obtenerOCrearGrupoPacientes();
    if (grupo) grupo.addContact(nuevoContacto);

    Logger.log("✅ Contacto creado: " + nombre + " · " + telefonoIntl);
  } catch (e) {
    Logger.log(
      "❌ Error contacto: " + e.message + " | " + nombre + " | " + telefono,
    );
  }
}

function obtenerOCrearGrupoPacientes() {
  const NOMBRE_GRUPO = "Pacientes Mediveins";
  try {
    const grupos = ContactsApp.getContactGroups();
    const existente = grupos.find((g) => g.getName() === NOMBRE_GRUPO);
    if (existente) return existente;

    const nuevo = ContactsApp.createContactGroup(NOMBRE_GRUPO);
    Logger.log("✅ Grupo creado: " + NOMBRE_GRUPO);
    return nuevo;
  } catch (e) {
    Logger.log("❌ Error grupo contactos: " + e.message);
    return null;
  }
}

//campanas de whatsapp

function exportarPacientesParaCampana() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPacientes = ss.getSheetByName(CONFIG.SHEET_PACIENTES);

  if (!shPacientes || shPacientes.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("No hay pacientes registrados.");
    return;
  }

  const pacientes = shPacientes
    .getRange(2, 1, shPacientes.getLastRow() - 1, 5)
    .getValues()
    .filter((r) => r[0]); // solo requiere nombre

  let expSheet = ss.getSheetByName("CAMPAÑA");
  if (!expSheet) expSheet = ss.insertSheet("CAMPAÑA");
  expSheet.clearContents();

  expSheet
    .getRange(1, 1, 1, 5)
    .setValues([
      ["Nombre", "Teléfono", "Teléfono Internacional", "Link WhatsApp", "Lote"],
    ])
    .setBackground("#1a73e8")
    .setFontColor("#fff")
    .setFontWeight("bold");

  const filas = pacientes.map((p, i) => {
    const telLimpio = String(p[1] || "").replace(/[\s\-\(\)]/g, "");
    const telIntl = telLimpio
      ? telLimpio.startsWith("+")
        ? telLimpio
        : "+506" + telLimpio
      : "";
    const lote = i < 256 ? "Lista 1" : "Lista 2";
    return [
      p[0],
      p[1],
      telIntl,
      telIntl ? "https://wa.me/" + telIntl.replace("+", "") : "",
      lote,
    ];
  });

  expSheet.getRange(2, 1, filas.length, 5).setValues(filas);
  expSheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(
    `✅ Listo!\n\nTotal exportados: ${filas.length}\n\n` +
      `⚠️ Algunos pueden no tener teléfono separado — revisa la hoja CAMPAÑA.\n\n` +
      `Archivo → Descargar → CSV`,
  );
}

//limpiador de datos

function limpiarYExportarCampana() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPacientes = ss.getSheetByName(CONFIG.SHEET_PACIENTES);
  const datos = shPacientes
    .getRange(2, 1, shPacientes.getLastRow() - 1, 2)
    .getValues();

  // Palabras que indican que NO es un paciente
  const basura = [
    "no agendar",
    "no citar",
    "feriado",
    "feria",
    "agenda",
    "consulta",
    "almuerzo",
    "fotos",
    "citas de",
    "interview",
    "psicología",
    "spa",
    "suarez",
    "teresita",
    "domicilio",
    "taller",
    "zoom",
    "virtual",
    "fuera",
    "lista de espera",
    "mayra",
    "it set",
    "deel",
    "dr.sáenz",
    "para don",
    "para idaly",
    "tratamiento domicilio",
    "no trabajo",
    "agendar 7",
    "agenda de 7",
    "agenda para",
    "no angendar",
  ];

  const esBasura = (texto) => {
    const t = texto.toLowerCase().trim();
    return basura.some((b) => t.includes(b)) || t.length < 4;
  };

  // Regex para extraer teléfonos de Costa Rica del texto
  const extraerTelefono = (texto) => {
    const match = texto.match(/[\+506\s]*([2-9]\d{3})[\s\-]?(\d{4})/);
    return match ? match[1] + match[2] : null;
  };

  // Limpiar nombre — quitar teléfono, prefijos C-, notas extras
  const limpiarNombre = (texto) => {
    return texto
      .replace(/[\+506\s]*[2-9]\d{3}[\s\-]?\d{4}/g, "") // quitar teléfonos
      .replace(/^[Cc]-\s*/g, "") // quitar prefijo C-
      .replace(/\s*\/\s*\d+/g, "") // quitar /número
      .replace(/\d{8,}/g, "") // quitar números solos largos
      .replace(/\s+/g, " ") // normalizar espacios
      .replace(/[,.]+$/, "") // quitar puntuación al final
      .trim();
  };

  const resultado = [];
  const telefonosVistos = new Set();
  const nombresVistos = new Set();

  datos.forEach((r) => {
    const nombreRaw = String(r[0] || "").trim();
    let telefonoRaw = String(r[1] || "").trim();

    if (!nombreRaw || esBasura(nombreRaw)) return;

    // Si no hay teléfono en columna B, intentar extraerlo del nombre
    if (!telefonoRaw) {
      telefonoRaw = extraerTelefono(nombreRaw) || "";
    }

    const nombreLimpio = limpiarNombre(nombreRaw);

    // Validar que parezca un nombre real (al menos 2 palabras o 5 chars)
    if (nombreLimpio.length < 5) return;
    if (nombreLimpio.split(" ").length < 2 && !telefonoRaw) return;

    // Limpiar teléfono
    const telLimpio = telefonoRaw.replace(/[\s\-\(\)\+506]/g, "");
    const telIntl = telLimpio ? "+506" + telLimpio.slice(-8) : "";

    // Evitar duplicados por teléfono o nombre
    const nombreKey = nombreLimpio.toLowerCase();
    if (nombresVistos.has(nombreKey)) return;
    if (telIntl && telefonosVistos.has(telIntl)) return;

    nombresVistos.add(nombreKey);
    if (telIntl) telefonosVistos.add(telIntl);

    resultado.push({
      nombre: nombreLimpio,
      telefono: telefonoRaw,
      telIntl,
      tieneTel: !!telIntl,
    });
  });

  // Separar los que tienen teléfono de los que no
  const conTel = resultado.filter((r) => r.tienetel !== false && r.telIntl);
  const sinTel = resultado.filter((r) => !r.telIntl);

  // Crear/limpiar hoja CAMPAÑA
  let expSheet = ss.getSheetByName("CAMPAÑA");
  if (!expSheet) expSheet = ss.insertSheet("CAMPAÑA");
  expSheet.clearContents();

  expSheet
    .getRange(1, 1, 1, 6)
    .setValues([
      [
        "Nombre",
        "Teléfono",
        "Teléfono Internacional",
        "Link WhatsApp",
        "Lote",
        "Estado",
      ],
    ])
    .setBackground("#1a73e8")
    .setFontColor("#fff")
    .setFontWeight("bold");

  const filas = resultado.map((p, i) => [
    p.nombre,
    p.telefono,
    p.telIntl,
    p.telIntl ? "https://wa.me/" + p.telIntl.replace("+", "") : "",
    i < 256 ? "Lista 1" : "Lista 2",
    p.telIntl ? "✅ Listo" : "⚠️ Sin teléfono",
  ]);

  if (filas.length > 0) {
    expSheet.getRange(2, 1, filas.length, 6).setValues(filas);
  }

  // Colorear filas sin teléfono en amarillo para revisión manual
  filas.forEach((f, i) => {
    if (!f[2]) {
      expSheet.getRange(i + 2, 1, 1, 6).setBackground("#fff8e1");
    }
  });

  expSheet.setFrozenRows(1);
  expSheet.autoResizeColumns(1, 6);

  SpreadsheetApp.getUi().alert(
    `✅ Limpieza completada\n\n` +
      `✅ Con teléfono (listos para campaña): ${conTel.length}\n` +
      `⚠️  Sin teléfono (fila amarilla, revisar manual): ${sinTel.length}\n` +
      `🗑️  Registros basura eliminados\n\n` +
      `Ve a la hoja CAMPAÑA, revisa las filas amarillas,\n` +
      `agrega los teléfonos que faltan y descarga el CSV.`,
  );
}

//validadcion de no numeros en el nombre

function protegerHojaPacientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_PACIENTES);

  // Validación en columna B (Teléfono) — solo números
  const rangoTel = sheet.getRange("B2:B1000");
  const regla = SpreadsheetApp.newDataValidation()
    .requireTextMatchesPattern("^[0-9\\s\\-\\+]+$")
    .setHelpText("Solo ingrese el número de teléfono (ej: 8888-1234)")
    .setAllowInvalid(false)
    .build();
  rangoTel.setDataValidation(regla);

  SpreadsheetApp.getUi().alert(
    "✅ Validación aplicada. La columna Teléfono solo aceptará números.",
  );
}

// ============================================================
// WHATSAPP AI AGENT — STAGING
// ============================================================

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    twilioSid: props.getProperty("TWILIO_ACCOUNT_SID"),
    twilioToken: props.getProperty("TWILIO_AUTH_TOKEN"),
    twilioFrom: props.getProperty("TWILIO_WHATSAPP_FROM"),
    claudeKey: props.getProperty("CLAUDE_API_KEY"),
  };
}

// ── Webhook principal ────────────────────────────────────────
function doPost(e) {
  try {
    const params = e.parameter || {};
    const mensaje = (params.Body || "").trim();
    const de = params.From || ""; // formato: whatsapp:+50688881234

    Logger.log("📩 Mensaje de: " + de + ' → "' + mensaje + '"');

    if (!mensaje || !de) {
      return xmlResponse();
    }

    // Extraer número limpio para buscar citas
    const telefono = de.replace("whatsapp:", "").replace("+506", "");

    // Responder
    const respuesta = procesarMensaje(mensaje, telefono, de);
    enviarWhatsApp(de, respuesta);

    return xmlResponse();
  } catch (err) {
    Logger.log("❌ Error doPost: " + err.message);
    return xmlResponse();
  }
}

// ── Procesar mensaje con Claude ──────────────────────────────
function procesarMensaje(mensaje, telefono, deWhatsApp) {
  const hoy = Utilities.formatDate(
    new Date(),
    "America/Costa_Rica",
    "yyyy-MM-dd",
  );

  // Datos reales del Sheet
  const citasPaciente = getCitas({ buscar: telefono });
  const doctores = getDoctores();
  const servicios = getTodosLosServicios();
  const historial = getHistorial(telefono);

  // ── Intercepción de reserva pendiente ───────────────────────
  const sistemaReserva = historial
    .filter(function (h) {
      return h.role === "system";
    })
    .slice(-1)[0];

  if (sistemaReserva) {
    const ctx = sistemaReserva.content || "";

    // Extraer datos del contexto sin optional chaining
    const doctorMatch = ctx.match(/doctor=([^|]+)/);
    const fechaMatch = ctx.match(/fecha=([^|]+)/);
    const horaMatch2 = ctx.match(/hora=([^|]+)/);
    const slotsMatch = ctx.match(/slots_disponibles=(.+)/);

    const ctxDoctor = doctorMatch ? doctorMatch[1].trim() : "";
    const ctxFecha = fechaMatch ? fechaMatch[1].trim() : "";
    const ctxHora = horaMatch2 ? horaMatch2[1].trim() : "";
    const slots = slotsMatch ? slotsMatch[1].trim().split(",") : [];

    // Detectar si el usuario menciona una hora
    const horaEnMensaje =
      mensaje.match(/\b([0-9]{1,2})[:\.]([0-9]{2})\b/) ||
      mensaje.match(/\ba\s*las?\s*([0-9]{1,2})\b/i) ||
      mensaje.match(/\b([0-9]{1,2})\s*(am|pm)\b/i);

    // Si estamos esperando nombre — agendar directo
    if (ctx.includes("esperando=nombre") && ctxDoctor && ctxFecha && ctxHora) {
      const nombresInvalidos = [
        "si",
        "sí",
        "no",
        "ok",
        "okay",
        "bien",
        "hola",
        "gracias",
        "nada",
      ];
      const nombreLimpio = mensaje.trim();

      if (
        nombreLimpio.length > 2 &&
        !nombresInvalidos.includes(nombreLimpio.toLowerCase())
      ) {
        const servicios2 = getTodosLosServicios();
        const servicio2 = servicios2.find(function (s) {
          return s.doctor === ctxDoctor;
        });

        // Limpiar teléfono a 8 dígitos
        const telLimpio2 = String(telefono)
          .replace(/[^0-9]/g, "")
          .slice(-8);

        const resultado = agregarCita({
          doctor: ctxDoctor,
          fecha: ctxFecha,
          hora: ctxHora,
          paciente: nombreLimpio,
          telefono: telLimpio2,
          servicio: servicio2 ? servicio2.servicio : "",
          precio: servicio2 ? servicio2.precio : 0,
          notas: "",
        });

        const respuesta = resultado.success
          ? "Su cita quedó agendada con " +
            ctxDoctor +
            " el " +
            ctxFecha +
            " a las " +
            ctxHora +
            ". Le esperamos en " +
            CONFIG.NOMBRE_CLINICA +
            "."
          : "Hubo un problema al agendar. Por favor llámenos directamente.";

        guardarHistorial(telefono, "user", mensaje);
        guardarHistorial(telefono, "assistant", respuesta);
        return respuesta;
      }
    }

    // Si el usuario menciona una hora y hay slots disponibles
    if (horaEnMensaje && slots.length > 0 && ctxDoctor && ctxFecha) {
      var hh, mm;
      const m1 = mensaje.match(/\b([0-9]{1,2})[:\.]([0-9]{2})\b/);
      const m2 = mensaje.match(/\ba\s*las?\s*([0-9]{1,2})\b/i);
      const m3 = mensaje.match(/\b([0-9]{1,2})\s*(am|pm)\b/i);

      if (m1) {
        hh = m1[1];
        mm = m1[2];
      } else if (m2) {
        hh = m2[1];
        mm = "00";
      } else if (m3) {
        hh = m3[1];
        mm = "00";
      } else {
        hh = "0";
        mm = "00";
      }

      const horaNorm =
        String(hh).padStart(2, "0") + ":" + String(mm).padStart(2, "0");

      if (slots.indexOf(horaNorm) !== -1) {
        // Hora disponible — pedir nombre
        const respuesta =
          "Con gusto le agendo el " +
          ctxFecha +
          " con " +
          ctxDoctor +
          " a las " +
          horaNorm +
          ". ¿Me confirma su nombre completo?";
        guardarHistorial(telefono, "user", mensaje);
        guardarHistorial(telefono, "assistant", respuesta);
        guardarHistorial(
          telefono,
          "system",
          "CONTEXTO_RESERVA: doctor=" +
            ctxDoctor +
            " | fecha=" +
            ctxFecha +
            " | hora=" +
            horaNorm +
            " | esperando=nombre",
        );
        return respuesta;
      } else {
        // Hora no disponible — mostrar alternativas
        const respuesta =
          "Lo sentimos, las " +
          horaNorm +
          " no está disponible. Las horas libres son:\n\n" +
          slots
            .map(function (h) {
              return "⏰ " + h;
            })
            .join("\n") +
          "\n\n¿Cuál le queda mejor?";
        guardarHistorial(telefono, "user", mensaje);
        guardarHistorial(telefono, "assistant", respuesta);
        return respuesta;
      }
    }
  }

  // ── Fin intercepción ─────────────────────────────────────────

  // Contexto de citas del paciente
  const contextoCitas =
    citasPaciente.length > 0
      ? citasPaciente
          .slice(0, 5)
          .map(
            (c) =>
              `- Fila ${c.rowIndex} | ${c.fecha} ${c.hora} | ${c.doctor} | ${c.servicio} | Estado: ${c.estado}`,
          )
          .join("\n")
      : "No tiene citas registradas con este número.";

  // Contexto de servicios reales
  const contextoServicios =
    servicios.length > 0
      ? servicios
          .map(
            (s) =>
              `- ${s.doctor} | ${s.servicio} | ₡${Number(s.precio).toLocaleString("es-CR")} | ${s.duracion} min`,
          )
          .join("\n")
      : "No hay servicios registrados.";

  // Contexto de doctores reales
  const contextoDoctores = doctores
    .map((d) => `- ${d.nombre} (${d.especialidad})`)
    .join("\n");

  // Recuperar contexto de reserva pendiente si existe
  const ultimoSistema = historial
    .filter((h) => h.role === "system")
    .slice(-1)[0];
  const contextoReserva = ultimoSistema
    ? "\n\n── RESERVA EN PROCESO ──\n" +
      ultimoSistema.content +
      "\nSi el usuario indica una hora, agendá directamente con ese doctor y fecha sin volver a mostrar disponibilidad."
    : "";

  const systemPrompt =
    `Sos Maureen, recepcionista de ${CONFIG.NOMBRE_CLINICA}, clínica médica en Costa Rica. ` +
    `Atendés pacientes por WhatsApp de forma amable, cálida y natural, como lo haría una recepcionista humana. ` +
    `\n\n⚠️ REGLAS IMPORTANTES:` +
    `\n- NUNCA digas que sos un bot o asistente virtual. Sos Maureen, la recepcionista.` +
    `\n- NUNCA saludes con "Hola" si ya hay mensajes previos en la conversación — solo saludás al inicio.` +
    `\n- NUNCA inventes información. Solo usá los datos reales que te doy abajo.` +
    `\n- Si el paciente pregunta disponibilidad para una fecha, pedile el nombre del doctor y la fecha exacta para consultarla.` +
    `\n- Si el paciente quiere hablar con un humano, decile que en este momento estás atendiendo vos pero que si lo prefiere puede llamar al número de la clínica.` +
    `\n- Si el paciente ya indicó una hora y solo falta el nombre u otro dato, NO muestres disponibilidad de nuevo — simplemente agendá con la hora que ya dijo.` +
    `\n- Si el paciente menciona una hora junto con la fecha en el mismo mensaje (ej: "el martes a las 9:30"), NO consultés disponibilidad — agendá directamente con esa hora. Solo mostrás disponibilidad cuando el paciente NO indica ninguna hora.` +
    `\n- Para agendar necesitás: nombre completo, teléfono (8 dígitos), email (opcional), fecha y hora. Pedí estos datos uno por uno de forma natural.` +
    `\n- Para CANCELAR o REAGENDAR: primero mostrá las citas encontradas para ese número y preguntá cuál es la que desea modificar. NUNCA cancelés sin que el paciente confirme cuál cita es.` +
    `\n- Al mostrar citas para cancelar/reagendar, presentalas numeradas así: "1. Jueves 22 de mayo a las 9:30 con [nombre del doctor]" y pedile que indique el número.` +
    `\n- Usá emojis con moderación — máximo 1 por mensaje y solo cuando sea realmente necesario. La mayoría de mensajes no necesitan emoji.` +
    `\n- Usás español de Costa Rica con tratamiento formal de "usted", natural y cálido.` +
    `\n- Cuando te referís a la persona por su nombre, usá "Don" para hombres y "Doña" para mujeres. Si no sabés el sexo por el nombre, omitís el título.` +
    `\n- Ejemplos correctos: "Don Miguel, su cita quedó confirmada", "Doña Ana, con mucho gusto".` +
    `\n\nFecha de hoy: ${hoy}` +
    `\n\n── DOCTORES DISPONIBLES ──\n${contextoDoctores}` +
    `\n\n── SERVICIOS Y PRECIOS REALES ──\n${contextoServicios}` +
    `\n\n── CITAS DE ESTE PACIENTE ──\n${contextoCitas}` +
    `\n\n── ACCIONES DISPONIBLES ──` +
    `\nRespondé SIEMPRE con un JSON puro sin markdown ni backticks:` +
    `\n- Responder: {"accion":"mensaje","texto":"tu respuesta"}` +
    `\n- Confirmar cita: {"accion":"confirmar","rowIndex":NUM,"texto":"mensaje al paciente"}` +
    `\n- Cancelar cita: {"accion":"cancelar","rowIndex":NUM,"texto":"mensaje al paciente"}` +
    `\n- Consultar disponibilidad: {"accion":"disponibilidad","doctor":"nombre exacto","fecha":"yyyy-MM-dd","texto":"mensaje"}` +
    `\n- Agendar cita nueva: {"accion":"agendar","doctor":"nombre exacto","fecha":"yyyy-MM-dd","hora":"HH:mm","paciente":"nombre completo","telefono":"SOLO 8 digitos sin espacios ni guiones","email":"correo o vacio","servicio":"nombre exacto","precio":numero,"texto":"mensaje de confirmacion"}`;

  const config = getConfig();

  // Construir mensajes con historial
  const mensajesConHistorial = [
    ...historial.filter((h) => h.role === "user" || h.role === "assistant"),
    { role: "user", content: mensaje },
  ];

  const payload = {
    model: "claude-sonnet-4-5",
    max_tokens: 1024,
    system: systemPrompt,
    messages: mensajesConHistorial,
  };

  const options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": config.claudeKey,
      "anthropic-version": "2023-06-01",
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(
    "https://api.anthropic.com/v1/messages",
    options,
  );
  const data = JSON.parse(response.getContentText());

  Logger.log("🤖 Claude raw: " + response.getContentText());

  if (!data.content || !data.content[0]) {
    return "Disculpe, tuve un problema. Por favor llámenos directamente.";
  }

  const rawText = data.content[0].text.trim();
  const cleanText = rawText
    .replace(/^```json\s*/i, "")
    .replace(/^```\s*/i, "")
    .replace(/```\s*$/i, "")
    .trim();

  Logger.log("🧹 Claude clean: " + cleanText);

  try {
    const json = JSON.parse(cleanText);

    // Acción: agendar cita nueva
    if (json.accion === "agendar") {
      // Saludo si es primer mensaje
      const esInicio = historial.length === 0;
      const saludo = esInicio
        ? "Buen día, soy Maureen de Clínica Integral Mediveins. "
        : "";

      // Validar que el nombre del paciente sea real
      const nombresInvalidos = [
        "pendiente",
        "cliente",
        "paciente",
        "usuario",
        "jimena",
        "nombre",
        "desconocido",
        "",
        "null",
        "undefined",
      ];
      const nombreValido =
        json.paciente &&
        !nombresInvalidos.includes(
          String(json.paciente).toLowerCase().trim(),
        ) &&
        String(json.paciente).trim().length > 2;

      if (!nombreValido) {
        // No tenemos nombre real — pedirlo antes de agendar
        const horaNorm = json.hora || "";
        const respuesta =
          `${saludo}Con gusto le agendo para el ${json.fecha} con ${json.doctor}` +
          (horaNorm ? ` a las ${horaNorm}` : "") +
          `. ¿Me confirma su nombre completo para completar la reserva?`;
        guardarHistorial(telefono, "user", mensaje);
        guardarHistorial(telefono, "assistant", respuesta);
        return respuesta;
      }

      // Verificar disponibilidad antes de agendar
      if (json.fecha && json.hora && json.doctor) {
        const disp = getDisponibilidad(json.doctor, json.fecha);
        const horaNorm = String(json.hora).substring(0, 5);
        const horaOcupada = disp.ocupados.some(
          (h) => h.substring(0, 5) === horaNorm,
        );

        if (horaOcupada) {
          const slots = disp.disponibles.slice(0, 6);
          const respuesta =
            slots.length > 0
              ? `${saludo}Lo sentimos, la hora ${horaNorm} ya no está disponible con ${json.doctor} el ${json.fecha}. ` +
                `Estas son las horas disponibles ese día:\n\n` +
                slots.map((h) => `⏰ ${h}`).join("\n") +
                `\n\n¿Cuál le queda mejor?`
              : `${saludo}Lo sentimos, no hay espacios disponibles con ${json.doctor} el ${json.fecha}. ¿Le gustaría intentar otro día?`;
          guardarHistorial(telefono, "user", mensaje);
          guardarHistorial(telefono, "assistant", respuesta);
          return respuesta;
        }
      }

      // Limpiar teléfono — preferir el que dio el paciente, si no usar el del WhatsApp
      const telPaciente = json.telefono
        ? String(json.telefono)
            .replace(/[^0-9]/g, "")
            .slice(-8)
        : String(telefono)
            .replace(/[^0-9]/g, "")
            .slice(-8);

      const resultado = agregarCita({
        doctor: json.doctor,
        fecha: json.fecha,
        hora: json.hora,
        paciente: json.paciente,
        telefono: telPaciente,
        servicio: json.servicio,
        precio: json.precio,
        notas: json.email ? "Email: " + json.email : "",
      });

      const respuesta = resultado.success
        ? saludo +
          (json.texto ||
            `Su cita quedó agendada con ${json.doctor} el ${json.fecha} a las ${json.hora}.`)
        : `${saludo}Hubo un problema al agendar. Por favor llámenos directamente.`;

      guardarHistorial(telefono, "user", mensaje);
      guardarHistorial(telefono, "assistant", respuesta);
      return respuesta;
    }

    // Acción: disponibilidad
    if (json.accion === "disponibilidad" && json.doctor && json.fecha) {
      const disp = getDisponibilidad(json.doctor, json.fecha);

      // Verificar si el usuario ya indicó una hora específica
      const horaMatch =
        mensaje.match(/\b([0-9]{1,2})[:\s\.]([0-9]{2})\b/) ||
        mensaje.match(/\b([0-9]{1,2})\s*(am|pm|a\.m\.|p\.m\.)\b/i);

      if (horaMatch) {
        // Normalizar la hora del usuario a formato HH:mm para comparar
        const rawHora = horaMatch[0].trim();
        const partes = rawHora.split(/[:\s\.]/);
        const hh = String(partes[0]).padStart(2, "0");
        const mm = String(partes[1] || "00").padStart(2, "0");
        const horaNorm = hh + ":" + mm; // ej: "09:30"

        const horaOcupada = disp.ocupados.some(
          (h) => h.substring(0, 5) === horaNorm,
        );

        if (!horaOcupada) {
          // Hora libre — pedir nombre y agendar
          const esInicio = historial.length === 0;
          const saludo = esInicio
            ? "Buen día, soy Maureen de Clínica Integral Mediveins. "
            : "";
          const respuesta = `${saludo}Con gusto le agendo para el ${json.fecha} con ${json.doctor} a las ${horaNorm}. ¿Me confirma su nombre completo?`;
          guardarHistorial(telefono, "assistant", respuesta);
          return respuesta;
        } else {
          // Hora ocupada — avisar y mostrar alternativas
          const slots = disp.disponibles.slice(0, 6);
          const respuesta =
            slots.length > 0
              ? `Lo sentimos, para esa hora ya no tenemos espacio disponible con ${json.doctor}. ` +
                `Estas son las horas disponibles ese día:\n\n` +
                slots.map((h) => `⏰ ${h}`).join("\n") +
                `\n\n¿Cuál le queda mejor?`
              : `Lo sentimos, no tenemos espacios disponibles con ${json.doctor} el ${json.fecha}. ¿Le gustaría intentar otro día?`;
          guardarHistorial(telefono, "assistant", respuesta);
          return respuesta;
        }
      }

      // El usuario no indicó hora — mostrar disponibilidad normal
      const slots = disp.disponibles.slice(0, 8);
      let respuesta;
      if (slots.length === 0) {
        respuesta = `No hay espacio disponible con ${json.doctor} el ${json.fecha}. ¿Le gustaría intentar otro día?`;
      } else {
        respuesta =
          `Para el ${json.fecha} con ${json.doctor} tenemos estos espacios:\n\n` +
          slots.map((h) => `⏰ ${h}`).join("\n") +
          `\n\n¿Cuál le queda mejor?`;
      }
      // Guardar contexto de reserva pendiente para el siguiente mensaje
      guardarHistorial(telefono, "assistant", respuesta);
      guardarHistorial(
        telefono,
        "system",
        `CONTEXTO_RESERVA: doctor=${json.doctor} | fecha=${json.fecha} | slots_disponibles=${slots.join(",")}`,
      );
      return respuesta;
    }

    // Acción: confirmar cita
    if (json.accion === "confirmar" && json.rowIndex) {
      actualizarEstadoCita(json.rowIndex, "Confirmada");
      const respuesta =
        json.texto || "✅ Perfecto, su cita quedó confirmada. ¡Le esperamos!";
      guardarHistorial(telefono, "assistant", respuesta);
      return respuesta;
    }

    // Acción: cancelar cita
    if (json.accion === "cancelar" && json.rowIndex) {
      actualizarEstadoCita(json.rowIndex, "Cancelada");
      const respuesta =
        json.texto ||
        "Su cita fue cancelada. Si desea reagendarla con gusto la ayudo.";
      guardarHistorial(telefono, "assistant", respuesta);
      return respuesta;
    }

    // Acción: mensaje normal
    if (json.accion === "mensaje") {
      const respuesta = json.texto || "Gracias por contactarnos.";
      guardarHistorial(telefono, "assistant", respuesta);
      return respuesta;
    }

    return "Gracias por contactarnos.";
  } catch (parseErr) {
    Logger.log("⚠️ Parse error: " + cleanText);
    const fallback =
      "Disculpe, tuve un problema. ¿Me puede repetir lo que necesita?";
    guardarHistorial(telefono, "assistant", fallback);
    return fallback;
  }
}

// ── Enviar mensaje por Twilio ────────────────────────────────
function enviarWhatsApp(para, mensaje) {
  const config = getConfig();
  const url =
    "https://api.twilio.com/2010-04-01/Accounts/" +
    config.twilioSid +
    "/Messages.json";

  const options = {
    method: "post",
    headers: {
      Authorization:
        "Basic " +
        Utilities.base64Encode(config.twilioSid + ":" + config.twilioToken),
    },
    payload: {
      To: para,
      From: config.twilioFrom,
      Body: mensaje,
    },
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  Logger.log("📤 Twilio send: " + response.getContentText());
}

// ── Helper TwiML vacío ───────────────────────────────────────
function xmlResponse() {
  return ContentService.createTextOutput("<Response></Response>").setMimeType(
    ContentService.MimeType.XML,
  );
}

function testAgente() {
  // Simula exactamente lo que haría doPost con un mensaje real
  const mensajePrueba = "Hola, quiero saber mis citas";
  const telefonoPrueba = "71702184"; // tu número sin +506
  const dePrueba = "whatsapp:+50671702184"; // tu número completo

  Logger.log("=== TEST INICIO ===");

  // Paso 1: verificar que las keys existen
  const config = getConfig();
  Logger.log("Twilio SID: " + (config.twilioSid ? "✅ OK" : "❌ VACÍO"));
  Logger.log("Twilio Token: " + (config.twilioToken ? "✅ OK" : "❌ VACÍO"));
  Logger.log("Twilio From: " + (config.twilioFrom ? "✅ OK" : "❌ VACÍO"));
  Logger.log("Claude Key: " + (config.claudeKey ? "✅ OK" : "❌ VACÍO"));

  // Paso 2: verificar que getCitas funciona
  const citas = getCitas({ buscar: telefonoPrueba });
  Logger.log("Citas encontradas: " + citas.length);

  // Paso 3: llamar Claude
  Logger.log("Llamando Claude...");
  const respuesta = procesarMensaje(mensajePrueba, telefonoPrueba, dePrueba);
  Logger.log("Respuesta Claude: " + respuesta);

  // Paso 4: enviar WhatsApp
  Logger.log("Enviando WhatsApp...");
  enviarWhatsApp(dePrueba, "🧪 TEST: " + respuesta);

  Logger.log("=== TEST FIN ===");
}

// ── Historial de conversación por número ─────────────────────
function getHistorial(telefono) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("CONVERSACIONES");
  if (!sheet) {
    sheet = ss.insertSheet("CONVERSACIONES");
    sheet
      .getRange(1, 1, 1, 4)
      .setValues([["Teléfono", "Timestamp", "Rol", "Mensaje"]]);
    sheet
      .getRange(1, 1, 1, 4)
      .setBackground("#1a73e8")
      .setFontColor("#fff")
      .setFontWeight("bold");
  }

  if (sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();

  // Últimos 10 mensajes de este número
  return data
    .filter((r) => String(r[0]) === String(telefono))
    .slice(-10)
    .map((r) => ({ role: r[2], content: r[3] }));
}

function guardarHistorial(telefono, rol, mensaje) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("CONVERSACIONES");
  if (!sheet) {
    sheet = ss.insertSheet("CONVERSACIONES");
    sheet
      .getRange(1, 1, 1, 4)
      .setValues([["Teléfono", "Timestamp", "Rol", "Mensaje"]]);
  }
  sheet.appendRow([
    telefono,
    Utilities.formatDate(
      new Date(),
      "America/Costa_Rica",
      "yyyy-MM-dd HH:mm:ss",
    ),
    rol,
    mensaje,
  ]);
}

// ── Disponibilidad real por doctor y fecha ───────────────────
function getDisponibilidad(nombreDoctor, fecha) {
  // Horarios base de la clínica (7am a 5pm cada 30 min)
  const todosLosHorarios = [];
  for (let h = 7; h <= 16; h++) {
    todosLosHorarios.push(String(h).padStart(2, "0") + ":00");
    todosLosHorarios.push(String(h).padStart(2, "0") + ":30");
  }
  todosLosHorarios.push("17:00");

  // Citas ya ocupadas para ese doctor y fecha
  const citasOcupadas = getCitas({ fecha: fecha, doctor: nombreDoctor })
    .filter((c) => c.estado !== "Cancelada")
    .map((c) => c.hora.substring(0, 5));

  // Retornar solo los espacios libres
  const disponibles = todosLosHorarios.filter(
    (h) => !citasOcupadas.includes(h),
  );

  return {
    doctor: nombreDoctor,
    fecha: fecha,
    disponibles: disponibles,
    ocupados: citasOcupadas,
  };
}

// ── Servicios reales del Sheet ───────────────────────────────
function getTodosLosServicios() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.SHEET_SERVICIOS,
  );
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 4)
    .getValues()
    .filter((r) => r[0])
    .map((r) => ({
      doctor: r[0],
      servicio: r[1],
      precio: r[2],
      duracion: r[3],
    }));
}

// ============================================================
// RECORDATORIOS AUTOMÁTICOS WHATSAPP — 48 horas antes
// ============================================================

function enviarRecordatoriosWhatsApp() {
  const tz = "America/Costa_Rica";

  // Calcular fecha de pasado mañana (48 horas adelante)
  const pasadoManana = new Date();
  pasadoManana.setDate(pasadoManana.getDate() + 2);
  const fechaObjetivo = Utilities.formatDate(pasadoManana, tz, "yyyy-MM-dd");

  Logger.log("📅 Buscando citas para: " + fechaObjetivo);

  // Obtener citas de ese día que no estén canceladas
  const basura = [
    "almuerzo",
    "feriado",
    "agenda",
    "citas de",
    "domicilio",
    "no agendar",
    "zoom",
    "virtual",
    "fuera",
    "taller",
    "it set",
  ];

  const citas = getCitas({ fecha: fechaObjetivo })
    .filter((c) => c.estado !== "Cancelada" && c.estado !== "Completada")
    .filter((c) => {
      const nombre = String(c.paciente || "").toLowerCase();
      return !basura.some((b) => nombre.includes(b)) && nombre.length > 3;
    });

  Logger.log("📋 Citas encontradas: " + citas.length);

  if (citas.length === 0) {
    Logger.log("✅ No hay citas para recordar ese día");
    return;
  }

  const config = getConfig();

  citas.forEach((cita) => {
    try {
      // Formatear fecha legible
      const [y, m, d] = cita.fecha.split("-").map(Number);
      const fechaObj = new Date(y, m - 1, d);
      const diasSemana = [
        "domingo",
        "lunes",
        "martes",
        "miércoles",
        "jueves",
        "viernes",
        "sábado",
      ];
      const meses = [
        "enero",
        "febrero",
        "marzo",
        "abril",
        "mayo",
        "junio",
        "julio",
        "agosto",
        "septiembre",
        "octubre",
        "noviembre",
        "diciembre",
      ];
      const fechaLegible = `${diasSemana[fechaObj.getDay()]} ${d} de ${meses[m - 1]}`;
      const hora12 = formatearHora12(cita.hora);
      const nombre = String(cita.paciente || "").split(" ")[0];

      const mensaje =
        `Hola ${nombre}, te recordamos tu cita en *${CONFIG.NOMBRE_CLINICA}*:\n\n` +
        `📅 ${fechaLegible} a las ${hora12}\n` +
        `👨‍⚕️ ${cita.doctor}\n` +
        `💉 ${cita.servicio}\n\n` +
        `Por favor confirmanos tu asistencia respondiendo *SÍ* o *NO* a este mensaje.\n\n` +
        `📍 ¿Cómo llegar? → ${CONFIG.LINK_UBICACION}`;

      // Validar teléfono antes de enviar
      const telLimpio = String(cita.telefono || "").replace(/[\s\-\(\)]/g, "");

      // Debe tener exactamente 8 dígitos para Costa Rica
      if (!/^\d{8}$/.test(telLimpio)) {
        Logger.log(
          "⚠️ Teléfono inválido, saltando: " +
            cita.paciente +
            ' · "' +
            cita.telefono +
            '"',
        );
        return; // ← salta esta cita sin enviar
      }
      const telIntl = "+506" + telLimpio;
      const para = "whatsapp:" + telIntl;

      const response = UrlFetchApp.fetch(
        "https://api.twilio.com/2010-04-01/Accounts/" +
          config.twilioSid +
          "/Messages.json",
        {
          method: "post",
          headers: {
            Authorization:
              "Basic " +
              Utilities.base64Encode(
                config.twilioSid + ":" + config.twilioToken,
              ),
          },
          payload: { To: para, From: config.twilioFrom, Body: mensaje },
          muteHttpExceptions: true,
        },
      );

      const twilioData = JSON.parse(response.getContentText());
      if (twilioData.error_code || twilioData.code) {
        Logger.log(
          "❌ Error Twilio para " + cita.paciente + ": " + twilioData.message,
        );
      } else {
        Logger.log(
          "✅ Recordatorio enviado a: " + cita.paciente + " · " + para,
        );
      }
    } catch (err) {
      Logger.log("❌ Error enviando a " + cita.paciente + ": " + err.message);
    }
  });

  Logger.log("🏁 Recordatorios completados: " + citas.length + " procesados");
}

// ── Trigger diario 8am para recordatorios WhatsApp ───────────
function configurarTriggerRecordatoriosWhatsApp() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === "enviarRecordatoriosWhatsApp")
    .forEach((t) => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("enviarRecordatoriosWhatsApp")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .inTimezone("America/Costa_Rica")
    .create();

  Logger.log(
    "✅ Trigger configurado: recordatorios WhatsApp cada día a las 8am",
  );
}

// ── Reset de conversación para testing ───────────────────────
function resetearConversacion() {
  const telefono = "tu numero aqui";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CONVERSACIONES");
  if (!sheet || sheet.getLastRow() <= 1) return;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]) === String(telefono)) {
      sheet.deleteRow(i + 2);
    }
  }
  Logger.log("✅ Historial borrado para: " + telefono);
}
