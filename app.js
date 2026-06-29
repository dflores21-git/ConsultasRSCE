require('dotenv').config();
const express = require('express');
const sql = require('mssql');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const path = require('path');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── Conexión SQL Server ─────────────────────────────────────────────────────
// Las credenciales ahora se leen del archivo .env (ver .env.example)
const dbConfig = {
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  server: process.env.DB_SERVER,
  database: process.env.DB_DATABASE,
  options: { encrypt: false, trustServerCertificate: true },
  pool: { max: 10, min: 0, idleTimeoutMillis: 30000 }
};

let pool;
async function getPool() {
  if (!pool) pool = await sql.connect(dbConfig);
  return pool;
}

// ── Helper: añade parámetro solo si tiene valor ─────────────────────────────
function addParam(request, name, type, value) {
  if (value !== null && value !== undefined && value !== '') {
    request.input(name, type, value);
    return true;
  }
  return false;
}

// ── Definición de todas las consultas ──────────────────────────────────────
const QUERIES = {

  socios: {
    title: 'Consulta de Socios',
    columns: ['Nº Socio','Apellidos','Nombre','DNI','Teléfono Móvil','Email','Tipo','F. Alta','F. Baja','Estado','Deudas','Entidad'],
    build: (p) => {
      let w = [];
      if (p.entidad)      w.push('S.SOCIOS_ENTIDAD = @entidad');
      if (p.soloActivos)  w.push('S.SOCIOS_FECHABAJA IS NULL');
      if (p.nombre)       w.push("P.PERSONA_APELLIDOS + ' ' + P.PERSONA_NOMBRE LIKE '%' + @nombre + '%'");
      const where = w.length ? 'AND ' + w.join(' AND ') : '';
      return {
        sql: `SELECT S.SOCIOS_SOCIO AS NumSocio, P.PERSONA_APELLIDOS AS Apellidos,
              P.PERSONA_NOMBRE AS Nombre, P.PERSONA_ID AS DNI,
              P.PERSONA_TELEFONOM AS Telefono, P.PERSONA_EMAILT AS Email,
              S.SOCIOS_AB AS Tipo, S.SOCIOS_FECHAALTA AS FechaAlta,
              S.SOCIOS_FECHABAJA AS FechaBaja,
              CASE WHEN S.SOCIOS_FECHABAJA IS NULL THEN 'Activo' ELSE 'Baja' END AS Estado,
              S.SOCIOS_DEUDAS AS Deudas, E.ENTIDAD_DENO AS Entidad
              FROM SOCIOS S
              INNER JOIN PERSONAS P ON P.PERSONA_CODIGO = S.SOCIOS_PERSONA
              INNER JOIN ENTIDADES E ON E.ENTIDAD_CODENT = S.SOCIOS_ENTIDAD
              WHERE 1=1 ${where}
              ORDER BY P.PERSONA_APELLIDOS, P.PERSONA_NOMBRE`,
        params: (req) => {
          if (p.entidad) req.input('entidad', sql.Int, parseInt(p.entidad));
          if (p.nombre)  req.input('nombre',  sql.VarChar, p.nombre);
        }
      };
    }
  },

  perros_propietario: {
    title: 'Perros por Propietario',
    columns: ['LOE','Nombre Perro','Raza','Sexo','F. Nacimiento','F. Inscripción','Apellidos Propietario','Nombre Propietario','Email','Prop. Desde','Prop. Hasta','Afijo','Entidad'],
    build: (p) => {
      let w = [];
      if (p.personaId)   w.push('PP.PPRO_IDPERSONA = @personaId');
      if (p.nombre)      w.push("PER.PERSONA_APELLIDOS + ' ' + PER.PERSONA_NOMBRE LIKE '%' + @nombre + '%'");
      if (p.soloActuales) w.push('PP.PPRO_PROPHASTA IS NULL');
      if (p.raza)        w.push('P.PERROS_CODRAZA = @raza');
      const where = w.length ? 'AND ' + w.join(' AND ') : '';
      return {
        sql: `SELECT P.PERROS_CODPERRO AS LOE, P.PERROS_NOMBRE AS NombrePerro,
              R.RAZA_DENORAZAESP AS Raza, P.PERROS_SEXO AS Sexo,
              P.PERROS_FECNAC AS FechaNacimiento, P.PERROS_FECINSCRIP AS FechaInscripcion,
              PER.PERSONA_APELLIDOS AS ApellidosPropietario, PER.PERSONA_NOMBRE AS NombrePropietario,
              PER.PERSONA_EMAILT AS Email,
              PP.PPRO_PROPDESDE AS PropDesde, PP.PPRO_PROPHASTA AS PropHasta,
              A.AFIJO_NOMBRE AS Afijo, E.ENTIDAD_DENO AS Entidad
              FROM PERROS P
              INNER JOIN PERROS_PROPIETARIOS PP ON PP.PPRO_CODPERRO = P.PERROS_CODPERRO
              INNER JOIN PERSONAS PER ON PER.PERSONA_CODIGO = PP.PPRO_IDPERSONA
              LEFT JOIN RAZAS R ON R.RAZA_CODRAZA = P.PERROS_CODRAZA
              LEFT JOIN AFIJOS A ON A.AFIJO_CODIGO = P.PERROS_PERRO_AFIJO
              LEFT JOIN ENTIDADES E ON E.ENTIDAD_CODENT = P.PERROS_CODENT
              WHERE 1=1 ${where}
              ORDER BY PER.PERSONA_APELLIDOS, PER.PERSONA_NOMBRE, P.PERROS_NOMBRE`,
        params: (req) => {
          if (p.personaId) req.input('personaId', sql.Int, parseInt(p.personaId));
          if (p.nombre)    req.input('nombre', sql.VarChar, p.nombre);
          if (p.raza)      req.input('raza', sql.Int, parseInt(p.raza));
        }
      };
    }
  },

  camadas_criador: {
    title: 'Camadas Inscritas por Criador',
    columns: ['Código','Apellidos Criador','Nombre Criador','Afijo','Raza','F. Nacimiento','F. Notificación','F. Inscripción','H. Notif.','H. Inscr.','M. Notif.','M. Inscr.','Total Notif.','Total Inscr.','Padre','Madre','Entidad'],
    build: (p) => {
      let w = ['C.CAMADA_FEC_INSCRIP IS NOT NULL'];
      if (p.criadorId)   w.push('PC.PCRI_IDPERSONA = @criadorId');
      if (p.entidad)     w.push('C.CAMADA_CODENT = @entidad');
      if (p.fechaDesde)  w.push('C.CAMADA_FEC_INSCRIP >= @fechaDesde');
      if (p.fechaHasta)  w.push('C.CAMADA_FEC_INSCRIP <= @fechaHasta');
      return {
        sql: `SELECT C.CAMADA_CODIGO AS Codigo,
              PER.PERSONA_APELLIDOS AS ApellidosCriador, PER.PERSONA_NOMBRE AS NombreCriador,
              A.AFIJO_NOMBRE AS Afijo, R.RAZA_DENORAZAESP AS Raza,
              C.CAMADA_FEC_NAC AS FechaNacimiento, C.CAMADA_FEC_NOTIFICACION AS FechaNotificacion,
              C.CAMADA_FEC_INSCRIP AS FechaInscripcion,
              C.CAMADA_H_NOTIFICADAS AS HNotif, C.CAMADA_H_INSCRITAS AS HInscr,
              C.CAMADA_M_NOTIFICADOS AS MNotif, C.CAMADA_M_INSCRITOS AS MInscr,
              (C.CAMADA_H_NOTIFICADAS + C.CAMADA_M_NOTIFICADOS) AS TotalNotif,
              (C.CAMADA_H_INSCRITAS + C.CAMADA_M_INSCRITOS) AS TotalInscr,
              PP.PERROS_NOMBRE AS Padre, PM.PERROS_NOMBRE AS Madre,
              E.ENTIDAD_DENO AS Entidad
              FROM CAMADAS C
              INNER JOIN PERROS_CRIADORES PC ON PC.PCRI_CODPERRO = C.CAMADA_PADRE_CODPERRO
              INNER JOIN PERSONAS PER ON PER.PERSONA_CODIGO = PC.PCRI_IDPERSONA
              LEFT JOIN AFIJOS A ON A.AFIJO_CODIGO = C.CAMADA_AFIJO
              LEFT JOIN PERROS PP ON PP.PERROS_CODPERRO = C.CAMADA_PADRE_CODPERRO
              LEFT JOIN PERROS PM ON PM.PERROS_CODPERRO = C.CAMADA_MADRE_CODPERRO
              LEFT JOIN RAZAS R ON R.RAZA_CODRAZA = PP.PERROS_CODRAZA
              LEFT JOIN ENTIDADES E ON E.ENTIDAD_CODENT = C.CAMADA_CODENT
              WHERE ${w.join(' AND ')}
              ORDER BY C.CAMADA_FEC_INSCRIP DESC`,
        params: (req) => {
          if (p.criadorId)  req.input('criadorId', sql.Int, parseInt(p.criadorId));
          if (p.entidad)    req.input('entidad', sql.Int, parseInt(p.entidad));
          if (p.fechaDesde) req.input('fechaDesde', sql.Date, p.fechaDesde);
          if (p.fechaHasta) req.input('fechaHasta', sql.Date, p.fechaHasta);
        }
      };
    }
  },

  perros_progenitores: {
    title: 'Perros por Progenitores',
    columns: ['LOE','Nombre','Raza','Sexo','F. Nacimiento','F. Inscripción','LOE Padre','Nombre Padre','LOE Madre','Nombre Madre','Afijo'],
    build: (p) => {
      let w = [];
      if (p.padre) w.push('P.PERROS_CODPERROPADRE = @padre');
      if (p.madre) w.push('P.PERROS_CODPERROMADRE = @madre');
      if (p.raza)  w.push('P.PERROS_CODRAZA = @raza');
      const where = w.length ? 'WHERE ' + w.join(' AND ') : 'WHERE 1=0';
      return {
        sql: `SELECT P.PERROS_CODPERRO AS LOE, P.PERROS_NOMBRE AS Nombre,
              R.RAZA_DENORAZAESP AS Raza, P.PERROS_SEXO AS Sexo,
              P.PERROS_FECNAC AS FechaNacimiento, P.PERROS_FECINSCRIP AS FechaInscripcion,
              PAD.PERROS_CODPERRO AS LOEPadre, PAD.PERROS_NOMBRE AS NombrePadre,
              MAD.PERROS_CODPERRO AS LOEMadre, MAD.PERROS_NOMBRE AS NombreMadre,
              A.AFIJO_NOMBRE AS Afijo
              FROM PERROS P
              LEFT JOIN PERROS PAD ON PAD.PERROS_CODPERRO = P.PERROS_CODPERROPADRE
              LEFT JOIN PERROS MAD ON MAD.PERROS_CODPERRO = P.PERROS_CODPERROMADRE
              LEFT JOIN RAZAS R ON R.RAZA_CODRAZA = P.PERROS_CODRAZA
              LEFT JOIN AFIJOS A ON A.AFIJO_CODIGO = P.PERROS_PERRO_AFIJO
              ${where}
              ORDER BY P.PERROS_FECNAC DESC`,
        params: (req) => {
          if (p.padre) req.input('padre', sql.VarChar, p.padre);
          if (p.madre) req.input('madre', sql.VarChar, p.madre);
          if (p.raza)  req.input('raza', sql.Int, parseInt(p.raza));
        }
      };
    }
  },

  perros_criador: {
    title: 'Perros por Criador',
    columns: ['LOE','Nombre Perro','Raza','Sexo','F. Nacimiento','F. Inscripción','Apellidos Criador','Nombre Criador','Afijo','Apellidos Propietario','Nombre Propietario'],
    build: (p) => {
      let w = [];
      if (p.criadorId)   w.push('PC.PCRI_IDPERSONA = @criadorId');
      if (p.raza)        w.push('P.PERROS_CODRAZA = @raza');
      if (p.fechaDesde)  w.push('P.PERROS_FECINSCRIP >= @fechaDesde');
      if (p.fechaHasta)  w.push('P.PERROS_FECINSCRIP <= @fechaHasta');
      const where = w.length ? 'AND ' + w.join(' AND ') : '';
      return {
        sql: `SELECT P.PERROS_CODPERRO AS LOE, P.PERROS_NOMBRE AS NombrePerro,
              R.RAZA_DENORAZAESP AS Raza, P.PERROS_SEXO AS Sexo,
              P.PERROS_FECNAC AS FechaNacimiento, P.PERROS_FECINSCRIP AS FechaInscripcion,
              CRI.PERSONA_APELLIDOS AS ApellidosCriador, CRI.PERSONA_NOMBRE AS NombreCriador,
              A.AFIJO_NOMBRE AS Afijo,
              PROP.PERSONA_APELLIDOS AS ApellidosPropietario, PROP.PERSONA_NOMBRE AS NombrePropietario
              FROM PERROS P
              INNER JOIN PERROS_CRIADORES PC ON PC.PCRI_CODPERRO = P.PERROS_CODPERRO
              INNER JOIN PERSONAS CRI ON CRI.PERSONA_CODIGO = PC.PCRI_IDPERSONA
              LEFT JOIN RAZAS R ON R.RAZA_CODRAZA = P.PERROS_CODRAZA
              LEFT JOIN AFIJOS A ON A.AFIJO_CODIGO = P.PERROS_PERRO_AFIJO
              LEFT JOIN PERROS_PROPIETARIOS PP ON PP.PPRO_CODPERRO = P.PERROS_CODPERRO AND PP.PPRO_PROPHASTA IS NULL
              LEFT JOIN PERSONAS PROP ON PROP.PERSONA_CODIGO = PP.PPRO_IDPERSONA
              WHERE 1=1 ${where}
              ORDER BY CRI.PERSONA_APELLIDOS, P.PERROS_FECINSCRIP DESC`,
        params: (req) => {
          if (p.criadorId)  req.input('criadorId', sql.Int, parseInt(p.criadorId));
          if (p.raza)       req.input('raza', sql.Int, parseInt(p.raza));
          if (p.fechaDesde) req.input('fechaDesde', sql.Date, p.fechaDesde);
          if (p.fechaHasta) req.input('fechaHasta', sql.Date, p.fechaHasta);
        }
      };
    }
  },

  criadores_raza: {
    title: 'Criadores de una Raza',
    columns: ['Código','Apellidos','Nombre','DNI','Email','Teléfono','Raza','Afijo','Perros Criados'],
    build: (p) => {
      let w = [];
      if (p.raza)  w.push('P.PERROS_CODRAZA = @raza');
      if (p.afijo) w.push('P.PERROS_PERRO_AFIJO = @afijo');
      const where = w.length ? 'WHERE ' + w.join(' AND ') : 'WHERE 1=1';
      return {
        sql: `SELECT CRI.PERSONA_CODIGO AS Codigo,
              CRI.PERSONA_APELLIDOS AS Apellidos, CRI.PERSONA_NOMBRE AS Nombre,
              CRI.PERSONA_ID AS DNI, CRI.PERSONA_EMAILT AS Email,
              CRI.PERSONA_TELEFONOM AS Telefono,
              R.RAZA_DENORAZAESP AS Raza, A.AFIJO_NOMBRE AS Afijo,
              COUNT(P.PERROS_CODPERRO) AS NumPerrosCriados
              FROM PERROS_CRIADORES PC
              INNER JOIN PERSONAS CRI ON CRI.PERSONA_CODIGO = PC.PCRI_IDPERSONA
              INNER JOIN PERROS P ON P.PERROS_CODPERRO = PC.PCRI_CODPERRO
              INNER JOIN RAZAS R ON R.RAZA_CODRAZA = P.PERROS_CODRAZA
              LEFT JOIN AFIJOS A ON A.AFIJO_CODIGO = P.PERROS_PERRO_AFIJO
              ${where}
              GROUP BY CRI.PERSONA_CODIGO, CRI.PERSONA_APELLIDOS, CRI.PERSONA_NOMBRE,
              CRI.PERSONA_ID, CRI.PERSONA_EMAILT, CRI.PERSONA_TELEFONOM,
              R.RAZA_DENORAZAESP, A.AFIJO_NOMBRE
              ORDER BY R.RAZA_DENORAZAESP, CRI.PERSONA_APELLIDOS`,
        params: (req) => {
          if (p.raza)  req.input('raza', sql.Int, parseInt(p.raza));
          if (p.afijo) req.input('afijo', sql.Int, parseInt(p.afijo));
        }
      };
    }
  },

  propietarios_raza: {
    title: 'Propietarios de una Raza',
    columns: ['Código','Apellidos','Nombre','DNI','Email','Teléfono','Socio','Raza','Nº Perros'],
    build: (p) => {
      let w = ['PP.PPRO_PROPHASTA IS NULL'];
      if (p.raza)       w.push('P.PERROS_CODRAZA = @raza');
      if (p.soloSocios) w.push('PER.PERSONA_SOCIO = 1');
      if (p.entidad)    w.push('P.PERROS_CODENT = @entidad');
      return {
        sql: `SELECT PER.PERSONA_CODIGO AS Codigo,
              PER.PERSONA_APELLIDOS AS Apellidos, PER.PERSONA_NOMBRE AS Nombre,
              PER.PERSONA_ID AS DNI, PER.PERSONA_EMAILT AS Email,
              PER.PERSONA_TELEFONOM AS Telefono,
              CASE WHEN PER.PERSONA_SOCIO = 1 THEN 'Sí' ELSE 'No' END AS Socio,
              R.RAZA_DENORAZAESP AS Raza, COUNT(P.PERROS_CODPERRO) AS NumPerros
              FROM PERROS_PROPIETARIOS PP
              INNER JOIN PERSONAS PER ON PER.PERSONA_CODIGO = PP.PPRO_IDPERSONA
              INNER JOIN PERROS P ON P.PERROS_CODPERRO = PP.PPRO_CODPERRO
              INNER JOIN RAZAS R ON R.RAZA_CODRAZA = P.PERROS_CODRAZA
              WHERE ${w.join(' AND ')}
              GROUP BY PER.PERSONA_CODIGO, PER.PERSONA_APELLIDOS, PER.PERSONA_NOMBRE,
              PER.PERSONA_ID, PER.PERSONA_EMAILT, PER.PERSONA_TELEFONOM,
              PER.PERSONA_SOCIO, R.RAZA_DENORAZAESP
              ORDER BY R.RAZA_DENORAZAESP, PER.PERSONA_APELLIDOS`,
        params: (req) => {
          if (p.raza)    req.input('raza', sql.Int, parseInt(p.raza));
          if (p.entidad) req.input('entidad', sql.Int, parseInt(p.entidad));
        }
      };
    }
  },

  perros_fechas: {
    title: 'Perros Inscritos entre Fechas',
    columns: ['LOE','Nombre','Raza','Sexo','F. Nacimiento','F. Inscripción','Afijo','Apellidos Propietario','Nombre Propietario','Apellidos Criador','Nombre Criador','Entidad'],
    build: (p) => {
      let w = ['P.PERROS_FECINSCRIP IS NOT NULL'];
      if (p.fechaDesde) w.push('P.PERROS_FECINSCRIP >= @fechaDesde');
      if (p.fechaHasta) w.push('P.PERROS_FECINSCRIP <= @fechaHasta');
      if (p.raza)       w.push('P.PERROS_CODRAZA = @raza');
      if (p.entidad)    w.push('P.PERROS_CODENT = @entidad');
      return {
        sql: `SELECT P.PERROS_CODPERRO AS LOE, P.PERROS_NOMBRE AS Nombre,
              R.RAZA_DENORAZAESP AS Raza, P.PERROS_SEXO AS Sexo,
              P.PERROS_FECNAC AS FechaNacimiento, P.PERROS_FECINSCRIP AS FechaInscripcion,
              A.AFIJO_NOMBRE AS Afijo,
              PER.PERSONA_APELLIDOS AS ApellidosPropietario, PER.PERSONA_NOMBRE AS NombrePropietario,
              CRI.PERSONA_APELLIDOS AS ApellidosCriador, CRI.PERSONA_NOMBRE AS NombreCriador,
              E.ENTIDAD_DENO AS Entidad
              FROM PERROS P
              LEFT JOIN RAZAS R ON R.RAZA_CODRAZA = P.PERROS_CODRAZA
              LEFT JOIN AFIJOS A ON A.AFIJO_CODIGO = P.PERROS_PERRO_AFIJO
              LEFT JOIN ENTIDADES E ON E.ENTIDAD_CODENT = P.PERROS_CODENT
              LEFT JOIN PERROS_PROPIETARIOS PP ON PP.PPRO_CODPERRO = P.PERROS_CODPERRO AND PP.PPRO_PROPHASTA IS NULL
              LEFT JOIN PERSONAS PER ON PER.PERSONA_CODIGO = PP.PPRO_IDPERSONA
              LEFT JOIN PERROS_CRIADORES PC ON PC.PCRI_CODPERRO = P.PERROS_CODPERRO
              LEFT JOIN PERSONAS CRI ON CRI.PERSONA_CODIGO = PC.PCRI_IDPERSONA
              WHERE ${w.join(' AND ')}
              ORDER BY P.PERROS_FECINSCRIP DESC`,
        params: (req) => {
          if (p.fechaDesde) req.input('fechaDesde', sql.Date, p.fechaDesde);
          if (p.fechaHasta) req.input('fechaHasta', sql.Date, p.fechaHasta);
          if (p.raza)       req.input('raza', sql.Int, parseInt(p.raza));
          if (p.entidad)    req.input('entidad', sql.Int, parseInt(p.entidad));
        }
      };
    }
  },

  camadas_inscritas: {
    title: 'Camadas Inscritas',
    columns: ['Código','F. Nacimiento','F. Notificación','F. Inscripción','H. Notif.','H. Inscr.','M. Notif.','M. Inscr.','Total','Afijo','Raza','Padre','Madre','Entidad'],
    build: (p) => {
      let w = ['C.CAMADA_FEC_INSCRIP IS NOT NULL'];
      if (p.entidad)    w.push('C.CAMADA_CODENT = @entidad');
      if (p.fechaDesde) w.push('C.CAMADA_FEC_INSCRIP >= @fechaDesde');
      if (p.fechaHasta) w.push('C.CAMADA_FEC_INSCRIP <= @fechaHasta');
      return {
        sql: `SELECT C.CAMADA_CODIGO AS Codigo,
              C.CAMADA_FEC_NAC AS FechaNacimiento, C.CAMADA_FEC_NOTIFICACION AS FechaNotificacion,
              C.CAMADA_FEC_INSCRIP AS FechaInscripcion,
              C.CAMADA_H_NOTIFICADAS AS HNotif, C.CAMADA_H_INSCRITAS AS HInscr,
              C.CAMADA_M_NOTIFICADOS AS MNotif, C.CAMADA_M_INSCRITOS AS MInscr,
              (C.CAMADA_H_INSCRITAS + C.CAMADA_M_INSCRITOS) AS Total,
              A.AFIJO_NOMBRE AS Afijo, R.RAZA_DENORAZAESP AS Raza,
              PP.PERROS_NOMBRE AS Padre, PM.PERROS_NOMBRE AS Madre,
              E.ENTIDAD_DENO AS Entidad
              FROM CAMADAS C
              LEFT JOIN AFIJOS A ON A.AFIJO_CODIGO = C.CAMADA_AFIJO
              LEFT JOIN PERROS PP ON PP.PERROS_CODPERRO = C.CAMADA_PADRE_CODPERRO
              LEFT JOIN PERROS PM ON PM.PERROS_CODPERRO = C.CAMADA_MADRE_CODPERRO
              LEFT JOIN RAZAS R ON R.RAZA_CODRAZA = PP.PERROS_CODRAZA
              LEFT JOIN ENTIDADES E ON E.ENTIDAD_CODENT = C.CAMADA_CODENT
              WHERE ${w.join(' AND ')}
              ORDER BY C.CAMADA_FEC_INSCRIP DESC`,
        params: (req) => {
          if (p.entidad)    req.input('entidad', sql.Int, parseInt(p.entidad));
          if (p.fechaDesde) req.input('fechaDesde', sql.Date, p.fechaDesde);
          if (p.fechaHasta) req.input('fechaHasta', sql.Date, p.fechaHasta);
        }
      };
    }
  },

  camadas_notificadas: {
    title: 'Camadas Notificadas',
    columns: ['Código','F. Notificación','F. Inscripción','H. Notif.','H. Inscr.','M. Notif.','M. Inscr.','Total Notif.','Total Inscr.','Pendientes','Afijo','Raza','Padre','Madre','Entidad'],
    build: (p) => {
      let w = ['C.CAMADA_FEC_NOTIFICACION IS NOT NULL'];
      if (p.entidad)      w.push('C.CAMADA_CODENT = @entidad');
      if (p.fechaDesde)   w.push('C.CAMADA_FEC_NOTIFICACION >= @fechaDesde');
      if (p.fechaHasta)   w.push('C.CAMADA_FEC_NOTIFICACION <= @fechaHasta');
      if (p.soloPendientes) w.push('(C.CAMADA_FEC_INSCRIP IS NULL OR (C.CAMADA_H_NOTIFICADAS + C.CAMADA_M_NOTIFICADOS > C.CAMADA_H_INSCRITAS + C.CAMADA_M_INSCRITOS))');
      return {
        sql: `SELECT C.CAMADA_CODIGO AS Codigo,
              C.CAMADA_FEC_NOTIFICACION AS FechaNotificacion, C.CAMADA_FEC_INSCRIP AS FechaInscripcion,
              C.CAMADA_H_NOTIFICADAS AS HNotif, C.CAMADA_H_INSCRITAS AS HInscr,
              C.CAMADA_M_NOTIFICADOS AS MNotif, C.CAMADA_M_INSCRITOS AS MInscr,
              (C.CAMADA_H_NOTIFICADAS + C.CAMADA_M_NOTIFICADOS) AS TotalNotif,
              (C.CAMADA_H_INSCRITAS + C.CAMADA_M_INSCRITOS) AS TotalInscr,
              (C.CAMADA_H_NOTIFICADAS + C.CAMADA_M_NOTIFICADOS - C.CAMADA_H_INSCRITAS - C.CAMADA_M_INSCRITOS) AS Pendientes,
              A.AFIJO_NOMBRE AS Afijo, R.RAZA_DENORAZAESP AS Raza,
              PP.PERROS_NOMBRE AS Padre, PM.PERROS_NOMBRE AS Madre,
              E.ENTIDAD_DENO AS Entidad
              FROM CAMADAS C
              LEFT JOIN AFIJOS A ON A.AFIJO_CODIGO = C.CAMADA_AFIJO
              LEFT JOIN PERROS PP ON PP.PERROS_CODPERRO = C.CAMADA_PADRE_CODPERRO
              LEFT JOIN PERROS PM ON PM.PERROS_CODPERRO = C.CAMADA_MADRE_CODPERRO
              LEFT JOIN RAZAS R ON R.RAZA_CODRAZA = PP.PERROS_CODRAZA
              LEFT JOIN ENTIDADES E ON E.ENTIDAD_CODENT = C.CAMADA_CODENT
              WHERE ${w.join(' AND ')}
              ORDER BY C.CAMADA_FEC_NOTIFICACION DESC`,
        params: (req) => {
          if (p.entidad)    req.input('entidad', sql.Int, parseInt(p.entidad));
          if (p.fechaDesde) req.input('fechaDesde', sql.Date, p.fechaDesde);
          if (p.fechaHasta) req.input('fechaHasta', sql.Date, p.fechaHasta);
        }
      };
    }
  },

  pedigris_entidad: {
    title: 'Pedigríes por Entidad',
    columns: ['ID','Fecha Emisión','Tipo','Impreso','Entidad','LOE','Nombre Perro','Raza','Sexo','Apellidos Propietario','Nombre Propietario'],
    build: (p) => {
      let w = [];
      if (p.entidad)    w.push('IP.IMPED_CODENT = @entidad');
      if (p.fechaDesde) w.push('IP.IMPED_FECHA >= @fechaDesde');
      if (p.fechaHasta) w.push('IP.IMPED_FECHA <= @fechaHasta');
      if (p.impreso !== undefined && p.impreso !== '') w.push('IP.IMPED_IMPRESO = @impreso');
      const where = w.length ? 'WHERE ' + w.join(' AND ') : '';
      return {
        sql: `SELECT IP.IMPED_ID AS ID, IP.IMPED_FECHA AS FechaEmision,
              IP.IMPED_TIPOPED AS Tipo,
              CASE WHEN IP.IMPED_IMPRESO = 1 THEN 'Sí' ELSE 'No' END AS Impreso,
              E.ENTIDAD_DENO AS Entidad,
              P.PERROS_CODPERRO AS LOE, P.PERROS_NOMBRE AS NombrePerro,
              R.RAZA_DENORAZAESP AS Raza, P.PERROS_SEXO AS Sexo,
              PER.PERSONA_APELLIDOS AS ApellidosPropietario, PER.PERSONA_NOMBRE AS NombrePropietario
              FROM PERROS_IMPRESION_PEDIGRIS IP
              INNER JOIN ENTIDADES E ON E.ENTIDAD_CODENT = IP.IMPED_CODENT
              INNER JOIN PERROS P ON P.PERROS_CODPERRO = IP.IMPED_CODPERRO
              LEFT JOIN RAZAS R ON R.RAZA_CODRAZA = P.PERROS_CODRAZA
              LEFT JOIN PERSONAS PER ON PER.PERSONA_CODIGO = IP.IMPED_PERSONA
              ${where}
              ORDER BY IP.IMPED_FECHA DESC`,
        params: (req) => {
          if (p.entidad)    req.input('entidad', sql.Int, parseInt(p.entidad));
          if (p.fechaDesde) req.input('fechaDesde', sql.Date, p.fechaDesde);
          if (p.fechaHasta) req.input('fechaHasta', sql.Date, p.fechaHasta);
          if (p.impreso !== undefined && p.impreso !== '') req.input('impreso', sql.Bit, parseInt(p.impreso));
        }
      };
    }
  },

  nuevos_propietarios: {
    title: 'Nuevos Propietarios',
    columns: ['Código','Apellidos','Nombre','DNI','Email','Teléfono','F. Alta','Socio','Dirección','Municipio','C. Postal','Nº Perros'],
    build: (p) => {
      let w = [];
      if (p.fechaDesde)  w.push('PER.PERSONA_FECALTA >= @fechaDesde');
      if (p.fechaHasta)  w.push('PER.PERSONA_FECALTA <= @fechaHasta');
      if (p.soloSocios)  w.push('PER.PERSONA_SOCIO = 1');
      const where = w.length ? 'AND ' + w.join(' AND ') : '';
      return {
        sql: `SELECT PER.PERSONA_CODIGO AS Codigo,
              PER.PERSONA_APELLIDOS AS Apellidos, PER.PERSONA_NOMBRE AS Nombre,
              PER.PERSONA_ID AS DNI, PER.PERSONA_EMAILT AS Email,
              PER.PERSONA_TELEFONOM AS Telefono, PER.PERSONA_FECALTA AS FechaAlta,
              CASE WHEN PER.PERSONA_SOCIO = 1 THEN 'Sí' ELSE 'No' END AS Socio,
              D.DIR_DIRECCION1 AS Direccion, MUN.MUNIC_NOMBRE AS Municipio,
              D.DIR_CODPOS AS CodigoPostal,
              COUNT(DISTINCT PP.PPRO_CODPERRO) AS NumPerros
              FROM PERSONAS PER
              LEFT JOIN DIRECCIONES D ON D.DIR_PERSOCOD = PER.PERSONA_CODIGO AND D.DIR_DIRPREFE = 1
              LEFT JOIN MUNICIPIOS MUN ON MUN.MUNIC_CLAVE = D.DIR_MUNICLAVE
              LEFT JOIN PERROS_PROPIETARIOS PP ON PP.PPRO_IDPERSONA = PER.PERSONA_CODIGO
              WHERE PER.PERSONA_FECALTA IS NOT NULL ${where}
              GROUP BY PER.PERSONA_CODIGO, PER.PERSONA_APELLIDOS, PER.PERSONA_NOMBRE,
              PER.PERSONA_ID, PER.PERSONA_EMAILT, PER.PERSONA_TELEFONOM,
              PER.PERSONA_FECALTA, PER.PERSONA_SOCIO, D.DIR_DIRECCION1,
              MUN.MUNIC_NOMBRE, D.DIR_CODPOS
              ORDER BY PER.PERSONA_FECALTA DESC`,
        params: (req) => {
          if (p.fechaDesde) req.input('fechaDesde', sql.Date, p.fechaDesde);
          if (p.fechaHasta) req.input('fechaHasta', sql.Date, p.fechaHasta);
        }
      };
    }
  },

  afijos: {
    title: 'Afijos',
    columns: ['Código','Nombre Afijo','Pre/Sufijo','Nº FCI','F. Concesión','Apellidos Propietario','Nombre Propietario','Email','Prop. Desde','Prop. Hasta'],
    build: (p) => {
      let w = [];
      if (p.nombre)    w.push("A.AFIJO_NOMBRE LIKE '%' + @nombre + '%'");
      if (p.personaId) w.push('APP.AFIJO_PROPI_PERSONAID = @personaId');
      if (p.raza)      w.push('AR.AFIRAZA_RAZA = @raza');
      const where = w.length ? 'AND ' + w.join(' AND ') : '';
      return {
        sql: `SELECT A.AFIJO_CODIGO AS Codigo, A.AFIJO_NOMBRE AS NombreAfijo,
              A.AFIJO_PRESUFIJO AS PreSufijo, A.AFIJO_NUMFCI AS NumFCI,
              A.AFIJO_FECHACONCE AS FechaConcesion,
              PER.PERSONA_APELLIDOS AS ApellidosPropietario, PER.PERSONA_NOMBRE AS NombrePropietario,
              PER.PERSONA_EMAILT AS Email,
              APP.AFIJO_PROPI_DESDE AS PropDesde, APP.AFIJO_PROPI_HASTA AS PropHasta
              FROM AFIJOS A
              LEFT JOIN AFIJOS_PERIODO_PROPIETARIOS APP ON APP.AFIJO_PROPI_CODIGO = A.AFIJO_CODIGO AND APP.AFIJO_PROPI_HASTA IS NULL
              LEFT JOIN PERSONAS PER ON PER.PERSONA_CODIGO = APP.AFIJO_PROPI_PERSONAID
              LEFT JOIN AFIJOS_RAZAS AR ON AR.AFIRAZA_CODIGO = A.AFIJO_CODIGO
              WHERE 1=1 ${where}
              GROUP BY A.AFIJO_CODIGO, A.AFIJO_NOMBRE, A.AFIJO_PRESUFIJO, A.AFIJO_NUMFCI,
              A.AFIJO_FECHACONCE, PER.PERSONA_APELLIDOS, PER.PERSONA_NOMBRE,
              PER.PERSONA_EMAILT, APP.AFIJO_PROPI_DESDE, APP.AFIJO_PROPI_HASTA
              ORDER BY A.AFIJO_NOMBRE`,
        params: (req) => {
          if (p.nombre)    req.input('nombre', sql.VarChar, p.nombre);
          if (p.personaId) req.input('personaId', sql.Int, parseInt(p.personaId));
          if (p.raza)      req.input('raza', sql.Int, parseInt(p.raza));
        }
      };
    }
  },

  participantes_exposicion: {
    title: 'Participantes en Exposiciones',
    columns: ['Código Expo','Exposición','Fecha','Estado','Recinto','Entidad','Clase','Dorsal','LOE','Nombre Perro','Raza','Sexo','Apellidos Expositor','Nombre Expositor','Email','F. Inscripción','Validada'],
    build: (p) => {
      let w = ['INS.PRU_ANULADA = 0'];
      if (p.exposicion) w.push('INS.PRU_EXPOSICION = @exposicion');
      if (p.entidad)    w.push('EX.PRUEXP_ENTCODIGO = @entidad');
      if (p.raza)       w.push('P.PERROS_CODRAZA = @raza');
      if (p.fechaDesde) w.push('EX.PRUEXP_FECHA >= @fechaDesde');
      if (p.fechaHasta) w.push('EX.PRUEXP_FECHA <= @fechaHasta');
      return {
        sql: `SELECT EX.PRUEXP_CODIGOBELLEZA AS CodigoExpo,
              EX.PRUEXP_DESCRIPCION AS Exposicion, EX.PRUEXP_FECHA AS FechaExposicion,
              EX.PRUEXP_ESTADO AS Estado, EX.PRUEXP_RECINTO AS Recinto,
              ENT.ENTIDAD_DENO AS Entidad,
              INS.PRU_CLASE AS Clase, INS.PRU_DORSAL AS Dorsal,
              P.PERROS_CODPERRO AS LOE, P.PERROS_NOMBRE AS NombrePerro,
              R.RAZA_DENORAZAESP AS Raza, P.PERROS_SEXO AS Sexo,
              PER.PERSONA_APELLIDOS AS ApellidosExpositor, PER.PERSONA_NOMBRE AS NombreExpositor,
              PER.PERSONA_EMAILT AS Email,
              INS.PRU_FECINSCRIPCION AS FechaInscripcion,
              CASE WHEN INS.PRU_VALIDADA = 1 THEN 'Sí' ELSE 'No' END AS Validada
              FROM PRUEBAS_INSCRIPCIONES INS
              INNER JOIN PRUEBAS_EXPOSICIONES EX ON EX.PRUEXP_CODIGOBELLEZA = INS.PRU_EXPOSICION
              INNER JOIN PERSONAS PER ON PER.PERSONA_CODIGO = INS.PRU_EXPOSITOR
              INNER JOIN PERROS P ON P.PERROS_CODPERRO = INS.PRU_PERRO
              LEFT JOIN RAZAS R ON R.RAZA_CODRAZA = P.PERROS_CODRAZA
              LEFT JOIN ENTIDADES ENT ON ENT.ENTIDAD_CODENT = EX.PRUEXP_ENTCODIGO
              WHERE ${w.join(' AND ')}
              ORDER BY EX.PRUEXP_FECHA DESC, R.RAZA_DENORAZAESP, INS.PRU_DORSAL`,
        params: (req) => {
          if (p.exposicion) req.input('exposicion', sql.VarChar, p.exposicion);
          if (p.entidad)    req.input('entidad', sql.Int, parseInt(p.entidad));
          if (p.raza)       req.input('raza', sql.Int, parseInt(p.raza));
          if (p.fechaDesde) req.input('fechaDesde', sql.Date, p.fechaDesde);
          if (p.fechaHasta) req.input('fechaHasta', sql.Date, p.fechaHasta);
        }
      };
    }
  }
};

// ── Endpoint: ejecutar consulta ─────────────────────────────────────────────
app.post('/api/query/:name', async (req, res) => {
  const name = req.params.name;
  if (!QUERIES[name]) return res.status(404).json({ error: 'Consulta no encontrada' });
  try {
    const pool = await getPool();
    const qdef = QUERIES[name].build(req.body);
    const request = pool.request();
    qdef.params(request);
    const result = await request.query(qdef.sql);
    res.json({ rows: result.recordset, columns: QUERIES[name].columns, title: QUERIES[name].title });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ── Endpoint: exportar Excel ────────────────────────────────────────────────
app.post('/api/export/excel/:name', async (req, res) => {
  const name = req.params.name;
  if (!QUERIES[name]) return res.status(404).json({ error: 'Consulta no encontrada' });
  try {
    const pool = await getPool();
    const qdef = QUERIES[name].build(req.body);
    const request = pool.request();
    qdef.params(request);
    const result = await request.query(qdef.sql);
    const rows = result.recordset;
    const cols = QUERIES[name].columns;
    const title = QUERIES[name].title;

    const wb = new ExcelJS.Workbook();
    wb.creator = 'RSCE';
    const ws = wb.addWorksheet(title.slice(0, 31));

    // Título
    ws.mergeCells(1, 1, 1, cols.length);
    ws.getCell('A1').value = title;
    ws.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    ws.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } };
    ws.getCell('A1').alignment = { horizontal: 'center' };
    ws.getRow(1).height = 28;

    // Fecha
    ws.mergeCells(2, 1, 2, cols.length);
    ws.getCell('A2').value = `Generado: ${new Date().toLocaleString('es-ES')}  |  Registros: ${rows.length}`;
    ws.getCell('A2').font = { italic: true, size: 10, color: { argb: 'FF555555' } };
    ws.getCell('A2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD5E8F0' } };
    ws.getRow(2).height = 16;

    // Cabeceras
    const hRow = ws.addRow(cols);
    hRow.eachCell(cell => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E5090' } };
      cell.alignment = { horizontal: 'center', wrapText: true };
      cell.border = { bottom: { style: 'thin', color: { argb: 'FFAAAAAA' } } };
    });
    ws.getRow(3).height = 20;

    // Datos
    if (rows.length > 0) {
      const keys = Object.keys(rows[0]);
      rows.forEach((row, i) => {
        const dRow = ws.addRow(keys.map(k => {
          const v = row[k];
          if (v instanceof Date) return v.toLocaleDateString('es-ES');
          return v;
        }));
        dRow.eachCell(cell => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: i % 2 === 0 ? 'FFFAFAFA' : 'FFFFFFFF' } };
          cell.border = { bottom: { style: 'hair', color: { argb: 'FFDDDDDD' } } };
        });
      });
    }

    // Anchos
    cols.forEach((_, i) => { ws.getColumn(i + 1).width = 18; });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${name}_${Date.now()}.xlsx"`);
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ── Endpoint: exportar PDF ──────────────────────────────────────────────────
app.post('/api/export/pdf/:name', async (req, res) => {
  const name = req.params.name;
  if (!QUERIES[name]) return res.status(404).json({ error: 'Consulta no encontrada' });
  try {
    const pool = await getPool();
    const qdef = QUERIES[name].build(req.body);
    const request = pool.request();
    qdef.params(request);
    const result = await request.query(qdef.sql);
    const rows = result.recordset;
    const cols = QUERIES[name].columns;
    const title = QUERIES[name].title;

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${name}_${Date.now()}.pdf"`);

    const doc = new PDFDocument({ margin: 30, size: 'A4', layout: 'landscape' });
    doc.pipe(res);

    // Encabezado
    doc.rect(0, 0, doc.page.width, 50).fill('#1F3864');
    doc.fillColor('white').fontSize(16).font('Helvetica-Bold')
       .text('RSCE — Real Sociedad Canina de España', 30, 12);
    doc.fontSize(10).font('Helvetica')
       .text(title, 30, 32);

    doc.fillColor('#333333').fontSize(8)
       .text(`Generado: ${new Date().toLocaleString('es-ES')}  |  Registros: ${rows.length}`,
             30, 58);

    // Tabla
    const colW = Math.min(120, Math.floor((doc.page.width - 60) / cols.length));
    const startX = 30;
    let y = 75;
    const rowH = 16;

    // Cabecera tabla
    doc.rect(startX, y, colW * cols.length, rowH).fill('#2E5090');
    doc.fillColor('white').fontSize(7).font('Helvetica-Bold');
    cols.forEach((c, i) => {
      doc.text(c, startX + i * colW + 2, y + 4, { width: colW - 4, ellipsis: true });
    });
    y += rowH;

    // Filas
    if (rows.length > 0) {
      const keys = Object.keys(rows[0]);
      rows.forEach((row, ri) => {
        if (y + rowH > doc.page.height - 30) {
          doc.addPage({ layout: 'landscape' });
          y = 30;
        }
        doc.rect(startX, y, colW * cols.length, rowH)
           .fill(ri % 2 === 0 ? '#F8F8F8' : '#FFFFFF');
        doc.fillColor('#222222').font('Helvetica').fontSize(7);
        keys.forEach((k, i) => {
          let v = row[k];
          if (v instanceof Date) v = v.toLocaleDateString('es-ES');
          if (v === null || v === undefined) v = '';
          doc.text(String(v), startX + i * colW + 2, y + 4, { width: colW - 4, ellipsis: true });
        });
        y += rowH;
      });
    }

    doc.end();
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ── Endpoint: lista de razas (para desplegables) ────────────────────────────
app.get('/api/razas', async (req, res) => {
  try {
    const pool = await getPool();
    const r = await pool.request().query('SELECT RAZA_CODRAZA AS id, RAZA_DENORAZAESP AS nombre FROM RAZAS WHERE RAZA_FECBAJA IS NULL ORDER BY RAZA_DENORAZAESP');
    res.json(r.recordset);
  } catch(e) { res.json([]); }
});

app.get('/api/entidades', async (req, res) => {
  try {
    const pool = await getPool();
    const r = await pool.request().query('SELECT ENTIDAD_CODENT AS id, ENTIDAD_DENO AS nombre FROM ENTIDADES WHERE ENTIDAD_FECHABAJA IS NULL ORDER BY ENTIDAD_DENO');
    res.json(r.recordset);
  } catch(e) { res.json([]); }
});

// ── Arranque ────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n✅  Aplicación RSCE iniciada`);
  console.log(`   Abre en tu navegador: http://localhost:${PORT}\n`);
});
