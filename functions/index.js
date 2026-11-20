const { initializeApp } = require('firebase-admin/app');
const { getFirestore } = require('firebase-admin/firestore');
const { HttpsError, onCall } = require('firebase-functions/v2/https');

initializeApp();

const db = getFirestore();
const appId = 'utepsa-sae-atencion';
const adminEmail = 'oscarjunior0327@gmail.com';

function getTimestampMillis(data) {
  if (typeof data.createdAtServer?.toMillis === 'function') {
    return data.createdAtServer.toMillis();
  }
  if (Number.isFinite(data.createdAt)) {
    return data.createdAt;
  }
  return null;
}

exports.getLatestStudentAttendance = onCall(async (request) => {
  if (!request.auth) {
    throw new HttpsError('unauthenticated', 'Debe iniciar sesión para consultar atenciones.');
  }

  const email = String(request.auth.token.email || '').toLowerCase();
  const isAdmin = email === adminEmail;
  if (!isAdmin) {
    const profileSnapshot = await db.collection('usuarios').doc(request.auth.uid).get();
    const profile = profileSnapshot.data();
    if (
      !profileSnapshot.exists
      || profile.rol !== 'Asesor'
      || String(profile.email || '').toLowerCase() !== email
    ) {
      throw new HttpsError('permission-denied', 'El usuario no tiene un perfil de asesor válido.');
    }
  }

  const studentCode = request.data?.studentCode;
  if (typeof studentCode !== 'string' || !/^(?:\d{4,6}|\d{5}A)$/.test(studentCode)) {
    throw new HttpsError('invalid-argument', 'El código debe contener de cuatro a seis dígitos o cinco dígitos seguidos de A.');
  }
  const targetMonth = request.data?.targetMonth;
  if (
    targetMonth !== undefined
    && (typeof targetMonth !== 'string' || !/^\d{4}-\d{2}$/.test(targetMonth))
  ) {
    throw new HttpsError('invalid-argument', 'El mes debe tener el formato YYYY-MM.');
  }

  const records = await db
    .collection('artifacts')
    .doc(appId)
    .collection('public')
    .doc('data')
    .collection('atenciones')
    .where('registro', '==', studentCode)
    .get();

  const matchingRecords = records.docs.filter((record) => {
    if (!targetMonth) return true;
    const data = record.data();
    const recordMonth = typeof data.mesAnio === 'string'
      ? data.mesAnio
      : typeof data.fecha === 'string'
        ? data.fecha.slice(0, 7)
        : '';
    return recordMonth === targetMonth;
  });

  if (matchingRecords.length === 0) {
    return { found: false };
  }

  const latest = matchingRecords
    .map((record) => {
      const data = record.data();
      const createdAtMillis = getTimestampMillis(data);
      const fechaMillis = /^\d{4}-\d{2}-\d{2}$/.test(data.fecha || '')
        ? Date.parse(`${data.fecha}T00:00:00Z`)
        : 0;
      return {
        asesor: String(data.asesor || '').trim() || 'Asesor SAE',
        fecha: typeof data.fecha === 'string' ? data.fecha : '',
        createdAtMillis,
        sortMillis: createdAtMillis ?? fechaMillis
      };
    })
    .sort((left, right) => right.sortMillis - left.sortMillis)[0];

  return {
    found: true,
    asesor: latest.asesor,
    fecha: latest.fecha,
    createdAtMillis: latest.createdAtMillis
  };
});
