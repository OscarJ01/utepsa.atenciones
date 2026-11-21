import { initializeApp, deleteApp, getApps } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-app.js";
    import {
      getAuth,
      signInWithCustomToken,
      signInWithEmailAndPassword,
      signOut
    } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-auth.js";
    import {
      initializeFirestore,
      persistentLocalCache,
      persistentMultipleTabManager,
      collection,
      doc,
      getDoc,
      setDoc,
      updateDoc,
      deleteDoc,
      getDocs,
      onSnapshot,
      query,
      where,
      writeBatch,
      serverTimestamp,
      Timestamp
    } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-firestore.js";

    // =========================================================================
    // 1. CONSTANTES INSTITUCIONALES Y CATÁLOGO EXACTO DEL EXCEL UTEPSA
    // =========================================================================
    const PROCEDURES = [
      { key: 'pensum', label: 'PENSUM', short: 'PENSUM', hotkey: '1' },
      { key: 'programacion', label: 'PROGRAMACION', short: 'PROGR.', hotkey: '2' },
      { key: 'confirmacion', label: 'CONFIRMACION', short: 'CONFIR.', hotkey: '3' },
      { key: 'plan_pago', label: 'PLAN DE PAGO', short: 'PLAN PAGO', hotkey: '4' },
      { key: 'recargo', label: 'RECARGO', short: 'RECARGO', hotkey: '5' },
      { key: 'congelamiento', label: 'CONGELAMIENTO', short: 'CONGEL.', hotkey: '6' },
      { key: 'pro_forma', label: 'PRO-FORMA', short: 'PRO-FORMA', hotkey: '7' },
      { key: 'abonos', label: 'ABONOS', short: 'ABONOS', hotkey: '8' },
      { key: 'informacion', label: 'INFORMACION', short: 'INFORM.', hotkey: '9' },
      { key: 'reingresante', label: 'REINGRESANTE', short: 'REINGR.', hotkey: '0' },
      { key: 'ex_grado', label: 'EX. DE GRADO / TESIS', short: 'EX. GRADO', hotkey: 'Q' },
    ];

    const QUICK_OBS_TAGS = [
      'MOD 1',
      'MOD 2',
      'MOD 3',
      'MOD 4',
      'MOD 5',
      'HISTORIAL DE NOTAS',
      'DEFENSA EXTERNA',
      'INFORMACION'
    ];

    const MONTH_NAMES_ES = [
      'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
      'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
    ];

    // Identificador de App para rutas estrictas de Firestore
    const appId = typeof __app_id !== 'undefined' ? __app_id : 'utepsa-sae-atencion';
    const FIREBASE_CONFIG = {
      apiKey: 'AIzaSyCZYmxKIiXXYb-iAm7o-jiPbJ9RX1e3VBg',
      authDomain: 'portalsaeutepsa.firebaseapp.com',
      projectId: 'portalsaeutepsa',
      storageBucket: 'portalsaeutepsa.firebasestorage.app',
      messagingSenderId: '448731681018',
      appId: '1:448731681018:web:42b72099bfa63eddda3943'
    };
    const INITIAL_ADMIN_EMAIL = 'oscarjunior0327@gmail.com';
    const CHANNEL_MODE_STORAGE_KEY_PREFIX = 'utepsa-sae-channel-mode:';

    function loadSavedChannelMode(uid) {
      try {
        return window.localStorage.getItem(`${CHANNEL_MODE_STORAGE_KEY_PREFIX}${uid}`) === 'wati'
          ? 'wati'
          : 'plataforma';
      } catch (error) {
        console.error('No se pudo recuperar el canal de atención guardado:', error);
        showToast('No se pudo recuperar el canal de atención anterior.', 'error');
        return 'plataforma';
      }
    }

    function saveChannelMode(uid, channel) {
      try {
        window.localStorage.setItem(`${CHANNEL_MODE_STORAGE_KEY_PREFIX}${uid}`, channel);
      } catch (error) {
        console.error('No se pudo guardar el canal de atención:', error);
        showToast('No se pudo guardar el canal de atención para la próxima actualización.', 'error');
      }
    }

    // =========================================================================
    // 2. ESTADO GLOBAL EN MEMORIA
    // =========================================================================
    const state = {
      currentUser: null, // { name, email, role, uid }
      channelAtencion: 'plataforma',
      activeTab: 'registro',
      records: [], // Todos los registros sincronizados (Firestore o sesión activa)
      selectedProcedures: new Set(),
      editSelectedProcedures: new Set(),
      historyPage: 1,
      charts: {
        procedures: null,
        daily: null,
        advisors: null
      },
      firebase: {
        app: null,
        auth: null,
        db: null,
        initialized: false,
        connected: false,
        pendingWrites: 0,
        unsubscribeSnapshot: null,
        summaryBackfillStarted: false
      },
      latestAttendanceLookup: { key: '', promise: null },
      latestAttendanceRequestId: 0
    };

    function normalizeStudentCodeInput(value) {
      let code = '';
      for (const character of String(value).toUpperCase()) {
        if (/\d/.test(character) && code.length < 6 && !code.endsWith('A')) {
          code += character;
        } else if (character === 'A' && code.length === 5) {
          code += character;
        }
      }
      return code;
    }

    function isValidStudentCode(value) {
      return /^(?:\d{4,6}|\d{5}A)$/.test(value);
    }

    // =========================================================================
    // 3. INICIALIZACIÓN Y GESTIÓN DE FIREBASE (REGLAS 1, 2 Y 3)
    // =========================================================================
    function finishInitialLoading() {
      document.getElementById('app-loading-screen')?.classList.add('hidden');
      if (!state.currentUser) {
        document.getElementById('login-screen')?.classList.remove('hidden');
      }
    }

    async function initFirebaseFromConfig(configObj) {
      try {
        if (!configObj || !configObj.apiKey || !configObj.projectId) {
          finishInitialLoading();
          return false;
        }

        // Limpiar suscripción previa si existía
        if (state.firebase.unsubscribeSnapshot) {
          state.firebase.unsubscribeSnapshot();
          state.firebase.unsubscribeSnapshot = null;
        }

        // Eliminar instancias previas si se reconfigura dinámicamente
        const existingApps = getApps();
        for (const existing of existingApps) {
          await deleteApp(existing);
        }

        const app = initializeApp(configObj);
        const auth = getAuth(app);
        const db = initializeFirestore(app, {
          localCache: persistentLocalCache({
            tabManager: persistentMultipleTabManager()
          })
        });

        state.firebase.app = app;
        state.firebase.auth = auth;
        state.firebase.db = db;
        state.firebase.initialized = true;
        state.firebase.connected = false;
        updateFirebaseStatusUI(false, configObj.projectId, true);

        await auth.authStateReady();

        if (typeof __initial_auth_token !== 'undefined' && __initial_auth_token) {
          await signInWithCustomToken(auth, __initial_auth_token);
          state.firebase.connected = false;
          updateFirebaseStatusUI(false, configObj.projectId, true);
        }

        if (auth.currentUser) {
          await completeAuthentication(auth.currentUser, auth.currentUser.email || '');
        }

        finishInitialLoading();
        return true;
      } catch (err) {
        console.error('Error al inicializar Firebase:', err);
        state.firebase.initialized = false;
        state.firebase.connected = false;
        updateFirebaseStatusUI(false);
        finishInitialLoading();
        const message = err?.code === 'auth/operation-not-allowed'
          ? 'Habilite el proveedor Correo/Contraseña en Authentication del proyecto Firebase.'
          : err?.code === 'auth/unauthorized-domain'
            ? 'Este dominio no está autorizado en Firebase Authentication. Agregue localhost y 127.0.0.1 a Dominios autorizados.'
            : err?.code === 'unimplemented'
              ? 'Este navegador no permite guardar registros pendientes sin conexión.'
              : err?.code === 'failed-precondition'
                ? 'No se pudo habilitar el almacenamiento local de Firebase. Cierre otras pestañas de la aplicación e inténtelo de nuevo.'
                : 'No se pudo conectar a Firebase. Revise la configuración y los proveedores de autenticación.';
        showToast(message, 'error');
        return false;
      }
    }

    function startFirestoreRealtimeSync() {
      // Guard estricto (Regla 3)
      if (!state.firebase.db || !state.firebase.auth?.currentUser) return;

      if (state.firebase.unsubscribeSnapshot) {
        state.firebase.unsubscribeSnapshot();
      }

      const recordsRef = collection(state.firebase.db, 'artifacts', appId, 'public', 'data', 'atenciones');
      const colRef = state.currentUser?.role === 'Asesor'
        ? query(recordsRef, where('asesorUid', '==', state.currentUser.uid))
        : recordsRef;

      state.firebase.unsubscribeSnapshot = onSnapshot(
        colRef,
        { includeMetadataChanges: true },
        (snapshot) => {
          const loaded = [];
          snapshot.forEach((docSnap) => {
            const data = docSnap.data();
            loaded.push({
              id: docSnap.id,
              fecha: data.fecha || getTodayISO(),
              mesAnio: data.mesAnio || (data.fecha ? data.fecha.slice(0, 7) : getTodayISO().slice(0, 7)),
              registro: String(data.registro || '').trim(),
              tramites: data.tramites || {},
              observaciones: data.observaciones || '',
              asesor: data.asesor || 'Oscar Vargas',
              asesorUid: data.asesorUid || '',
              canalAtencion: data.canalAtencion === 'wati' ? 'wati' : 'plataforma',
              createdAt: Number.isFinite(data.createdAt) ? data.createdAt : 0,
              createdAtServer: data.createdAtServer || null,
              pendingSync: docSnap.metadata.hasPendingWrites
            });
          });
          // Ordenar en memoria de forma descendente por fecha y marca de tiempo
          loaded.sort((a, b) => {
            if (b.fecha !== a.fecha) return b.fecha.localeCompare(a.fecha);
            return (b.createdAt || 0) - (a.createdAt || 0);
          });
          state.records = loaded;
          state.firebase.pendingWrites = loaded.filter((record) => record.pendingSync).length;
          state.firebase.connected = !snapshot.metadata.fromCache;
          if (
            state.currentUser?.role === 'Administrador'
            && snapshot.docChanges({ includeMetadataChanges: false }).length > 0
          ) {
            state.firebase.summaryBackfillStarted = false;
          }
          if (
            state.currentUser?.role === 'Administrador'
            && !state.firebase.summaryBackfillStarted
            && !snapshot.metadata.fromCache
          ) {
            state.firebase.summaryBackfillStarted = true;
            syncLatestAttendanceSummaries(loaded)
              .then((updatedCount) => {
                if (updatedCount > 0) {
                  showToast('Se preparó el historial de últimas atenciones para los asesores.', 'success');
                }
              })
              .catch((error) => {
                console.error('Error preparando resúmenes de última atención:', error);
                state.firebase.summaryBackfillStarted = false;
                showToast('No se pudieron preparar los resúmenes de última atención.', 'error');
              });
          }
          updateFirebaseStatusUI(
            state.firebase.connected,
            state.firebase.app?.options.projectId,
            true
          );
          refreshAllViews();
        },
        (error) => {
          console.error('Error en sincronización Firestore:', error);
          state.firebase.connected = false;
          updateFirebaseStatusUI(false, state.firebase.app?.options.projectId, true);
          const message = error?.code === 'permission-denied'
            ? 'Firestore denegó el acceso de lectura. Revise las reglas para la ruta de atenciones.'
            : `Error al sincronizar Firestore (${error?.code || 'error-desconocido'}).`;
          showToast(message, 'error');
        }
      );
    }

    function getAttendanceTimestampMillis(record) {
      if (typeof record.createdAtServer?.toMillis === 'function') {
        return record.createdAtServer.toMillis();
      }
      if (Number.isFinite(record.createdAt) && record.createdAt > 0) return record.createdAt;
      return /^\d{4}-\d{2}-\d{2}$/.test(record.fecha || '')
        ? Date.parse(`${record.fecha}T00:00:00Z`)
        : 0;
    }

    function toAttendanceTimestamp(record) {
      if (typeof record.createdAtServer?.toMillis === 'function') return record.createdAtServer;
      return Timestamp.fromMillis(getAttendanceTimestampMillis(record));
    }

    function getLatestAttendanceSummaryRef(studentCode) {
      return doc(
        state.firebase.db,
        'artifacts',
        appId,
        'public',
        'data',
        'ultimaAtencion',
        studentCode
      );
    }

    function getAdvisorLatestAttendanceSummaryRef(studentCode) {
      return doc(
        state.firebase.db,
        'artifacts',
        appId,
        'public',
        'data',
        'ultimaAtencionAsesores',
        studentCode
      );
    }

    function createAttendanceSummaryData(studentCode, record) {
      return {
        registro: studentCode,
        asesor: record.asesor,
        asesorUid: record.asesorUid,
        fecha: record.fecha,
        mesAnio: record.mesAnio,
        createdAt: record.createdAt,
        createdAtServer: toAttendanceTimestamp(record),
        recordId: record.id
      };
    }

    function createAdvisorAttendanceSummaryData(studentCode, record) {
      return {
        registro: studentCode,
        asesor: record.asesor,
        fecha: record.fecha,
        createdAtServer: record.createdAtServer || toAttendanceTimestamp(record),
        recordId: record.id
      };
    }

    async function updateLatestAttendanceSummary(recordRef, recordData) {
      const savedRecordSnapshot = await getDoc(recordRef);
      if (!savedRecordSnapshot.exists()) {
        throw new Error('El registro no aparece en Firestore después de guardarse.');
      }

      const savedRecord = savedRecordSnapshot.data();
      const savedTimestamp = savedRecord.createdAtServer;
      if (typeof savedTimestamp?.toMillis !== 'function') {
        throw new Error('El registro guardado no contiene la marca de tiempo del servidor.');
      }

      const summaryRef = getLatestAttendanceSummaryRef(recordData.registro);
      const currentSummary = await getDoc(summaryRef);
      const currentData = currentSummary.data();
      const currentTimestamp = currentData?.createdAtServer;
      if (
        typeof currentTimestamp?.toMillis === 'function'
        && (
          currentTimestamp.toMillis() > savedTimestamp.toMillis()
          || (
            currentTimestamp.toMillis() === savedTimestamp.toMillis()
            && currentData.recordId >= recordRef.id
          )
        )
      ) return;

      const summaryData = createAttendanceSummaryData(recordData.registro, {
        ...recordData,
        createdAtServer: savedTimestamp,
        id: recordRef.id
      });
      try {
        await setDoc(summaryRef, summaryData);
      } catch {
        const latestSummary = await getDoc(summaryRef);
        const latestData = latestSummary.data();
        const latestTimestamp = latestData?.createdAtServer;
        if (
          typeof latestTimestamp?.toMillis === 'function'
          && (
            latestTimestamp.toMillis() > savedTimestamp.toMillis()
            || (
              latestTimestamp.toMillis() === savedTimestamp.toMillis()
              && latestData.recordId >= recordRef.id
            )
          )
        ) return;
        try {
          await setDoc(summaryRef, summaryData);
        } catch (retryError) {
          console.error('Error reintentando el resumen de última atención:', retryError);
          throw retryError;
        }
      }
    }

    async function syncLatestAttendanceSummaries(records) {
      const summariesByStudent = new Map();
      records.forEach((record) => {
        if (record.pendingSync || !isValidStudentCode(record.registro)) return;
        const current = summariesByStudent.get(record.registro);
        if (
          !current
          || getAttendanceTimestampMillis(record) > getAttendanceTimestampMillis(current)
          || (
            getAttendanceTimestampMillis(record) === getAttendanceTimestampMillis(current)
            && record.id > current.id
          )
        ) {
          summariesByStudent.set(record.registro, record);
        }
      });

      const summaryCollection = collection(
        state.firebase.db,
        'artifacts',
        appId,
        'public',
        'data',
        'ultimaAtencion'
      );
      const existingSummaries = await getDocs(summaryCollection);
      const currentSummaryByStudent = new Map();
      existingSummaries.forEach((summary) => {
        currentSummaryByStudent.set(summary.id, summary.data());
      });
      const advisorSummaryCollection = collection(
        state.firebase.db,
        'artifacts',
        appId,
        'public',
        'data',
        'ultimaAtencionAsesores'
      );
      const existingAdvisorSummaries = await getDocs(advisorSummaryCollection);
      const currentAdvisorSummaryByStudent = new Map();
      existingAdvisorSummaries.forEach((summary) => {
        currentAdvisorSummaryByStudent.set(summary.id, summary.data());
      });

      const changes = [];
      for (const [studentCode, record] of summariesByStudent) {
        const current = currentSummaryByStudent.get(studentCode);
        const recordTimestamp = toAttendanceTimestamp(record);
        const privateSummaryMatches = Boolean(
          current
          && current.recordId === record.id
          && current.createdAtServer?.toMillis?.() === recordTimestamp.toMillis()
          && current.asesor === record.asesor
          && current.asesorUid === record.asesorUid
          && current.fecha === record.fecha
          && current.mesAnio === record.mesAnio
          && current.createdAt === record.createdAt
        );
        const privateSummaryTimestamp = current?.createdAtServer?.toMillis?.() ?? 0;
        const privateSummaryIsNewer = current
          && (
            privateSummaryTimestamp > recordTimestamp.toMillis()
            || (
              privateSummaryTimestamp === recordTimestamp.toMillis()
              && current.recordId > record.id
            )
          );
        if (!privateSummaryMatches && !privateSummaryIsNewer) {
          changes.push({
            type: 'set',
            ref: getLatestAttendanceSummaryRef(studentCode),
            data: createAttendanceSummaryData(studentCode, record)
          });
        }

        const advisorSummary = createAdvisorAttendanceSummaryData(studentCode, record);
        const currentAdvisorSummary = currentAdvisorSummaryByStudent.get(studentCode);
        const advisorSummaryMatches = Boolean(
          currentAdvisorSummary
          && currentAdvisorSummary.registro === advisorSummary.registro
          && currentAdvisorSummary.asesor === advisorSummary.asesor
          && currentAdvisorSummary.fecha === advisorSummary.fecha
          && currentAdvisorSummary.recordId === advisorSummary.recordId
          && currentAdvisorSummary.createdAtServer?.toMillis?.()
            === advisorSummary.createdAtServer.toMillis()
        );
        const advisorSummaryTimestamp =
          currentAdvisorSummary?.createdAtServer?.toMillis?.() ?? 0;
        const advisorSummaryIsNewer = currentAdvisorSummary
          && (
            advisorSummaryTimestamp > recordTimestamp.toMillis()
            || (
              advisorSummaryTimestamp === recordTimestamp.toMillis()
              && currentAdvisorSummary.recordId > record.id
            )
          );
        if (!advisorSummaryMatches && !advisorSummaryIsNewer) {
          changes.push({
            type: 'set',
            ref: getAdvisorLatestAttendanceSummaryRef(studentCode),
            data: advisorSummary
          });
        }
      }

      for (const [studentCode] of currentSummaryByStudent) {
        if (summariesByStudent.has(studentCode)) continue;
        changes.push({
          type: 'delete',
          ref: getLatestAttendanceSummaryRef(studentCode)
        });
      }

      for (const [studentCode] of currentAdvisorSummaryByStudent) {
        if (summariesByStudent.has(studentCode)) continue;
        changes.push({
          type: 'delete',
          ref: getAdvisorLatestAttendanceSummaryRef(studentCode)
        });
      }

      for (let index = 0; index < changes.length; index += 10) {
        const batch = writeBatch(state.firebase.db);
        changes.slice(index, index + 10).forEach((change) => {
          if (change.type === 'delete') {
            batch.delete(change.ref);
          } else {
            batch.set(change.ref, change.data);
          }
        });
        await batch.commit();
      }
      return changes.length;
    }

    function updateFirebaseStatusUI(isConnected, projectId = '', isReady = false) {
      const loginDot = document.getElementById('login-fb-dot');
      const loginText = document.getElementById('login-fb-text');
      const headerDot = document.getElementById('header-fb-dot');
      const headerText = document.getElementById('header-fb-text');

      if (isConnected) {
        loginDot.className = 'w-2 h-2 rounded-full bg-emerald-500 animate-pulse';
        loginText.textContent = `Firebase Conectado (${projectId})`;
        headerDot.className = 'w-2 h-2 rounded-full bg-emerald-500 animate-pulse';
        headerText.textContent = state.firebase.pendingWrites > 0
          ? `Sincronizando (${state.firebase.pendingWrites})`
          : 'Firebase Live';
      } else if (isReady) {
        loginDot.className = 'w-2 h-2 rounded-full bg-amber-500';
        loginText.textContent = state.firebase.auth?.currentUser
          ? 'Firebase Auth conectado · Firestore pendiente'
          : 'Firebase listo';
        headerDot.className = 'w-2 h-2 rounded-full bg-amber-500';
        headerText.textContent = state.firebase.pendingWrites > 0
          ? `Sin conexión · ${state.firebase.pendingWrites} pendiente(s)`
          : state.firebase.auth?.currentUser ? 'Firestore pendiente' : 'Firebase listo';
      } else {
        loginDot.className = 'w-2 h-2 rounded-full bg-amber-500';
        loginText.textContent = 'Sin conexión con Firebase';
        headerDot.className = 'w-2 h-2 rounded-full bg-amber-500';
        headerText.textContent = 'Sin conexión';
      }
    }

    // =========================================================================
    // 4. AUTENTICACIÓN Y NAVEGACIÓN
    // =========================================================================
    async function completeAuthentication(authUser, emailFallback) {
      const email = String(authUser.email || emailFallback).trim().toLowerCase();
      const isInitialAdmin = email === INITIAL_ADMIN_EMAIL;
      let profile = {};

      try {
        const profileSnapshot = await getDoc(doc(state.firebase.db, 'usuarios', authUser.uid));
        if (profileSnapshot.exists()) {
          profile = profileSnapshot.data();
        }
      } catch (err) {
        console.error('Error leyendo el perfil del usuario:', err);
        showToast(`No se pudo cargar el perfil desde Firestore (${err?.code || 'error-desconocido'}).`, 'error');
        return;
      }

      const authEmail = String(authUser.email || emailFallback).trim().toLowerCase();
      const profileEmail = String(profile.email || (isInitialAdmin ? authEmail : '')).trim().toLowerCase();
      const isAdmin = isInitialAdmin
        || (profile.rol === 'Administrador' && profileEmail === authEmail);
      const isAdvisor = profile.rol === 'Asesor' && profileEmail === authEmail;
      const role = isAdmin ? 'Administrador' : isAdvisor ? 'Asesor' : '';
      const name = String(profile.nombre || (isInitialAdmin ? 'Oscar Vargas' : '')).trim();
      const phone = String(profile.telefono || '').trim();

      if (!role || !name || profileEmail !== authEmail || (!isAdmin && !phone)) {
        if (!isInitialAdmin) {
          await signOut(state.firebase.auth);
        }
        showToast(
          isInitialAdmin
            ? 'El perfil del administrador no coincide con su cuenta de Firebase.'
            : 'No existe un perfil válido. Registre nombre, teléfono, correo y rol "Asesor" o "Administrador" en Firestore.',
          'error'
        );
        return;
      }

      state.firebase.connected = false;
      state.records = [];
      state.latestAttendanceLookup = { key: '', promise: null };
      state.latestAttendanceRequestId += 1;
      updateFirebaseStatusUI(false, state.firebase.app?.options.projectId, true);
      enterApplication({
        uid: authUser.uid,
        name,
        email: authUser.email || emailFallback,
        phone,
        role
      });
      startFirestoreRealtimeSync();
    }

    window.handleAuthSubmit = async function (e) {
      e.preventDefault();
      const email = document.getElementById('auth-email').value.trim().toLowerCase();
      const password = document.getElementById('auth-password').value;

      if (!state.firebase.initialized || !state.firebase.auth) {
        showToast('Firebase no está disponible. Verifique la conexión e inténtelo de nuevo.', 'error');
        return;
      }

      let credential;
      try {
        credential = await signInWithEmailAndPassword(state.firebase.auth, email, password);
      } catch (err) {
        console.error('Error al autenticar con Firebase:', err);
        state.firebase.connected = false;
        updateFirebaseStatusUI(false, state.firebase.app?.options.projectId, true);
        const message = err?.code === 'auth/invalid-credential' || err?.code === 'auth/wrong-password' || err?.code === 'auth/user-not-found'
          ? 'Correo o contraseña incorrectos.'
          : err?.code === 'auth/operation-not-allowed'
            ? 'El proveedor Correo/Contraseña no está habilitado en Firebase Authentication.'
            : 'Firebase rechazó el inicio de sesión. Verifique la cuenta y la configuración de Authentication.';
        showToast(message, 'error');
        return;
      }

      await completeAuthentication(credential.user, email);
    };

    function enterApplication(user) {
      state.currentUser = user;
      state.channelAtencion = loadSavedChannelMode(user.uid);
      updateChannelModeUI();
      updateReportAdvisorsHeading();
      state.firebase.summaryBackfillStarted = false;
      document.getElementById('login-screen').classList.add('hidden');
      const appShell = document.getElementById('app-shell');
      appShell.classList.remove('hidden');
      appShell.classList.add('flex');

      // Actualizar cabecera de usuario
      document.getElementById('header-user-name').textContent = user.name;
      document.getElementById('header-user-role').textContent = `${user.role} SAE`;
      const initials = user.name
        .split(' ')
        .slice(0, 2)
        .map((n) => n[0])
        .join('')
        .toUpperCase();
      document.getElementById('user-avatar-initials').textContent = initials || 'US';

      // Inicializar fecha por defecto hoy
      const today = getTodayISO();
      const entryDateInput = document.getElementById('entry-fecha');
      entryDateInput.value = today;
      const canEditAttendanceDate = user.role === 'Administrador';
      setDateFieldEditable(entryDateInput, canEditAttendanceDate);
      document.getElementById('entry-fecha-help').textContent = canEditAttendanceDate
        ? 'El administrador puede ajustar la fecha de atención'
        : 'Fecha automática; los asesores no pueden modificarla';
      document.getElementById('quick-stat-today-date').textContent = formatDateDisplay(today);

      // Si es supervisor, abrir directamente en Reporte Mensual; si es asesor, en Registro Rápido
      refreshAllViews();
      if (user.role === 'Supervisor' || user.role === 'Administrador') {
        switchTab('reportes');
      } else {
        switchTab('registro');
        setTimeout(() => document.getElementById('entry-registro')?.focus(), 100);
      }
    }

    window.confirmLogout = function () {
      openConfirmModal(
        'Cerrar Sesión',
        '¿Desea salir del sistema de atención estudiantil UTEPSA SAE?',
        async () => {
          if (state.firebase.auth) {
            try {
              await signOut(state.firebase.auth);
            } catch (err) {
              console.error('Error al cerrar sesión de Firebase:', err);
              showToast('No se pudo cerrar la sesión de Firebase.', 'error');
            }
          }
          if (state.firebase.unsubscribeSnapshot) {
            state.firebase.unsubscribeSnapshot();
            state.firebase.unsubscribeSnapshot = null;
          }
          state.firebase.connected = false;
          if (state.firebase.initialized) {
            updateFirebaseStatusUI(false, state.firebase.app?.options.projectId, true);
          }
          state.currentUser = null;
          state.channelAtencion = 'plataforma';
          updateChannelModeUI();
          state.records = [];
          state.latestAttendanceLookup = { key: '', promise: null };
          state.latestAttendanceRequestId += 1;
          state.firebase.summaryBackfillStarted = false;
          document.getElementById('app-shell').classList.add('hidden');
          document.getElementById('app-shell').classList.remove('flex');
          document.getElementById('login-screen').classList.remove('hidden');
          showToast('Sesión cerrada correctamente.', 'info');
        }
      );
    };

    window.switchTab = function (tabName) {
      state.activeTab = tabName;
      const tabs = ['registro', 'historial', 'reportes'];

      tabs.forEach((t) => {
        const viewEl = document.getElementById(`view-${t}`);
        const navBtn = document.getElementById(`nav-tab-${t}`);
        const mobBtn = document.getElementById(`mob-tab-${t}`);

        if (t === tabName) {
          viewEl.classList.remove('hidden');
          navBtn.className = 'px-4 py-2 rounded-lg text-xs font-bold transition flex items-center gap-2 bg-[#C8102E] text-white shadow-sm';
          mobBtn.className = 'py-2.5 text-xs font-bold text-[#C8102E] border-b-2 border-[#C8102E] bg-white';
        } else {
          viewEl.classList.add('hidden');
          navBtn.className = 'px-4 py-2 rounded-lg text-xs font-bold transition flex items-center gap-2 text-slate-600 hover:text-slate-900';
          mobBtn.className = 'py-2.5 text-xs font-bold text-slate-600';
        }
      });

      if (tabName === 'historial') {
        renderHistoryTable();
      } else if (tabName === 'reportes') {
        renderMonthlyReport();
      } else if (tabName === 'registro') {
        renderTodayFeed();
        document.getElementById('entry-registro')?.focus();
      }
    };

    function updateChannelModeUI() {
      const control = document.getElementById('channel-mode-control');
      const button = document.getElementById('channel-mode-toggle');
      const indicator = document.getElementById('channel-mode-indicator');
      const label = document.getElementById('channel-mode-label');
      if (!control || !button || !indicator || !label) return;

      const canChooseChannel = ['Asesor', 'Administrador'].includes(state.currentUser?.role);
      const isWati = state.channelAtencion === 'wati';
      control.classList.toggle('hidden', !canChooseChannel);
      label.textContent = isWati ? 'WATI' : 'Plataforma';
      button.setAttribute('aria-checked', String(isWati));
      button.setAttribute(
        'aria-label',
        `Canal de atención: ${isWati ? 'WATI' : 'Plataforma'}`
      );
      button.className = isWati
        ? 'inline-flex items-center gap-2 px-2.5 py-2 rounded-xl text-[10px] sm:text-xs font-bold border bg-emerald-50 text-emerald-800 border-emerald-200 hover:bg-emerald-100 transition'
        : 'inline-flex items-center gap-2 px-2.5 py-2 rounded-xl text-[10px] sm:text-xs font-bold border bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100 transition';
      indicator.className = isWati
        ? 'w-2.5 h-2.5 rounded-full bg-emerald-500'
        : 'w-2.5 h-2.5 rounded-full bg-slate-400';
    }

    window.toggleChannelMode = function () {
      if (!['Asesor', 'Administrador'].includes(state.currentUser?.role)) return;
      state.channelAtencion = state.channelAtencion === 'wati' ? 'plataforma' : 'wati';
      saveChannelMode(state.currentUser.uid, state.channelAtencion);
      updateChannelModeUI();
    };

    // =========================================================================
    // 6. REGISTRO RÁPIDO Y DETECCIÓN AUTOMÁTICA DE RE-CONSULTAS (NARANJA)
    // =========================================================================
    function renderProcedureButtons() {
      const container = document.getElementById('procedures-grid');
      container.innerHTML = PROCEDURES.map((proc) => {
        const active = state.selectedProcedures.has(proc.key);
        return `
          <button
            type="button"
            onclick="toggleProcedure('${proc.key}')"
            class="p-3 rounded-2xl border-2 text-left transition flex items-center justify-between gap-2 select-none ${
              active
                ? 'bg-[#C8102E] border-[#C8102E] text-white shadow-md shadow-red-600/15'
                : 'bg-slate-50/70 border-slate-200 text-slate-700 hover:border-[#C8102E]/60 hover:bg-white'
            }"
          >
            <div class="flex items-center gap-2.5 min-w-0">
              <span class="w-6 h-6 rounded-lg flex items-center justify-center font-black text-xs shrink-0 ${
                active ? 'bg-white text-[#C8102E]' : 'bg-white border border-slate-300 text-slate-400'
              }">
                ${active ? 'X' : ''}
              </span>
              <span class="text-xs font-extrabold tracking-tight truncate">${proc.label}</span>
            </div>
            <span class="text-[10px] font-bold px-1.5 py-0.5 rounded ${
              active ? 'bg-black/20 text-white' : 'bg-slate-200/70 text-slate-500'
            }">${proc.hotkey}</span>
          </button>
        `;
      }).join('');

      document.getElementById('selected-procs-count').textContent = `${state.selectedProcedures.size} seleccionado${state.selectedProcedures.size === 1 ? '' : 's'}`;
    }

    window.toggleProcedure = function (key) {
      if (state.selectedProcedures.has(key)) {
        state.selectedProcedures.delete(key);
      } else {
        state.selectedProcedures.add(key);
      }
      renderProcedureButtons();
    };

    function renderQuickObsTags() {
      const container = document.getElementById('quick-obs-tags');
      container.innerHTML = QUICK_OBS_TAGS.map((tag) => `
        <button
          type="button"
          onclick="applyQuickObsTag('${tag}')"
          class="px-2.5 py-1 rounded-lg text-[11px] font-bold bg-slate-100 hover:bg-[#FDF2F4] text-slate-700 hover:text-[#C8102E] border border-slate-200 hover:border-red-200 transition"
        >
          + ${tag}
        </button>
      `).join('');
    }

    window.applyQuickObsTag = function (tag) {
      const input = document.getElementById('entry-observaciones');
      const current = input.value.trim();
      const moduleMatch = /^MOD\s+(\d+)$/i.exec(tag);
      const observationTag = moduleMatch ? `Modulo ${moduleMatch[1]}` : tag;
      if (!current) {
        input.value = observationTag;
      } else if (!current.toUpperCase().includes(observationTag.toUpperCase())) {
        input.value = `${current} - ${observationTag}`;
      }
      input.focus();
    };

    function getStudentAttendanceSummary(studentCode) {
      const key = studentCode;
      const cachedLookup = state.latestAttendanceLookup;
      if (cachedLookup.key === key && cachedLookup.promise) return cachedLookup.promise;

      let promise;
      promise = getDoc(getAdvisorLatestAttendanceSummaryRef(studentCode))
        .then((snapshot) => {
          if (!snapshot.exists()) return { found: false };
          const summary = snapshot.data();
          return {
            found: true,
            fecha: summary.fecha,
            asesor: summary.asesor,
            createdAtMillis: summary.createdAtServer?.toMillis?.() ?? null
          };
        })
        .finally(() => {
          if (state.latestAttendanceLookup.promise === promise) {
            state.latestAttendanceLookup = { key: '', promise: null };
          }
        });
      state.latestAttendanceLookup = { key, promise };
      return promise;
    }

    async function showAdvisorLatestAttendance(studentCode, matches, requestId) {
      const alertBox = document.getElementById('repeat-alert-box');
      const latestAttendance = document.getElementById('repeat-alert-history');
      const hasLocalMatches = matches.length > 0;
      if (!hasLocalMatches) {
        alertBox.classList.remove('hidden');
        document.getElementById('repeat-alert-badge').textContent = 'Última atención';
        document.getElementById('repeat-alert-summary').textContent =
          'Consultando la última atención registrada globalmente.';
        latestAttendance.replaceChildren();
      }

      const card = document.createElement('div');
      card.className = 'bg-white/90 border border-orange-200 rounded-xl p-3 shadow-sm space-y-3';

      const header = document.createElement('div');
      header.className = 'flex flex-wrap items-center justify-between gap-2';
      const heading = document.createElement('span');
      heading.className = 'inline-flex items-center gap-1.5 text-[10px] font-extrabold uppercase tracking-wider text-orange-800';
      heading.textContent = 'Última atención global';
      const dateTime = document.createElement('span');
      dateTime.className = 'text-[11px] font-semibold text-slate-500';
      dateTime.textContent = 'Consultando...';
      header.append(heading, dateTime);
      card.appendChild(header);

      const advisor = document.createElement('div');
      const advisorLabel = document.createElement('p');
      advisorLabel.className = 'text-[10px] font-bold uppercase tracking-wider text-slate-400';
      advisorLabel.textContent = 'Asesor';
      const advisorName = document.createElement('p');
      advisorName.className = 'text-sm font-extrabold text-[#C8102E]';
      advisorName.textContent = 'Consultando...';
      advisor.append(advisorLabel, advisorName);
      card.appendChild(advisor);
      latestAttendance.appendChild(card);

      try {
        const latest = await getStudentAttendanceSummary(studentCode);
        if (
          requestId !== state.latestAttendanceRequestId
          || document.getElementById('entry-registro').value.trim() !== studentCode
          || state.currentUser?.role !== 'Asesor'
        ) return;

        if (!latest.found) {
          card.remove();
          if (!hasLocalMatches) {
            alertBox.classList.add('hidden');
          } else {
            showToast(
              'No se encontró una atención previa para este estudiante.',
              'info'
            );
          }
          return;
        }

        dateTime.textContent =
          `${formatDateDisplay(latest.fecha)} · ${formatTimeDisplay(latest.createdAtMillis)}`;
        advisorName.textContent = latest.asesor;
        if (!hasLocalMatches) {
          document.getElementById('repeat-alert-summary').textContent =
            `Última atención registrada para el estudiante ${studentCode}.`;
        }
      } catch (error) {
        console.error('Error al consultar el resumen de última atención:', error);
        if (requestId !== state.latestAttendanceRequestId) return;
        if (!hasLocalMatches) {
          document.getElementById('repeat-alert-summary').textContent =
            'No se pudo consultar la última atención. Inténtelo nuevamente.';
        }
        dateTime.textContent = 'No disponible';
        advisorName.textContent = 'No disponible';
        showToast(
          `No se pudo consultar la última atención (${error?.code || 'error-desconocido'}).`,
          'error'
        );
      }
    }

    // Detección instantánea usando los registros disponibles para la sesión actual.
    window.handleStudentCodeInput = function (rawVal) {
      const requestId = ++state.latestAttendanceRequestId;
      const input = document.getElementById('entry-registro');
      const clean = normalizeStudentCodeInput(rawVal);
      input.value = clean;
      document.getElementById('registro-char-counter').textContent = `${clean.length}/6`;

      const alertBox = document.getElementById('repeat-alert-box');
      if (!isValidStudentCode(clean)) {
        alertBox.classList.add('hidden');
        return;
      }

      const selectedDate = document.getElementById('entry-fecha').value || getTodayISO();
      const targetMonth = selectedDate.slice(0, 7);
      const matches = state.records.filter(
        (r) => r.registro === clean && r.mesAnio === targetMonth
      );

      if (matches.length === 0) {
        if (state.currentUser?.role === 'Asesor') {
          showAdvisorLatestAttendance(clean, matches, requestId);
          return;
        }
        alertBox.classList.add('hidden');
        return;
      }

      const todayMatches = matches.filter((r) => r.fecha === selectedDate);
      const latestMatch = [...matches].sort((a, b) => {
        const aTimestamp = a.createdAt || Date.parse(`${a.fecha}T00:00:00Z`) || 0;
        const bTimestamp = b.createdAt || Date.parse(`${b.fecha}T00:00:00Z`) || 0;
        return bTimestamp - aTimestamp;
      })[0];
      alertBox.classList.remove('hidden');

      document.getElementById('repeat-alert-badge').textContent =
        todayMatches.length > 0
          ? `¡Atendido HOY (${todayMatches.length}x)!`
          : `${matches.length} atención(es) este mes`;

      document.getElementById('repeat-alert-summary').textContent =
        `El estudiante ${clean} ya registra ${matches.length} atención(es) en ${formatMonthName(targetMonth)}. Al guardarse se marcará en naranja como re-consulta.`;

      const attendanceTime = formatTimeDisplay(latestMatch.createdAt);
      const attendanceProcedures = PROCEDURES
        .filter((procedure) => latestMatch.tramites?.[procedure.key])
        .map((procedure) => procedure.label);
      const latestAttendance = document.getElementById('repeat-alert-history');
      latestAttendance.replaceChildren();

      const card = document.createElement('div');
      card.className = 'bg-white/90 border border-orange-200 rounded-xl p-3 shadow-sm space-y-3';

      const header = document.createElement('div');
      header.className = 'flex flex-wrap items-center justify-between gap-2';

      const heading = document.createElement('span');
      heading.className = 'inline-flex items-center gap-1.5 text-[10px] font-extrabold uppercase tracking-wider text-orange-800';
      heading.textContent = 'Última atención';

      const dateTime = document.createElement('span');
      dateTime.className = 'text-[11px] font-semibold text-slate-500';
      dateTime.textContent = `${formatDateDisplay(latestMatch.fecha)} · ${attendanceTime}`;

      header.append(heading, dateTime);
      card.appendChild(header);

      const advisor = document.createElement('div');
      const advisorLabel = document.createElement('p');
      advisorLabel.className = 'text-[10px] font-bold uppercase tracking-wider text-slate-400';
      advisorLabel.textContent = 'Asesor';
      const advisorName = document.createElement('p');
      advisorName.className = 'text-sm font-extrabold text-[#C8102E]';
      advisorName.textContent = latestMatch.asesor;
      advisor.append(advisorLabel, advisorName);
      card.appendChild(advisor);

      const proceduresSection = document.createElement('div');
      const proceduresLabel = document.createElement('p');
      proceduresLabel.className = 'text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-1.5';
      proceduresLabel.textContent = 'Atención recibida';
      const procedureTags = document.createElement('div');
      procedureTags.className = 'flex flex-wrap gap-1.5';
      const procedureNames = attendanceProcedures.length > 0 ? attendanceProcedures : ['Sin trámite'];
      procedureNames.forEach((name) => {
        const tag = document.createElement('span');
        tag.className = 'px-2 py-1 rounded-lg bg-[#FDF2F4] border border-red-100 text-[10px] font-bold text-[#C8102E]';
        tag.textContent = name;
        procedureTags.appendChild(tag);
      });
      proceduresSection.append(proceduresLabel, procedureTags);
      card.appendChild(proceduresSection);

      const observation = document.createElement('div');
      observation.className = 'rounded-lg bg-orange-50/80 px-2.5 py-2';
      const observationLabel = document.createElement('p');
      observationLabel.className = 'text-[10px] font-bold uppercase tracking-wider text-orange-700';
      observationLabel.textContent = 'Observación';
      const observationText = document.createElement('p');
      observationText.className = 'text-xs text-slate-600 mt-0.5 break-words';
      observationText.textContent = latestMatch.observaciones || '-';
      observation.append(observationLabel, observationText);
      card.appendChild(observation);

      latestAttendance.appendChild(card);
      if (state.currentUser?.role === 'Asesor') {
        showAdvisorLatestAttendance(clean, matches, requestId);
      }
    };

    window.handleEditStudentCodeInput = function (rawVal) {
      document.getElementById('edit-registro').value = normalizeStudentCodeInput(rawVal);
    };

    window.handleSaveRecord = async function (e) {
      e.preventDefault();
      const registroInput = document.getElementById('entry-registro');
      const registro = normalizeStudentCodeInput(registroInput.value);
      registroInput.value = registro;
      const fecha = state.currentUser?.role === 'Administrador'
        ? document.getElementById('entry-fecha').value || getTodayISO()
        : getTodayISO();
      const observaciones = document.getElementById('entry-observaciones').value.trim().toUpperCase();

      if (!isValidStudentCode(registro)) {
        showToast('Ingrese un código de 4 a 6 dígitos o cinco dígitos seguidos de A.', 'error');
        registroInput.focus();
        return;
      }

      if (state.selectedProcedures.size === 0) {
        showToast('Marque al menos un trámite solicitado con (X).', 'error');
        return;
      }

      const tramitesObj = {};
      PROCEDURES.forEach((p) => {
        tramitesObj[p.key] = state.selectedProcedures.has(p.key);
      });

      const mesAnio = fecha.slice(0, 7);
      const advisorName = state.currentUser?.name || 'Oscar Vargas';
      const newPayload = {
        fecha,
        mesAnio,
        registro,
        tramites: tramitesObj,
        observaciones: observaciones || 'INFORMACION',
        asesor: advisorName,
        asesorUid: state.currentUser.uid,
        canalAtencion: state.channelAtencion,
        createdAt: Date.now(),
        createdAtServer: serverTimestamp()
      };

      const saveButton = document.querySelector('#fast-entry-form button[type="submit"]');
      if (!state.firebase.db || !state.firebase.auth?.currentUser) {
        showToast('Inicie sesión en Firebase antes de guardar. El formulario se conserva.', 'error');
        return;
      }

      saveButton.disabled = true;
      try {
        const colRef = collection(state.firebase.db, 'artifacts', appId, 'public', 'data', 'atenciones');
        const recordRef = doc(colRef);
        let writePromise;
        if (state.currentUser?.role === 'Asesor') {
          const batch = writeBatch(state.firebase.db);
          batch.set(recordRef, newPayload);
          batch.set(
            getAdvisorLatestAttendanceSummaryRef(registro),
            createAdvisorAttendanceSummaryData(registro, {
              ...newPayload,
              id: recordRef.id
            })
          );
          writePromise = batch.commit();
        } else {
          writePromise = setDoc(recordRef, newPayload);
        }
        const queuedRecord = { id: recordRef.id, ...newPayload, pendingSync: true };

        state.records = [queuedRecord, ...state.records.filter((record) => record.id !== recordRef.id)];
        state.firebase.pendingWrites = state.records.filter((record) => record.pendingSync).length;
        updateFirebaseStatusUI(state.firebase.connected, state.firebase.app?.options.projectId, true);
        refreshAllViews();

        const chosenNames = PROCEDURES.filter((p) => tramitesObj[p.key]).map((p) => p.short).join(', ');
        document.getElementById('quick-stat-last-code').textContent = registro;
        document.getElementById('quick-stat-last-proc').textContent = chosenNames;
        resetFastForm();
        saveButton.disabled = false;
        showToast('Registro guardado en este dispositivo; sincronización con Firebase pendiente.', 'info');

        writePromise.then(() => {
          state.records = state.records.map((record) =>
            record.id === recordRef.id ? { ...record, pendingSync: false } : record
          );
          state.firebase.pendingWrites = state.records.filter((record) => record.pendingSync).length;
          updateFirebaseStatusUI(state.firebase.connected, state.firebase.app?.options.projectId, true);
          refreshAllViews();
          const summaryUpdate = state.currentUser?.role === 'Administrador'
            ? updateLatestAttendanceSummary(recordRef, newPayload)
            : Promise.resolve();
          summaryUpdate
            .then(() => {
              showToast(`Registro ${registro} sincronizado con Firebase.`, 'success');
            })
            .catch((summaryError) => {
              console.error('Error actualizando resumen de última atención:', summaryError);
              showToast(
                `El registro se guardó, pero no se actualizó la última atención (${summaryError?.code || 'error-desconocido'}).`,
                'error'
              );
            });
        }).catch((err) => {
          console.error('Error guardando en Firestore:', err);
          state.records = state.records.filter((record) => record.id !== recordRef.id);
          state.firebase.pendingWrites = state.records.filter((record) => record.pendingSync).length;
          updateFirebaseStatusUI(state.firebase.connected, state.firebase.app?.options.projectId, true);
          refreshAllViews();
          const errorCode = err?.code || 'error-desconocido';
          const message = err?.code === 'permission-denied'
            ? 'Firestore rechazó el registro por permisos. Revise las reglas de escritura.'
            : `Firebase no pudo guardar el registro (${errorCode}). Revise la conexión e inténtelo de nuevo.`;
          if (
            !document.getElementById('entry-registro').value.trim() &&
            state.selectedProcedures.size === 0
          ) {
            document.getElementById('entry-registro').value = registro;
            document.getElementById('entry-observaciones').value = observaciones;
            state.selectedProcedures = new Set(
              PROCEDURES.filter((procedure) => tramitesObj[procedure.key]).map((procedure) => procedure.key)
            );
            renderProcedureButtons();
          }
          showToast(message, 'error');
        });
      } catch (err) {
        console.error('Error preparando el registro para Firestore:', err);
        saveButton.disabled = false;
        showToast(`No se pudo poner el registro en cola (${err?.code || 'error-desconocido'}).`, 'error');
      }
    };

    window.resetFastForm = function () {
      document.getElementById('entry-registro').value = '';
      document.getElementById('registro-char-counter').textContent = '0/6';
      document.getElementById('entry-observaciones').value = '';
      document.getElementById('repeat-alert-box').classList.add('hidden');
      state.selectedProcedures.clear();
      renderProcedureButtons();
      document.getElementById('entry-registro').focus();
    };

    // Atajos de teclado rápidos cuando no se está escribiendo en un input de texto
    document.addEventListener('keydown', (e) => {
      if (!state.currentUser || state.activeTab !== 'registro') return;
      const activeTag = document.activeElement?.tagName;
      if (activeTag === 'INPUT' || activeTag === 'TEXTAREA' || activeTag === 'SELECT') return;

      const key = e.key.toUpperCase();
      const matchedProc = PROCEDURES.find((p) => p.hotkey === key);
      if (matchedProc) {
        e.preventDefault();
        toggleProcedure(matchedProc.key);
      }
    });

    // =========================================================================
    // 7. CÁLCULO DE RE-CONSULTAS Y RENDERIZADO DE VISTA RÁPIDA DE HOY
    // =========================================================================
    function getStudentMonthCounts(recordsList) {
      // Mapa: "YYYY-MM|REGISTRO" -> conteo
      const counts = new Map();
      recordsList.forEach((r) => {
        const key = `${r.mesAnio}|${r.registro}`;
        counts.set(key, (counts.get(key) || 0) + 1);
      });
      return counts;
    }

    function escapeHtml(value) {
      return String(value ?? '').replace(/[&<>"']/g, (character) => ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;'
      })[character]);
    }

    function bindRecordActionHandlers(container) {
      container.onclick = (event) => {
        const button = event.target?.closest?.('button[data-record-action]');
        if (!button) return;

        if (button.dataset.recordAction === 'edit') {
          window.openEditModal(button.dataset.recordId);
        } else if (button.dataset.recordAction === 'delete') {
          window.requestDeleteRecord(button.dataset.recordId, button.dataset.recordCode);
        }
      };
    }

    function renderTodayFeed() {
      const todayISO = document.getElementById('entry-fecha').value || getTodayISO();
      const currentMonth = todayISO.slice(0, 7);
      const advisorName = state.currentUser?.name || '';
      const isSupervisor = state.currentUser?.role === 'Supervisor' || state.currentUser?.role === 'Administrador';
      const canDeleteRecords = state.currentUser?.role === 'Administrador';

      const monthCounts = getStudentMonthCounts(state.records);
      const todayRecords = state.records.filter(
        (r) => r.fecha === todayISO && (isSupervisor || r.asesor === advisorName)
      );

      document.getElementById('today-feed-subtitle').textContent =
        `${formatDateDisplay(todayISO)} • ${todayRecords.length} atenciones registradas`;

      const container = document.getElementById('today-feed-container');
      if (todayRecords.length === 0) {
        container.innerHTML = `
          <div class="p-10 text-center">
            <p class="text-sm font-bold text-slate-600">Aún no hay registros en esta fecha</p>
            <p class="text-xs text-slate-400 mt-1">Complete el formulario de la izquierda para comenzar.</p>
          </div>
        `;
        return;
      }

      container.innerHTML = todayRecords.slice(0, 30).map((rec) => {
        const isRepeat = (monthCounts.get(`${rec.mesAnio}|${rec.registro}`) || 0) > 1;
        const activeLabels = PROCEDURES.filter((p) => rec.tramites?.[p.key]).map((p) => p.short);
        return `
          <div class="p-3.5 hover:bg-slate-50 transition flex items-center justify-between gap-3">
            <div class="min-w-0 flex-1">
              <div class="flex items-center gap-2">
                <span class="px-2.5 py-0.5 rounded-lg text-xs font-black tracking-wide ${
                  isRepeat
                    ? 'bg-orange-100 text-orange-900 border border-orange-300'
                    : 'bg-slate-100 text-slate-800'
                }">${escapeHtml(rec.registro)}</span>
                ${isRepeat ? '<span class="text-[10px] font-extrabold text-orange-600 uppercase">Re-consulta</span>' : ''}
                ${rec.pendingSync ? '<span class="text-[10px] font-extrabold text-amber-700 uppercase">Pendiente de sincronizar</span>' : ''}
                <span class="px-1.5 py-0.5 rounded-md text-[9px] font-extrabold uppercase ${
                  rec.canalAtencion === 'wati'
                    ? 'bg-emerald-100 text-emerald-800'
                    : 'bg-slate-100 text-slate-600'
                }">${rec.canalAtencion === 'wati' ? 'WATI' : 'Plataforma'}</span>
                <span class="text-[11px] text-slate-400 truncate">• ${escapeHtml(rec.observaciones || '-')}</span>
              </div>
              <div class="flex flex-wrap gap-1 mt-1.5">
                ${activeLabels.map((lbl) => `
                  <span class="px-2 py-0.5 rounded-md text-[10px] font-extrabold bg-[#FDF2F4] text-[#C8102E] border border-red-100">X ${lbl}</span>
                `).join('')}
              </div>
            </div>
            <div class="flex items-center gap-1 shrink-0">
              <button type="button" data-record-action="edit" data-record-id="${escapeHtml(rec.id)}" title="Editar" class="p-1.5 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-100">
                <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M16.862 4.487l1.687-1.688a1.875 1.875 0 112.652 2.652L10.582 16.07a4.5 4.5 0 01-1.897 1.13L6 18l.8-2.685a4.5 4.5 0 011.13-1.897l8.932-8.931z"/></svg>
              </button>
              ${canDeleteRecords ? `<button type="button" data-record-action="delete" data-record-id="${escapeHtml(rec.id)}" data-record-code="${escapeHtml(rec.registro)}" title="Eliminar" class="p-1.5 rounded-lg text-slate-400 hover:text-[#C8102E] hover:bg-red-50">
                <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M14.74 9l-.346 9m-4.788 0L9.26 9m9.968-3.21c.342.052.682.107 1.022.166m-1.022-.165L18.16 19.673a2.25 2.25 0 01-2.244 2.077H8.084a2.25 2.25 0 01-2.244-2.077L4.772 5.79m14.456 0a48.108 48.108 0 00-3.478-.397m-12 .562c.34-.059.68-.114 1.022-.165m0 0a48.11 48.11 0 013.478-.397m7.5 0v-.916c0-1.18-.91-2.164-2.09-2.201a51.964 51.964 0 00-3.32 0c-1.18.037-2.09 1.022-2.09 2.201v.916m7.5 0a48.667 48.667 0 00-7.5 0"/></svg>
              </button>` : ''}
            </div>
          </div>
        `;
      }).join('');
      bindRecordActionHandlers(container);

      // Actualizar tarjetas superiores de Registro Rápido
      const advisorMonthRecords = state.records.filter(
        (r) => r.mesAnio === currentMonth && (isSupervisor || r.asesor === advisorName)
      );
      const goalTarget = isSupervisor ? 2500 : 1250;
      const pct = Math.min(100, Math.round((advisorMonthRecords.length / goalTarget) * 100));

      let repeatRowsCount = 0;
      advisorMonthRecords.forEach((r) => {
        if ((monthCounts.get(`${r.mesAnio}|${r.registro}`) || 0) > 1) repeatRowsCount++;
      });

      document.getElementById('quick-stat-today').textContent = todayRecords.length.toLocaleString('es-BO');
      document.getElementById('quick-stat-month').textContent = advisorMonthRecords.length.toLocaleString('es-BO');
      document.getElementById('quick-stat-month-pct').textContent = `${pct}%`;
      document.getElementById('quick-stat-month-bar').style.width = `${pct}%`;
      document.getElementById('quick-stat-repeats').textContent = repeatRowsCount.toLocaleString('es-BO');
    }

    // =========================================================================
    // 8. HISTORIAL PAGINADO (SOPORTE PARA 1,500+ REGISTROS)
    // =========================================================================
    function getDateRangeFilter(prefix) {
      const from = document.getElementById(`${prefix}-date-from`)?.value || '';
      const to = document.getElementById(`${prefix}-date-to`)?.value || '';
      const valid = !from || !to || from <= to;
      const error = document.getElementById(`${prefix}-range-error`);
      if (error) error.classList.toggle('hidden', valid);
      return { from, to, hasRange: Boolean(from || to), valid };
    }

    function isDateInRange(date, range) {
      if (!range.valid) return false;
      if (range.from && date < range.from) return false;
      if (range.to && date > range.to) return false;
      return true;
    }

    function getFilteredHistoryRecords() {
      const search = (document.getElementById('hist-search')?.value || '').trim().toLowerCase();
      const monthFilter = document.getElementById('hist-filter-month')?.value || 'ALL';
      const procFilter = document.getElementById('hist-filter-proc')?.value || 'ALL';
      const advisorFilter = document.getElementById('hist-filter-advisor')?.value || 'ALL';
      const onlyRepeats = document.getElementById('hist-filter-repeats')?.checked || false;
      const dateRange = getDateRangeFilter('hist');

      const monthCounts = getStudentMonthCounts(state.records);

      return state.records.filter((r) => {
        if (!isDateInRange(r.fecha || '', dateRange)) return false;
        if (!dateRange.hasRange && monthFilter !== 'ALL' && r.mesAnio !== monthFilter) return false;
        if (advisorFilter !== 'ALL' && r.asesor !== advisorFilter) return false;
        if (procFilter !== 'ALL' && !r.tramites?.[procFilter]) return false;
        if (onlyRepeats && (monthCounts.get(`${r.mesAnio}|${r.registro}`) || 0) <= 1) return false;
        if (search) {
          const matchCode = r.registro.toLowerCase().includes(search);
          const matchObs = (r.observaciones || '').toLowerCase().includes(search);
          if (!matchCode && !matchObs) return false;
        }
        return true;
      });
    }

    window.resetHistoryPageAndRender = function () {
      state.historyPage = 1;
      renderHistoryTable();
    };

    window.clearHistoryFilters = function () {
      document.getElementById('hist-search').value = '';
      document.getElementById('hist-date-from').value = '';
      document.getElementById('hist-date-to').value = '';
      document.getElementById('hist-filter-month').value = 'ALL';
      document.getElementById('hist-filter-proc').value = 'ALL';
      document.getElementById('hist-filter-advisor').value =
        state.currentUser?.role === 'Asesor' ? state.currentUser.name : 'ALL';
      document.getElementById('hist-filter-repeats').checked = false;
      state.historyPage = 1;
      renderHistoryTable();
    };

    window.changeHistoryPage = function (dir) {
      const filtered = getFilteredHistoryRecords();
      const pageSize = parseInt(document.getElementById('pagination-size').value, 10) || 50;
      const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));

      if (dir === 'first') state.historyPage = 1;
      else if (dir === 'prev' && state.historyPage > 1) state.historyPage--;
      else if (dir === 'next' && state.historyPage < totalPages) state.historyPage++;
      else if (dir === 'last') state.historyPage = totalPages;

      renderHistoryTable();
    };

    function renderHistoryTable() {
      const filtered = getFilteredHistoryRecords();
      const pageSize = parseInt(document.getElementById('pagination-size').value, 10) || 50;
      const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
      if (state.historyPage > totalPages) state.historyPage = totalPages;

      const startIdx = (state.historyPage - 1) * pageSize;
      const pageSlice = filtered.slice(startIdx, startIdx + pageSize);
      const monthCounts = getStudentMonthCounts(state.records);
      const canDeleteRecords = state.currentUser?.role === 'Administrador';

      const tbody = document.getElementById('history-table-body');
      if (pageSlice.length === 0) {
        tbody.innerHTML = `
          <tr>
            <td colspan="17" class="py-12 text-center text-slate-500">
              No se encontraron registros que coincidan con los filtros seleccionados.
            </td>
          </tr>
        `;
      } else {
        tbody.innerHTML = pageSlice.map((rec) => {
          const repeatTotal = monthCounts.get(`${rec.mesAnio}|${rec.registro}`) || 0;
          const isRepeat = repeatTotal > 1;

          const procCells = PROCEDURES.map((p) => {
            const marked = Boolean(rec.tramites?.[p.key]);
            return `
              <td class="py-2.5 px-2 text-center border-r border-slate-100">
                ${marked
                  ? '<span class="inline-flex items-center justify-center w-5 h-5 rounded bg-[#FDF2F4] text-[#C8102E] font-black text-[11px]">X</span>'
                  : '<span class="text-slate-300">-</span>'}
              </td>
            `;
          }).join('');

          return `
            <tr class="hover:bg-slate-50/90 transition">
              <td class="py-2.5 px-3 font-semibold text-slate-700 border-r border-slate-100">
                <span class="block">${escapeHtml(formatDateDisplay(rec.fecha))}</span>
                <span class="block text-[10px] font-medium text-slate-400 mt-0.5">${formatTimeDisplay(rec.createdAt)}</span>
              </td>
              <td class="py-2.5 px-3 border-r border-slate-100 ${isRepeat ? 'bg-orange-100/90' : ''}">
                <span class="font-black ${isRepeat ? 'text-orange-900' : 'text-slate-900'}">${escapeHtml(rec.registro)}</span>
                ${isRepeat ? `<span class="ml-1 text-[10px] font-bold text-orange-700" title="Atendido ${repeatTotal} veces en este mes">(${repeatTotal}x)</span>` : ''}
              </td>
              ${procCells}
              <td class="py-2.5 px-3 font-medium text-slate-700 border-r border-slate-100 max-w-[220px] truncate" title="${escapeHtml(rec.observaciones || '')}">${escapeHtml(rec.observaciones || '-')}</td>
              <td class="py-2.5 px-3 border-r border-slate-100">
                <span class="px-2 py-1 rounded-md text-[10px] font-bold ${
                  rec.canalAtencion === 'wati'
                    ? 'bg-emerald-100 text-emerald-800'
                    : 'bg-slate-100 text-slate-600'
                }">${rec.canalAtencion === 'wati' ? 'WATI' : 'Plataforma'}</span>
              </td>
              <td class="py-2.5 px-3 text-slate-600 font-semibold border-r border-slate-100">${escapeHtml(rec.asesor)}</td>
              <td class="py-2.5 px-3 text-center">
                <div class="inline-flex items-center gap-1">
                  <button type="button" data-record-action="edit" data-record-id="${escapeHtml(rec.id)}" class="px-2 py-1 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-[11px]">Editar</button>
                  ${canDeleteRecords ? `<button type="button" data-record-action="delete" data-record-id="${escapeHtml(rec.id)}" data-record-code="${escapeHtml(rec.registro)}" class="px-2 py-1 rounded-lg bg-red-50 hover:bg-red-100 text-[#C8102E] font-bold text-[11px]">Eliminar</button>` : ''}
                </div>
              </td>
            </tr>
          `;
        }).join('');
      }
      bindRecordActionHandlers(tbody);

      const showingFrom = filtered.length === 0 ? 0 : startIdx + 1;
      const showingTo = Math.min(filtered.length, startIdx + pageSize);
      document.getElementById('pagination-info').textContent =
        `Mostrando ${showingFrom.toLocaleString('es-BO')} - ${showingTo.toLocaleString('es-BO')} de ${filtered.length.toLocaleString('es-BO')} registros`;
      document.getElementById('pagination-page-indicator').textContent =
        `Página ${state.historyPage} de ${totalPages}`;

      document.getElementById('btn-page-first').disabled = state.historyPage <= 1;
      document.getElementById('btn-page-prev').disabled = state.historyPage <= 1;
      document.getElementById('btn-page-next').disabled = state.historyPage >= totalPages;
      document.getElementById('btn-page-last').disabled = state.historyPage >= totalPages;
    }

    // =========================================================================
    // 9. EDICIÓN Y ELIMINACIÓN DE REGISTROS (SIN confirm() NATIVO)
    // =========================================================================
    window.openEditModal = function (recordId) {
      const rec = state.records.find((r) => r.id === recordId);
      if (!rec) return;

      document.getElementById('edit-record-id').value = rec.id;
      document.getElementById('edit-registro').value = rec.registro;
      const canEditStudentCode = state.currentUser?.role === 'Administrador';
      setRecordCodeEditable(document.getElementById('edit-registro'), canEditStudentCode);
      const editDateInput = document.getElementById('edit-fecha');
      editDateInput.value = rec.fecha;
      const canEditAttendanceDate = state.currentUser?.role === 'Administrador';
      setDateFieldEditable(editDateInput, canEditAttendanceDate);
      document.getElementById('edit-fecha-help').textContent = canEditAttendanceDate
        ? 'El administrador puede modificar esta fecha'
        : 'Los asesores no pueden modificar la fecha del registro';
      document.getElementById('edit-observaciones').value = rec.observaciones || '';

      state.editSelectedProcedures.clear();
      PROCEDURES.forEach((p) => {
        if (rec.tramites?.[p.key]) state.editSelectedProcedures.add(p.key);
      });

      renderEditProceduresGrid();
      document.getElementById('modal-edit').classList.remove('hidden');
    };

    function renderEditProceduresGrid() {
      const container = document.getElementById('edit-procedures-grid');
      container.innerHTML = PROCEDURES.map((p) => {
        const active = state.editSelectedProcedures.has(p.key);
        return `
          <button
            type="button"
            onclick="toggleEditProcedure('${p.key}')"
            class="p-2.5 rounded-xl border text-left text-xs font-bold flex items-center gap-2 transition ${
              active
                ? 'bg-[#C8102E] border-[#C8102E] text-white'
                : 'bg-slate-50 border-slate-200 text-slate-700 hover:border-[#C8102E]'
            }"
          >
            <span class="w-4 h-4 rounded flex items-center justify-center text-[10px] font-black ${
              active ? 'bg-white text-[#C8102E]' : 'bg-white border border-slate-300'
            }">${active ? 'X' : ''}</span>
            <span class="truncate">${p.label}</span>
          </button>
        `;
      }).join('');
    }

    window.toggleEditProcedure = function (key) {
      if (state.editSelectedProcedures.has(key)) state.editSelectedProcedures.delete(key);
      else state.editSelectedProcedures.add(key);
      renderEditProceduresGrid();
    };

    window.handleUpdateRecord = async function (e) {
      e.preventDefault();
      const id = document.getElementById('edit-record-id').value;
      const existingRecord = state.records.find((record) => record.id === id);
      if (!existingRecord) {
        showToast('No se encontró el registro que intenta editar.', 'error');
        return;
      }
      const canEditStudentCode = state.currentUser?.role === 'Administrador';
      const registro = canEditStudentCode
        ? normalizeStudentCodeInput(document.getElementById('edit-registro').value)
        : existingRecord.registro;
      const fecha = state.currentUser?.role === 'Administrador'
        ? document.getElementById('edit-fecha').value
        : existingRecord.fecha;
      const observaciones = document.getElementById('edit-observaciones').value.trim().toUpperCase();

      if (!isValidStudentCode(registro)) {
        showToast('Ingrese un código de 4 a 6 dígitos o cinco dígitos seguidos de A.', 'error');
        return;
      }
      if (canEditStudentCode) {
        document.getElementById('edit-registro').value = registro;
      }
      if (!fecha) {
        showToast('La fecha de atención es obligatoria.', 'error');
        return;
      }
      if (state.editSelectedProcedures.size === 0) {
        showToast('Seleccione al menos un trámite.', 'error');
        return;
      }

      const tramites = {};
      PROCEDURES.forEach((p) => {
        tramites[p.key] = state.editSelectedProcedures.has(p.key);
      });
      const mesAnio = fecha.slice(0, 7);

      if (state.firebase.db && state.firebase.auth?.currentUser && !id.startsWith('local_')) {
        try {
          const docRef = doc(state.firebase.db, 'artifacts', appId, 'public', 'data', 'atenciones', id);
          const updates = canEditStudentCode
            ? { registro, fecha, mesAnio, tramites, observaciones }
            : { tramites, observaciones };
          await updateDoc(docRef, updates);
        } catch (err) {
          console.error('Error actualizando en Firestore:', err);
          showToast('Error al actualizar en Firestore.', 'error');
          return;
        }
      } else {
        const idx = state.records.findIndex((r) => r.id === id);
        if (idx !== -1) {
          state.records[idx] = { ...state.records[idx], registro, fecha, mesAnio, tramites, observaciones };
          refreshAllViews();
        }
      }

      closeModal('modal-edit');
      showToast('Registro actualizado correctamente.', 'success');
    };

    window.requestDeleteRecord = function (id, registroCode) {
      if (state.currentUser?.role !== 'Administrador') {
        showToast('Solo un administrador puede eliminar registros.', 'error');
        return;
      }

      openConfirmModal(
        'Eliminar Registro',
        `¿Está seguro que desea eliminar la atención del estudiante ${registroCode}?`,
        async () => {
          if (state.firebase.db && state.firebase.auth?.currentUser && !id.startsWith('local_')) {
            try {
              const docRef = doc(state.firebase.db, 'artifacts', appId, 'public', 'data', 'atenciones', id);
              await deleteDoc(docRef);
            } catch (err) {
              console.error('Error eliminando en Firestore:', err);
              showToast('Error al eliminar en Firestore.', 'error');
              return;
            }
          } else {
            state.records = state.records.filter((r) => r.id !== id);
            refreshAllViews();
          }
          showToast(`Registro ${registroCode} eliminado.`, 'info');
        }
      );
    };

    // =========================================================================
    // 10. DASHBOARD DE REPORTE MENSUAL, KPIs Y GRÁFICOS (Chart.js)
    // =========================================================================
    function getFilteredReportRecords(advisorFilter) {
      const selectedMonth = document.getElementById('report-month-select').value || getTodayISO().slice(0, 7);
      const selectedAdvisor = advisorFilter
        || document.getElementById('report-advisor-select').value
        || 'ALL';
      const dateRange = getDateRangeFilter('report');

      return state.records.filter((record) => {
        if (!isDateInRange(record.fecha || '', dateRange)) return false;
        if (!dateRange.hasRange && record.mesAnio !== selectedMonth) return false;
        return selectedAdvisor === 'ALL' || record.asesor === selectedAdvisor;
      });
    }

    function renderMonthlyReport() {
      updateReportAdvisorsHeading();
      const selectedMonth = document.getElementById('report-month-select').value || getTodayISO().slice(0, 7);
      const selectedAdvisor = document.getElementById('report-advisor-select').value || 'ALL';
      const dateRange = getDateRangeFilter('report');

      const periodAllRecords = getFilteredReportRecords('ALL');
      const filteredRecords = getFilteredReportRecords(selectedAdvisor);

      const totalAttentions = filteredRecords.length;
      const watiAttentions = filteredRecords.filter((record) => record.canalAtencion === 'wati').length;
      const platformAttentions = totalAttentions - watiAttentions;
      const platformPct = totalAttentions > 0
        ? Math.round((platformAttentions / totalAttentions) * 100)
        : 0;
      const watiPct = totalAttentions > 0
        ? Math.round((watiAttentions / totalAttentions) * 100)
        : 0;

      document.getElementById('kpi-channel-platform-count').textContent =
        platformAttentions.toLocaleString('es-BO');
      document.getElementById('kpi-channel-platform-pct').textContent = `${platformPct}%`;
      document.getElementById('kpi-channel-platform-bar').style.width = `${platformPct}%`;
      document.getElementById('kpi-channel-wati-count').textContent =
        watiAttentions.toLocaleString('es-BO');
      document.getElementById('kpi-channel-wati-pct').textContent = `${watiPct}%`;
      document.getElementById('kpi-channel-wati-bar').style.width = `${watiPct}%`;

      // Meta mensual (1,000 - 1,500 por asesor; escalable si es vista global de varios asesores)
      const uniqueAdvisorsInPeriod = new Set(periodAllRecords.map((r) => r.asesor));
      const advisorMultiplier = selectedAdvisor === 'ALL' ? Math.max(1, uniqueAdvisorsInPeriod.size) : 1;
      const minGoal = 1000 * advisorMultiplier;
      const maxGoal = 1500 * advisorMultiplier;
      const midGoal = 1250 * advisorMultiplier;
      const goalPct = dateRange.hasRange
        ? null
        : Math.min(100, Math.round((totalAttentions / midGoal) * 100));

      document.getElementById('kpi-goal-badge').textContent =
        dateRange.hasRange
          ? 'Rango seleccionado'
          : `Meta: ${minGoal.toLocaleString('es-BO')}-${maxGoal.toLocaleString('es-BO')}`;
      document.getElementById('kpi-total-attentions').textContent = totalAttentions.toLocaleString('es-BO');
      document.getElementById('kpi-goal-pct').textContent = goalPct === null ? '—' : `${goalPct}%`;
      document.getElementById('kpi-goal-bar').style.width = `${goalPct || 0}%`;
      document.getElementById('kpi-total-label').textContent =
        dateRange.hasRange ? 'Total Atenciones del Periodo' : 'Total Atenciones Mes';
      document.getElementById('kpi-goal-caption').textContent =
        dateRange.hasRange ? 'Meta mensual no aplicable' : 'Cumplimiento meta';

      // Estudiantes únicos vs Re-consultas
      const studentFreq = new Map();
      filteredRecords.forEach((r) => {
        studentFreq.set(r.registro, (studentFreq.get(r.registro) || 0) + 1);
      });
      const uniqueStudentsCount = studentFreq.size;
      const repeatAttentionsCount = Math.max(0, totalAttentions - uniqueStudentsCount);
      const repeatPct = totalAttentions > 0 ? ((repeatAttentionsCount / totalAttentions) * 100).toFixed(1) : '0.0';

      document.getElementById('kpi-unique-students').textContent = uniqueStudentsCount.toLocaleString('es-BO');
      document.getElementById('kpi-repeat-count').textContent = `+${repeatAttentionsCount.toLocaleString('es-BO')} re-consultas`;
      document.getElementById('kpi-repeat-rate').textContent =
        `${repeatPct}% de atenciones recurrentes en el ${dateRange.hasRange ? 'periodo' : 'mes'}`;

      // Conteo de los 11 trámites (Qué más pidieron los estudiantes)
      const procCounts = PROCEDURES.map((p) => {
        let count = 0;
        filteredRecords.forEach((r) => {
          if (r.tramites?.[p.key]) count++;
        });
        return { ...p, count };
      }).sort((a, b) => b.count - a.count);

      const totalProcMarks = procCounts.reduce((acc, item) => acc + item.count, 0);
      document.getElementById('total-marks-badge').textContent = `${totalProcMarks.toLocaleString('es-BO')} trámites (X)`;

      const topProc = procCounts[0];
      if (topProc && topProc.count > 0) {
        const sharePct = totalProcMarks > 0 ? ((topProc.count / totalProcMarks) * 100).toFixed(1) : '0.0';
        document.getElementById('kpi-top-procedure').textContent = topProc.label;
        document.getElementById('kpi-top-procedure-stats').textContent =
          `${topProc.count.toLocaleString('es-BO')} solicitudes (${sharePct}% de trámites)`;
      } else {
        document.getElementById('kpi-top-procedure').textContent = 'Sin datos';
        document.getElementById('kpi-top-procedure-stats').textContent = '0 solicitudes';
      }

      // Evolución diaria y Promedio diario
      const dailyMap = new Map();
      filteredRecords.forEach((r) => {
        dailyMap.set(r.fecha, (dailyMap.get(r.fecha) || 0) + 1);
      });
      const activeDays = dailyMap.size;
      const dailyAvg = activeDays > 0 ? (totalAttentions / activeDays).toFixed(1) : '0';

      let peakDate = '-';
      let peakCount = 0;
      dailyMap.forEach((cnt, dt) => {
        if (cnt > peakCount) {
          peakCount = cnt;
          peakDate = formatDateDisplay(dt);
        }
      });

      document.getElementById('kpi-daily-avg').textContent = dailyAvg;
      document.getElementById('kpi-peak-day').textContent =
        peakCount > 0 ? `Día pico: ${peakDate} (${peakCount} est.)` : 'Día pico: -';

      // Tabla de desglose de trámites
      document.getElementById('report-procedures-table').innerHTML = procCounts.map((item, idx) => {
        const pct = totalProcMarks > 0 ? ((item.count / totalProcMarks) * 100).toFixed(1) : '0.0';
        return `
          <tr class="hover:bg-slate-50">
            <td class="py-2 px-2.5 font-bold text-slate-800 flex items-center gap-2">
              <span class="w-4 h-4 rounded text-[10px] font-black flex items-center justify-center ${
                idx < 3 ? 'bg-[#C8102E] text-white' : 'bg-slate-100 text-slate-600'
              }">${idx + 1}</span>
              <span>${item.label}</span>
            </td>
            <td class="py-2 px-2 text-right font-extrabold text-slate-900">${item.count.toLocaleString('es-BO')}</td>
            <td class="py-2 px-2.5 text-right font-bold text-[#C8102E]">${pct}%</td>
          </tr>
        `;
      }).join('');

      // Actualizar metadatos para impresión ejecutiva
      const periodLabel = dateRange.hasRange
        ? `Del ${dateRange.from ? formatDateDisplay(dateRange.from) : 'inicio'} al ${dateRange.to ? formatDateDisplay(dateRange.to) : 'presente'}`
        : formatMonthName(selectedMonth);
      document.getElementById('report-title').textContent =
        dateRange.hasRange ? 'Reporte de Atención SAE' : 'Reporte Mensual de Atención SAE';
      document.getElementById('print-meta-period').textContent = `Periodo: ${periodLabel}`;
      document.getElementById('print-meta-advisor').textContent =
        `Asesor: ${selectedAdvisor === 'ALL' ? 'Consolidado General SAE' : selectedAdvisor}`;
      document.getElementById('print-meta-date').textContent = `Generado: ${formatDateDisplay(getTodayISO())}`;
      // Renderizar los 3 gráficos interactivos
      renderProceduresChart(procCounts);
      renderDailyTrendChart(selectedMonth, dailyMap, dateRange.hasRange);
      renderAdvisorsChart(periodAllRecords);
    }

    window.renderMonthlyReport = renderMonthlyReport;

    window.clearReportFilters = function () {
      document.getElementById('report-month-select').value = getTodayISO().slice(0, 7);
      document.getElementById('report-date-from').value = '';
      document.getElementById('report-date-to').value = '';
      document.getElementById('report-advisor-select').value =
        state.currentUser?.role === 'Asesor' ? state.currentUser.name : 'ALL';
      renderMonthlyReport();
    };

    function updateReportAdvisorsHeading() {
      const isAdvisor = state.currentUser?.role === 'Asesor';
      document.getElementById('report-advisors-title').textContent = isAdvisor
        ? 'Atención global del asesor'
        : 'Comparativa de Atenciones por Asesor';
      document.getElementById('report-advisors-description').textContent = isAdvisor
        ? 'Total de atenciones registradas por usted en el periodo seleccionado'
        : 'Distribución de carga operativa entre asesores SAE';
    }

    function renderProceduresChart(procCounts) {
      const ctx = document.getElementById('chart-procedures').getContext('2d');
      if (state.charts.procedures) state.charts.procedures.destroy();

      state.charts.procedures = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: procCounts.map((p) => p.label),
          datasets: [{
            label: 'Solicitudes (X)',
            data: procCounts.map((p) => p.count),
            backgroundColor: procCounts.map((_, i) => i === 0 ? '#C8102E' : i < 3 ? '#E11D48' : '#63666A'),
            borderRadius: 6,
            barPercentage: 0.72
          }]
        },
        options: {
          indexAxis: 'y',
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              backgroundColor: '#1E293B',
              padding: 10,
              cornerRadius: 10
            }
          },
          scales: {
            x: {
              grid: { color: '#F1F5F9' },
              ticks: { font: { size: 11, weight: '600' }, color: '#64748B' }
            },
            y: {
              grid: { display: false },
              ticks: { font: { size: 11, weight: '700' }, color: '#1E293B' }
            }
          }
        }
      });
    }

    function renderDailyTrendChart(selectedMonth, dailyMap, hasDateRange) {
      const ctx = document.getElementById('chart-daily-trend').getContext('2d');
      if (state.charts.daily) state.charts.daily.destroy();

      const labels = [];
      const values = [];

      if (hasDateRange) {
        Array.from(dailyMap.entries())
          .sort(([dateA], [dateB]) => dateA.localeCompare(dateB))
          .forEach(([date, count]) => {
            labels.push(formatDateDisplay(date));
            values.push(count);
          });
      } else {
        const [yStr, mStr] = selectedMonth.split('-');
        const daysInMonth = new Date(parseInt(yStr, 10), parseInt(mStr, 10), 0).getDate();
        for (let d = 1; d <= daysInMonth; d++) {
          const iso = `${selectedMonth}-${String(d).padStart(2, '0')}`;
          labels.push(`${String(d).padStart(2, '0')}/${mStr}`);
          values.push(dailyMap.get(iso) || 0);
        }
      }

      state.charts.daily = new Chart(ctx, {
        type: 'line',
        data: {
          labels,
          datasets: [{
            label: 'Estudiantes Atendidos',
            data: values,
            borderColor: '#C8102E',
            backgroundColor: 'rgba(200, 16, 46, 0.12)',
            borderWidth: 2.5,
            fill: true,
            tension: 0.32,
            pointRadius: 3,
            pointHoverRadius: 6,
            pointBackgroundColor: '#C8102E'
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              backgroundColor: '#1E293B',
              padding: 10,
              cornerRadius: 10
            }
          },
          scales: {
            x: {
              grid: { display: false },
              ticks: { maxRotation: 0, autoSkip: true, maxTicksLimit: 12, font: { size: 10 }, color: '#64748B' }
            },
            y: {
              beginAtZero: true,
              grid: { color: '#F1F5F9' },
              ticks: { font: { size: 11 }, color: '#64748B' }
            }
          }
        }
      });
    }

    function renderAdvisorsChart(periodRecords) {
      const ctx = document.getElementById('chart-advisors').getContext('2d');
      if (state.charts.advisors) state.charts.advisors.destroy();

      const advMap = new Map();
      periodRecords.forEach((r) => {
        advMap.set(r.asesor, (advMap.get(r.asesor) || 0) + 1);
      });

      const sortedAdvisors = Array.from(advMap.entries()).sort((a, b) => b[1] - a[1]);

      state.charts.advisors = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: sortedAdvisors.map((a) => a[0]),
          datasets: [{
            label: 'Atenciones del Periodo',
            data: sortedAdvisors.map((a) => a[1]),
            backgroundColor: ['#C8102E', '#63666A', '#94A3B8', '#E11D48'],
            borderRadius: 8,
            barPercentage: 0.55
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              backgroundColor: '#1E293B',
              padding: 10,
              cornerRadius: 10
            }
          },
          scales: {
            x: {
              grid: { display: false },
              ticks: { font: { size: 11, weight: '700' }, color: '#1E293B' }
            },
            y: {
              beginAtZero: true,
              grid: { color: '#F1F5F9' },
              ticks: { font: { size: 11 }, color: '#64748B' }
            }
          }
        }
      });
    }

    // =========================================================================
    // 11. EXPORTACIÓN EXCEL UTEPSA (.XLS Y .CSV) E IMPRESIÓN EJECUTIVA
    // =========================================================================
    function getRecordsForExport(source) {
      if (source === 'historial') {
        return getFilteredHistoryRecords();
      }
      return getFilteredReportRecords();
    }

    window.exportToCSV = function (source) {
      const rows = getRecordsForExport(source);
      if (rows.length === 0) {
        showToast('No hay registros para exportar en el filtro actual.', 'error');
        return;
      }

      const headers = [
        'FECHA', 'REGISTRO', 'CANAL',
        ...PROCEDURES.map((p) => p.label),
        'OBSERVACIONES', 'ASESOR'
      ];

      const csvLines = [headers.join(';')];
      rows.forEach((r) => {
        const line = [
          formatDateDisplay(r.fecha),
          r.registro,
          r.canalAtencion === 'wati' ? 'WATI' : 'PLATAFORMA',
          ...PROCEDURES.map((p) => (r.tramites?.[p.key] ? 'X' : '')),
          `"${(r.observaciones || '').replace(/"/g, '""')}"`,
          `"${(r.asesor || '').replace(/"/g, '""')}"`
        ];
        csvLines.push(line.join(';'));
      });

      const blob = new Blob(['\uFEFF' + csvLines.join('\n')], { type: 'text/csv;charset=utf-8;' });
      triggerDownload(blob, `Reporte_Atencion_UTEPSA_SAE_${getTodayISO()}.csv`);
      showToast(`CSV exportado (${rows.length} registros).`, 'success');
    };

    window.exportToExcelXLS = function (source) {
      const rows = getRecordsForExport(source);
      if (rows.length === 0) {
        showToast('No hay registros para exportar en el filtro actual.', 'error');
        return;
      }

      const monthCounts = getStudentMonthCounts(state.records);
      const advisorTitle = state.currentUser?.name || 'SAE UTEPSA';

      const tableRowsHtml = rows.map((r) => {
        const isRepeat = (monthCounts.get(`${r.mesAnio}|${r.registro}`) || 0) > 1;
        const regBg = isRepeat ? 'background-color:#F8CBAD;font-weight:bold;' : '';
        const procCols = PROCEDURES.map((p) =>
          `<td style="text-align:center;border:1px solid #D9D9D9;">${r.tramites?.[p.key] ? 'X' : ''}</td>`
        ).join('');

        return `
          <tr>
            <td style="text-align:center;border:1px solid #D9D9D9;">${escapeHtml(formatDateDisplay(r.fecha))}</td>
            <td style="text-align:center;border:1px solid #D9D9D9;${regBg}">${escapeHtml(r.registro)}</td>
            <td style="text-align:center;border:1px solid #D9D9D9;">${r.canalAtencion === 'wati' ? 'WATI' : 'PLATAFORMA'}</td>
            ${procCols}
            <td style="border:1px solid #D9D9D9;">${escapeHtml(r.observaciones || '')}</td>
            <td style="border:1px solid #D9D9D9;">${escapeHtml(r.asesor || '')}</td>
          </tr>
        `;
      }).join('');

      const htmlExcel = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <head><meta charset="UTF-8"></head>
        <body>
          <table>
            <tr>
              <td colspan="4" style="font-size:16px;font-weight:bold;color:#C8102E;">UTEPSA - UNIVERSIDAD TECNOLÓGICA PRIVADA DE SANTA CRUZ</td>
              <td colspan="7" style="text-align:center;font-weight:bold;color:#C8102E;">"UN GRAN PODER CONLLEVA UNA GRAN RESPONSABILIDAD"</td>
              <td colspan="5" style="text-align:right;font-weight:bold;color:#C8102E;">${escapeHtml(advisorTitle.toUpperCase())}</td>
            </tr>
            <tr style="background-color:#C8102E;color:#FFFFFF;font-weight:bold;text-align:center;">
              <th style="border:1px solid #9E0B22;padding:6px;">FECHA</th>
              <th style="border:1px solid #9E0B22;padding:6px;">REGISTRO</th>
              <th style="border:1px solid #9E0B22;padding:6px;">CANAL</th>
              ${PROCEDURES.map((p) => `<th style="border:1px solid #9E0B22;padding:6px;">${p.label}</th>`).join('')}
              <th style="border:1px solid #9E0B22;padding:6px;">OBSERVACIONES</th>
              <th style="border:1px solid #9E0B22;padding:6px;">ASESOR</th>
            </tr>
            ${tableRowsHtml}
          </table>
        <\/body>
        <\/html>
      `;

      const blob = new Blob(['\uFEFF' + htmlExcel], { type: 'application/vnd.ms-excel;charset=utf-8;' });
      triggerDownload(blob, `Planilla_Atencion_UTEPSA_SAE_${getTodayISO()}.xls`);
      showToast(`Planilla Excel UTEPSA exportada (${rows.length} registros).`, 'success');
    };

    window.printExecutiveReport = function () {
      window.print();
    };

    function triggerDownload(blob, filename) {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }

    // =========================================================================
    // 12. SINCRONIZACIÓN DE SELECTORES Y PRECARGA DE MUESTRA
    // =========================================================================
    function refreshAllViews() {
      document.getElementById('nav-count-badge').textContent = state.records.length.toLocaleString('es-BO');
      populateDynamicFilters();
      renderTodayFeed();
      if (state.activeTab === 'historial') renderHistoryTable();
      if (state.activeTab === 'reportes') renderMonthlyReport();
    }

    function populateDynamicFilters() {
      const currentMonth = getTodayISO().slice(0, 7);
      const monthsSet = new Set([currentMonth]);
      const advisorsSet = new Set();
      if (state.currentUser?.name) advisorsSet.add(state.currentUser.name);

      state.records.forEach((r) => {
        if (r.mesAnio) monthsSet.add(r.mesAnio);
        if (r.asesor) advisorsSet.add(r.asesor);
      });

      const sortedMonths = Array.from(monthsSet).sort((a, b) => b.localeCompare(a));
      const sortedAdvisors = Array.from(advisorsSet).sort();

      // Filtros de Historial
      const histMonth = document.getElementById('hist-filter-month');
      const prevHistMonth = histMonth.value;
      histMonth.innerHTML = '<option value="ALL">Todos los meses</option>' +
        sortedMonths.map((m) => `<option value="${escapeHtml(m)}">${escapeHtml(formatMonthName(m))}</option>`).join('');
      if (sortedMonths.includes(prevHistMonth)) histMonth.value = prevHistMonth;

      const histProc = document.getElementById('hist-filter-proc');
      if (histProc.options.length <= 1) {
        PROCEDURES.forEach((p) => {
          const opt = document.createElement('option');
          opt.value = p.key;
          opt.textContent = p.label;
          histProc.appendChild(opt);
        });
      }

      const histAdv = document.getElementById('hist-filter-advisor');
      const prevHistAdv = histAdv.value;
      const currentAdvisor = state.currentUser?.role === 'Asesor' ? state.currentUser.name : '';
      histAdv.disabled = Boolean(currentAdvisor);
      histAdv.replaceChildren();
      if (currentAdvisor) {
        const option = document.createElement('option');
        option.value = currentAdvisor;
        option.textContent = currentAdvisor;
        histAdv.appendChild(option);
      } else {
        const allOption = document.createElement('option');
        allOption.value = 'ALL';
        allOption.textContent = 'Todos los asesores';
        histAdv.appendChild(allOption);
        sortedAdvisors.forEach((advisor) => {
          const option = document.createElement('option');
          option.value = advisor;
          option.textContent = advisor;
          histAdv.appendChild(option);
        });
        histAdv.value = sortedAdvisors.includes(prevHistAdv) ? prevHistAdv : 'ALL';
      }

      // Filtros de Reportes Mensuales
      const repMonth = document.getElementById('report-month-select');
      const prevRepMonth = repMonth.value;
      repMonth.innerHTML = sortedMonths.map((m) => `<option value="${escapeHtml(m)}">${escapeHtml(formatMonthName(m))}</option>`).join('');
      repMonth.value = sortedMonths.includes(prevRepMonth) ? prevRepMonth : currentMonth;

      const repAdv = document.getElementById('report-advisor-select');
      const prevRepAdv = repAdv.value;
      repAdv.disabled = Boolean(currentAdvisor);
      repAdv.replaceChildren();
      if (currentAdvisor) {
        const option = document.createElement('option');
        option.value = currentAdvisor;
        option.textContent = currentAdvisor;
        repAdv.appendChild(option);
      } else {
        const allOption = document.createElement('option');
        allOption.value = 'ALL';
        allOption.textContent = 'Todos los Asesores (Global)';
        repAdv.appendChild(allOption);
        sortedAdvisors.forEach((advisor) => {
          const option = document.createElement('option');
          option.value = advisor;
          option.textContent = advisor;
          repAdv.appendChild(option);
        });
        repAdv.value = sortedAdvisors.includes(prevRepAdv) ? prevRepAdv : 'ALL';
      }
    }

    // =========================================================================
    // 13. MODALES, UTILIDADES DE FECHA Y NOTIFICACIONES TOAST
    // =========================================================================
    window.closeModal = function (id) {
      document.getElementById(id).classList.add('hidden');
    };

    function openConfirmModal(title, message, onAccept) {
      document.getElementById('confirm-title').textContent = title;
      document.getElementById('confirm-message').textContent = message;
      const btn = document.getElementById('confirm-accept-btn');
      const newBtn = btn.cloneNode(true);
      btn.parentNode.replaceChild(newBtn, btn);
      newBtn.addEventListener('click', () => {
        closeModal('modal-confirm');
        onAccept();
      });
      document.getElementById('modal-confirm').classList.remove('hidden');
    }

    function showToast(message, type = 'info') {
      const container = document.getElementById('toast-container');
      const colors = {
        success: 'bg-emerald-600 text-white',
        error: 'bg-[#C8102E] text-white',
        info: 'bg-slate-800 text-white'
      };
      const el = document.createElement('div');
      el.className = `pointer-events-auto px-4 py-3 rounded-2xl shadow-lg text-xs font-bold flex items-center justify-between gap-3 animate-fade-in ${colors[type] || colors.info}`;
      const messageEl = document.createElement('span');
      messageEl.textContent = message;
      const dismissButton = document.createElement('button');
      dismissButton.type = 'button';
      dismissButton.className = 'opacity-80 hover:opacity-100';
      dismissButton.textContent = '✕';
      dismissButton.addEventListener('click', () => el.remove());
      el.append(messageEl, dismissButton);
      container.appendChild(el);
      setTimeout(() => {
        if (el.parentNode) el.remove();
      }, 3800);
    }

    function getTodayISO() {
      const now = new Date();
      const y = now.getFullYear();
      const m = String(now.getMonth() + 1).padStart(2, '0');
      const d = String(now.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }

    function formatDateDisplay(isoDate) {
      if (!isoDate || !isoDate.includes('-')) return isoDate || '-';
      const [y, m, d] = isoDate.split('-');
      return `${d}/${m}/${y}`;
    }

    function setDateFieldEditable(input, editable) {
      input.disabled = !editable;
      input.setAttribute('aria-disabled', String(!editable));
      input.classList.toggle('bg-slate-100', !editable);
      input.classList.toggle('border-slate-200', !editable);
      input.classList.toggle('text-slate-500', !editable);
      input.classList.toggle('cursor-not-allowed', !editable);
      input.classList.toggle('bg-slate-50', editable);
      input.classList.toggle('border-slate-300', editable);
      input.classList.toggle('text-slate-800', editable);
      input.classList.toggle('cursor-pointer', editable);
    }

    function setRecordCodeEditable(input, editable) {
      input.readOnly = !editable;
      input.setAttribute('aria-readonly', String(!editable));
      input.classList.toggle('bg-slate-100', !editable);
      input.classList.toggle('border-slate-200', !editable);
      input.classList.toggle('text-slate-500', !editable);
      input.classList.toggle('cursor-not-allowed', !editable);
      input.classList.toggle('bg-white', editable);
      input.classList.toggle('border-slate-300', editable);
      input.classList.toggle('text-slate-900', editable);
      input.classList.toggle('cursor-text', editable);
    }

    function formatTimeDisplay(timestamp) {
      if (!Number.isFinite(timestamp) || timestamp <= 0) return '--:--:--';
      return new Intl.DateTimeFormat('es-BO', {
        timeZone: 'America/La_Paz',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hourCycle: 'h23'
      }).format(new Date(timestamp));
    }

    function formatMonthName(ym) {
      if (!ym || !ym.includes('-')) return ym;
      const [y, m] = ym.split('-');
      const idx = parseInt(m, 10) - 1;
      return `${MONTH_NAMES_ES[idx] || m} ${y}`;
    }

    // =========================================================================
    // 14. ARRANQUE INICIAL DE LA APLICACIÓN
    // =========================================================================
    window.addEventListener('DOMContentLoaded', async () => {
      renderProcedureButtons();
      renderQuickObsTags();

      try {
        const envConfig = typeof __firebase_config !== 'undefined' && __firebase_config
          ? (typeof __firebase_config === 'string' ? JSON.parse(__firebase_config) : __firebase_config)
          : FIREBASE_CONFIG;
        await initFirebaseFromConfig(envConfig);
      } catch (err) {
        console.error('No se pudo cargar la configuración de Firebase:', err);
        showToast('No se pudo cargar la configuración de Firebase.', 'error');
      }
    });
