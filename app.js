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
      { key: 'pensum', label: 'PENSUM', short: 'PENSUM', hotkey: '1', specialOnly: true },
      { key: 'programacion', label: 'PROGRAMACION', short: 'PROGR.', hotkey: '2', specialOnly: false },
      { key: 'confirmacion', label: 'CONFIRMACION', short: 'CONFIR.', hotkey: '3', specialOnly: false },
      { key: 'plan_pago', label: 'PLAN DE PAGO', short: 'PLAN PAGO', hotkey: '4', specialOnly: false },
     { key: 'recargo', label: 'RECARGO', short: 'RECARGO', hotkey: '5', specialOnly: true },
      { key: 'congelamiento', label: 'RETIRO / CONGELAMIENTO', short: 'RET./CONG.', hotkey: '6', specialOnly: false },
      { key: 'pro_forma', label: 'PRO-FORMA', short: 'PRO-FORMA', hotkey: '7', specialOnly: true },
      { key: 'abonos', label: 'ABONOS', short: 'ABONOS', hotkey: '8', specialOnly: true },
      { key: 'informacion', label: 'INFORMACION', short: 'INFORM.', hotkey: '9', specialOnly: false },
      { key: 'reingresante', label: 'REINGRESANTE', short: 'REINGR.', hotkey: '0', specialOnly: false },
      { key: 'ex_grado', label: 'EX. DE GRADO / TESIS', short: 'EX. GRADO', hotkey: 'Q', specialOnly: false },
      { key: 'cambio_materia', label: 'CAMBIO DE MATERIA', short: 'CAMB. MAT.', hotkey: 'W', specialOnly: false },
    ];

    function isAdvisorRole(role = state.currentUser?.role) {
      return role === 'Asesor' || role === 'Asesor Especial';
    }

    function canUserAccessProcedure(proc, role = state.currentUser?.role) {
      if (!proc?.specialOnly) return true;
      return role === 'Administrador' || role === 'Asesor Especial';
    }

    function getAllowedProceduresForCurrentUser() {
      return PROCEDURES.filter((proc) => canUserAccessProcedure(proc));
    }

    const QUICK_MODULE_TAGS = ['MOD 0', 'MOD 1', 'MOD 2', 'MOD 3', 'MOD 4', 'MOD 5'];

    const QUICK_OBS_TAGS = [
      { label: 'OFERTA', value: 'OFERTA', type: 'tramite' },
      { label: 'IMP. BOLETA', value: 'IMPRESION DE BOLETA', type: 'tramite' },
      { label: 'VINC. CONTRATO', value: 'VINCULACION DE CONTRATO', type: 'tramite' },
      { label: 'CARTA PLAZO', value: 'CARTA FUERA DE PLAZO', type: 'tramite' },
      { label: 'DEUDAS', value: 'DEUDAS', type: 'tramite' },
      { label: 'HIST. NOTAS', value: 'HISTORIAL DE NOTAS', type: 'tramite' },
      { label: 'DEF. EXTERNA', value: 'DEFENSA EXTERNA', type: 'tramite' },
      { label: 'INFORMACION', value: 'INFORMACION', type: 'tramite' },
      { label: 'DERIV. REGISTRO', value: 'DERIVACION A REGISTRO', type: 'derivacion' },
      { label: 'DERIV. CAJA', value: 'DERIVACION A CAJA', type: 'derivacion' }
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
      currentUser: null,
      channelAtencion: 'plataforma',
      editChannelAtencion: 'plataforma',
      reportChannelFilter: 'ALL',
      activeTab: 'registro',
      records: [], // Todos los registros sincronizados (Firestore o sesión activa)
     selectedProcedures: new Set(),
      editSelectedProcedures: new Set(),
      congelamientoFee: '',
      editCongelamientoFee: '',
      historyPage: 1,
     charts: {
        procedures: null,
        daily: null,
        advisors: null,
        hourly: null
      },
      reportSections: {
        kpis: true,
        tramites: true,
        diario: true,
        asesores: true,
        congelamientos: true
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
      latestAttendanceRequestId: 0,
      loadedMonths: new Set(),
      loadingMonths: new Set(),
      studentSummaryCache: new Map(),
      studentLookupTimer: null,
      activeDateISO: getTodayISO(),
      dateWatcherInterval: null
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

    function mapFirestoreRecord(docSnap) {
      const data = docSnap.data();
      return {
        id: docSnap.id,
        fecha: data.fecha || getTodayISO(),
        mesAnio: data.mesAnio || (data.fecha ? data.fecha.slice(0, 7) : getTodayISO().slice(0, 7)),
        registro: String(data.registro || '').trim(),
        tramites: data.tramites || {},
        congelamientoDetalle: data.congelamientoDetalle || null,
        observaciones: data.observaciones || '',
        asesor: data.asesor || 'Oscar Vargas',
        asesorUid: data.asesorUid || '',
        canalAtencion: data.canalAtencion === 'wati' ? 'wati' : 'plataforma',
        createdAt: Number.isFinite(data.createdAt) ? data.createdAt : 0,
        createdAtServer: data.createdAtServer || null,
        pendingSync: docSnap.metadata.hasPendingWrites
      };
    }

    function startFirestoreRealtimeSync() {
      if (!state.firebase.db || !state.firebase.auth?.currentUser) return;

      if (state.firebase.unsubscribeSnapshot) {
        state.firebase.unsubscribeSnapshot();
      }

      const currentMonth = getTodayISO().slice(0, 7);
      state.loadedMonths.add(currentMonth);

      const recordsRef = collection(state.firebase.db, 'artifacts', appId, 'public', 'data', 'atenciones');
      const colRef = isAdvisorRole()
        ? query(
            recordsRef,
            where('asesorUid', '==', state.currentUser.uid),
            where('mesAnio', '==', currentMonth)
          )
        : query(recordsRef, where('mesAnio', '==', currentMonth));

      state.firebase.unsubscribeSnapshot = onSnapshot(
        colRef,
        { includeMetadataChanges: true },
        (snapshot) => {
          const loadedCurrentMonth = [];
          snapshot.forEach((docSnap) => {
            loadedCurrentMonth.push(mapFirestoreRecord(docSnap));
          });

          const historicalRecords = state.records.filter((r) => r.mesAnio !== currentMonth);
          const merged = [...loadedCurrentMonth, ...historicalRecords];
          merged.sort((a, b) => {
            if (b.fecha !== a.fecha) return b.fecha.localeCompare(a.fecha);
            return (b.createdAt || 0) - (a.createdAt || 0);
          });

          state.records = merged;
          state.firebase.pendingWrites = merged.filter((record) => record.pendingSync).length;
          state.firebase.connected = !snapshot.metadata.fromCache;

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

    function getRecentMonthsList(count = 12) {
      const now = new Date();
      let y = now.getFullYear();
      let m = now.getMonth() + 1;
      const list = [];
      for (let i = 0; i < count; i++) {
        list.push(`${y}-${String(m).padStart(2, '0')}`);
        m--;
        if (m < 1) {
          m = 12;
          y--;
        }
      }
      return list;
    }

    function getMonthsBetweenDates(fromISO, toISO) {
      const endYM = (toISO || getTodayISO()).slice(0, 7);
      const startYM = (fromISO || getRecentMonthsList(12).slice(-1)[0] + '-01').slice(0, 7);
      if (startYM > endYM) return [];
      const result = [];
      let [y, m] = startYM.split('-').map(Number);
      const [endY, endM] = endYM.split('-').map(Number);
      while (y < endY || (y === endY && m <= endM)) {
        result.push(`${y}-${String(m).padStart(2, '0')}`);
        m++;
        if (m > 12) {
          m = 1;
          y++;
        }
      }
      return result;
    }

    async function ensureMonthsLoaded(monthsToLoad) {
      if (!state.firebase.db || !state.firebase.auth?.currentUser) return false;
      const missing = monthsToLoad.filter(
        (ym) => ym && !state.loadedMonths.has(ym) && !state.loadingMonths.has(ym)
      );
      if (missing.length === 0) return false;

      missing.forEach((ym) => state.loadingMonths.add(ym));
      try {
        const recordsRef = collection(state.firebase.db, 'artifacts', appId, 'public', 'data', 'atenciones');
        const fetchedAll = [];
        for (let i = 0; i < missing.length; i += 10) {
          const chunk = missing.slice(i, i + 10);
          const q = isAdvisorRole()
            ? query(recordsRef, where('asesorUid', '==', state.currentUser.uid), where('mesAnio', 'in', chunk))
            : query(recordsRef, where('mesAnio', 'in', chunk));
          const snap = await getDocs(q);
          snap.forEach((docSnap) => fetchedAll.push(mapFirestoreRecord(docSnap)));
        }

        const loadedSet = new Set(missing);
        const kept = state.records.filter((r) => !loadedSet.has(r.mesAnio));
        const merged = [...kept, ...fetchedAll];
        merged.sort((a, b) => {
          if (b.fecha !== a.fecha) return b.fecha.localeCompare(a.fecha);
          return (b.createdAt || 0) - (a.createdAt || 0);
        });

        state.records = merged;
        missing.forEach((ym) => state.loadedMonths.add(ym));
        document.getElementById('nav-count-badge').textContent = state.records.length.toLocaleString('es-BO');
        populateDynamicFilters();
        return true;
      } catch (err) {
        console.error('Error cargando meses anteriores:', err);
        showToast('No se pudieron cargar los registros de meses anteriores.', 'error');
        return false;
      } finally {
        missing.forEach((ym) => state.loadingMonths.delete(ym));
      }
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

    function createAdvisorAttendanceSummaryData(studentCode, record) {
      return {
        registro: studentCode,
        asesor: record.asesor,
        fecha: record.fecha,
        tramites: record.tramites || {},
        observaciones: record.observaciones || 'INFORMACION',
        createdAtServer: record.createdAtServer || toAttendanceTimestamp(record),
        recordId: record.id
      };
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
    const isSpecialAdvisor = profile.rol === 'Asesor Especial' && profileEmail === authEmail;
    const isAdvisor = profile.rol === 'Asesor' && profileEmail === authEmail;
    const role = isAdmin
      ? 'Administrador'
      : isSpecialAdvisor
        ? 'Asesor Especial'
        : isAdvisor
          ? 'Asesor'
          : '';
    const name = String(profile.nombre || (isInitialAdmin ? 'Oscar Vargas' : '')).trim();
    const phone = String(profile.telefono || '').trim();

    if (!role || !name || profileEmail !== authEmail || (!isAdmin && !phone)) {
      if (!isInitialAdmin) {
        await signOut(state.firebase.auth);
      }
      showToast(
        isInitialAdmin
          ? 'El perfil del administrador no coincide con su cuenta de Firebase.'
          : 'No existe un perfil válido. Registre nombre, teléfono, correo y rol "Asesor", "Asesor Especial" o "Administrador" en Firestore.',
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
  const passwordInput = document.getElementById('auth-password');
  passwordInput.type = 'password';
  const password = passwordInput.value;

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
      startDateRolloverWatcher();

      renderProcedureButtons();

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
          state.reportChannelFilter = 'ALL';
          updateChannelModeUI();
         state.records = [];
         state.loadedMonths.clear();
          state.loadingMonths.clear();
          state.studentSummaryCache.clear();
          clearTimeout(state.studentLookupTimer);
          if (state.dateWatcherInterval) {
            clearInterval(state.dateWatcherInterval);
            state.dateWatcherInterval = null;
          }
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
        const paginationSelect = document.getElementById('pagination-size');
        if (paginationSelect) paginationSelect.value = '10';
        state.historyPage = 1;
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
        : 'inline-flex items-center gap-2 px-2.5 py-2 rounded-xl text-[10px] sm:text-xs font-bold border bg-slate-50 text-[#C8102E] border-slate-200 hover:bg-slate-100 transition';
      indicator.className = isWati
        ? 'w-2.5 h-2.5 rounded-full bg-emerald-500'
        : 'w-2.5 h-2.5 rounded-full bg-[#C8102E]';
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
    const visibleProcedures = getAllowedProceduresForCurrentUser();
    container.innerHTML = visibleProcedures.map((proc) => {
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

    function updateCongelamientoSubpanelUI() {
      const subpanel = document.getElementById('congelamiento-subpanel');
      if (!subpanel) return;
      const isSelected = state.selectedProcedures.has('congelamiento');
      subpanel.classList.toggle('hidden', !isSelected);
      if (!isSelected) {
        state.congelamientoFee = '';
        const check = document.getElementById('cong-carta-check');
        if (check) check.checked = false;
      }
      ['GRATIS', '300', '500'].forEach((fee) => {
        const btn = document.getElementById(`cong-fee-btn-${fee}`);
        if (!btn) return;
        const active = state.congelamientoFee === fee;
        btn.className = active
          ? 'py-2.5 px-3 rounded-xl border-2 border-[#C8102E] bg-[#C8102E] text-white text-xs font-extrabold shadow-sm transition text-center'
          : 'py-2.5 px-3 rounded-xl border-2 border-slate-200 bg-white text-slate-700 hover:border-[#C8102E] text-xs font-extrabold transition text-center';
      });
    }

    window.selectCongelamientoFee = function (fee) {
      state.congelamientoFee = fee;
      updateCongelamientoSubpanelUI();
    };

    window.toggleProcedure = function (key) {
    const proc = PROCEDURES.find((p) => p.key === key);
    if (!proc || !canUserAccessProcedure(proc)) return;
    if (state.selectedProcedures.has(key)) {
      state.selectedProcedures.delete(key);
    } else {
      state.selectedProcedures.add(key);
    }
    renderProcedureButtons();
    updateCongelamientoSubpanelUI();
  };

    function renderQuickObsTags() {
      const container = document.getElementById('quick-obs-tags');

      const modulesHtml = `
        <div class="w-full flex flex-wrap items-center gap-1.5 pb-1.5 mb-0.5 border-b border-slate-100">
          <span class="text-[10px] font-extrabold uppercase tracking-wider text-slate-400 mr-1">Módulo:</span>
          ${QUICK_MODULE_TAGS.map((mod) => {
            const num = mod.replace('MOD ', '');
            return `
              <button
                type="button"
                onclick="applyQuickObsTag('${mod}')"
                title="Insertar Modulo ${num}"
                class="w-7 h-7 rounded-lg text-xs font-extrabold bg-slate-100 hover:bg-[#C8102E] text-slate-700 hover:text-white border border-slate-200 hover:border-[#C8102E] transition flex items-center justify-center"
              >
                ${num}
              </button>
            `;
          }).join('')}
        </div>
      `;

      const tagsHtml = QUICK_OBS_TAGS.map((item) => {
        const isDerivacion = item.type === 'derivacion';
        const styleClasses = isDerivacion
          ? 'bg-amber-50/80 hover:bg-amber-100 text-amber-800 border-amber-200 hover:border-amber-300'
          : 'bg-slate-50 hover:bg-[#FDF2F4] text-slate-700 hover:text-[#C8102E] border-slate-200 hover:border-red-200';

        return `
          <button
            type="button"
            onclick="applyQuickObsTag('${item.value}')"
            title="${item.value}"
            class="px-2.5 py-1 rounded-lg text-[11px] font-bold border transition ${styleClasses}"
          >
            + ${item.label}
          </button>
        `;
      }).join('');

      container.innerHTML = `${modulesHtml}${tagsHtml}`;
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
      const cached = state.studentSummaryCache.get(studentCode);
      if (cached && (Date.now() - cached.cachedAt) < 60000) {
        return Promise.resolve(cached.data);
      }

      const cachedLookup = state.latestAttendanceLookup;
      if (cachedLookup.key === studentCode && cachedLookup.promise) return cachedLookup.promise;

      const promise = getDoc(getAdvisorLatestAttendanceSummaryRef(studentCode))
        .then((snapshot) => {
          const data = !snapshot.exists()
            ? { found: false }
            : {
                found: true,
                recordId: snapshot.data().recordId || '',
                fecha: snapshot.data().fecha,
                asesor: snapshot.data().asesor,
                tramites: snapshot.data().tramites || {},
                observaciones: snapshot.data().observaciones || '',
                createdAtMillis: snapshot.data().createdAtServer?.toMillis?.() ?? null
              };
          state.studentSummaryCache.set(studentCode, { data, cachedAt: Date.now() });
          return data;
        })
        .finally(() => {
          if (state.latestAttendanceLookup.promise === promise) {
            state.latestAttendanceLookup = { key: '', promise: null };
          }
        });
      state.latestAttendanceLookup = { key: studentCode, promise };
      return promise;
    }

    function buildDetailedAttendanceCard({ headingText, fecha, createdAtMillis, asesor, tramites, observaciones }) {
      const attendanceTime = formatTimeDisplay(createdAtMillis);
      const attendanceProcedures = PROCEDURES
        .filter((procedure) => tramites?.[procedure.key])
        .map((procedure) => procedure.label);

      const card = document.createElement('div');
      card.className = 'bg-white/90 border border-orange-200 rounded-xl p-3 shadow-sm space-y-3';

      const header = document.createElement('div');
      header.className = 'flex flex-wrap items-center justify-between gap-2';

      const heading = document.createElement('span');
      heading.className = 'inline-flex items-center gap-1.5 text-[10px] font-extrabold uppercase tracking-wider text-orange-800';
      heading.textContent = headingText;

      const dateTime = document.createElement('span');
      dateTime.className = 'text-[11px] font-semibold text-slate-500';
      dateTime.textContent = `${formatDateDisplay(fecha)} · ${attendanceTime}`;

      header.append(heading, dateTime);
      card.appendChild(header);

      const advisor = document.createElement('div');
      const advisorLabel = document.createElement('p');
      advisorLabel.className = 'text-[10px] font-bold uppercase tracking-wider text-slate-400';
      advisorLabel.textContent = 'Asesor';
      const advisorName = document.createElement('p');
      advisorName.className = 'text-sm font-extrabold text-[#C8102E]';
      advisorName.textContent = asesor || 'No disponible';
      advisor.append(advisorLabel, advisorName);
      card.appendChild(advisor);

      const proceduresSection = document.createElement('div');
      const proceduresLabel = document.createElement('p');
      proceduresLabel.className = 'text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-1.5';
      proceduresLabel.textContent = 'Atención recibida';
      const procedureTags = document.createElement('div');
      procedureTags.className = 'flex flex-wrap gap-1.5';
      const procedureNames = attendanceProcedures.length > 0 ? attendanceProcedures : ['Sin detalle registrado'];
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
      observationText.textContent = observaciones || '-';
      observation.append(observationLabel, observationText);
      card.appendChild(observation);

      return card;
    }

    function setStudentLookupSpinner(visible) {
      const spinner = document.getElementById('student-lookup-spinner');
      if (spinner) spinner.classList.toggle('hidden', !visible);
    }

    async function showAdvisorLatestAttendance(studentCode, matches, localLatestId, requestId) {
      const alertBox = document.getElementById('repeat-alert-box');
      const latestAttendance = document.getElementById('repeat-alert-history');
      const hasLocalMatches = matches.length > 0;

      const cached = state.studentSummaryCache.get(studentCode);
      const isCachedFresh = cached && (Date.now() - cached.cachedAt) < 60000;
      if (!isCachedFresh) {
        setStudentLookupSpinner(true);
      }

      try {
        const latest = await getStudentAttendanceSummary(studentCode);
        if (
          requestId !== state.latestAttendanceRequestId
          || document.getElementById('entry-registro').value.trim() !== studentCode
          || !isAdvisorRole()
        ) return;

        if (!latest.found) {
          if (!hasLocalMatches) {
            alertBox.classList.add('hidden');
          }
          return;
        }

        // Si el último registro global es el mismo que ya se mostró localmente, no duplicar tarjeta
        if (hasLocalMatches && latest.recordId && latest.recordId === localLatestId) {
          return;
        }

        const selectedDate = document.getElementById('entry-fecha').value || getTodayISO();
        const isToday = latest.fecha === selectedDate;
        const detailedCard = buildDetailedAttendanceCard({
          headingText: 'Última atención en SAE',
          fecha: latest.fecha,
          createdAtMillis: latest.createdAtMillis,
          asesor: latest.asesor,
          tramites: latest.tramites,
          observaciones: latest.observaciones
        });

        alertBox.classList.remove('hidden');
        if (!hasLocalMatches) {
          document.getElementById('repeat-alert-badge').textContent = isToday ? '¡Atendido HOY!' : 'Atención previa';
          document.getElementById('repeat-alert-summary').textContent =
            `El estudiante ${studentCode} ya fue atendido previamente en SAE. Al guardarse se marcará en naranja como re-consulta.`;
          latestAttendance.replaceChildren(detailedCard);
        } else {
          latestAttendance.replaceChildren(detailedCard);
        }
      } catch (error) {
        console.error('Error al consultar el resumen de última atención:', error);
        if (requestId !== state.latestAttendanceRequestId) return;
        if (!hasLocalMatches) {
          alertBox.classList.add('hidden');
        }
        showToast(
          `No se pudo consultar la última atención (${error?.code || 'error-desconocido'}).`,
          'error'
        );
      } finally {
        if (requestId === state.latestAttendanceRequestId) {
          setStudentLookupSpinner(false);
        }
      }
    }

    // Detección instantánea usando los registros disponibles para la sesión actual.
    window.handleStudentCodeInput = function (rawVal) {
      const requestId = ++state.latestAttendanceRequestId;
      const input = document.getElementById('entry-registro');
      const clean = normalizeStudentCodeInput(rawVal);
      input.value = clean;
      document.getElementById('registro-char-counter').textContent = `${clean.length}/6`;

      clearTimeout(state.studentLookupTimer);
      setStudentLookupSpinner(false);

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

      clearTimeout(state.studentLookupTimer);

      if (matches.length === 0) {
     if (isAdvisorRole()) {
          state.studentLookupTimer = setTimeout(() => {
            showAdvisorLatestAttendance(clean, matches, '', requestId);
          }, 300);
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

      const latestAttendance = document.getElementById('repeat-alert-history');
      latestAttendance.replaceChildren(
        buildDetailedAttendanceCard({
          headingText: 'Última atención',
          fecha: latestMatch.fecha,
          createdAtMillis: latestMatch.createdAt,
          asesor: latestMatch.asesor,
          tramites: latestMatch.tramites,
          observaciones: latestMatch.observaciones
        })
      );

     if (isAdvisorRole()) {
        state.studentLookupTimer = setTimeout(() => {
          showAdvisorLatestAttendance(clean, matches, latestMatch.id, requestId);
        }, 300);
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

      const hasCongelamiento = state.selectedProcedures.has('congelamiento');
      if (hasCongelamiento && !['GRATIS', '300', '500'].includes(state.congelamientoFee)) {
        showToast('Seleccione si el congelamiento es Retiro Gratis, 300 Bs o 500 Bs.', 'error');
        return;
      }

      const congelamientoDetalle = hasCongelamiento
        ? {
            cartaRecepcionada: Boolean(document.getElementById('cong-carta-check')?.checked),
            tipoArancel: state.congelamientoFee,
            montoBs: state.congelamientoFee === '300' ? 300 : state.congelamientoFee === '500' ? 500 : 0
          }
        : null;

      const tramitesObj = {};
    PROCEDURES.forEach((p) => {
      tramitesObj[p.key] = canUserAccessProcedure(p) && state.selectedProcedures.has(p.key);
    });

      const mesAnio = fecha.slice(0, 7);
      const advisorName = state.currentUser?.name || 'Oscar Vargas';

      // Detectar posible duplicado accidental (< 2 minutos con mismos trámites y canal)
      const nowMs = Date.now();
      const recentDuplicate = state.records.find((r) => {
        if (r.registro !== registro || r.fecha !== fecha) return false;
        if (r.asesorUid && r.asesorUid !== state.currentUser?.uid) return false;
        if ((r.canalAtencion || 'plataforma') !== state.channelAtencion) return false;
        if (!r.createdAt || (nowMs - r.createdAt) > 120000) return false;
        return PROCEDURES.every((p) => Boolean(r.tramites?.[p.key]) === Boolean(tramitesObj[p.key]));
      });

      if (recentDuplicate && !e.__confirmedDuplicate) {
        const secsAgo = Math.max(1, Math.round((nowMs - recentDuplicate.createdAt) / 1000));
        const chosenNames = PROCEDURES.filter((p) => tramitesObj[p.key]).map((p) => p.label).join(', ');
        openConfirmModal(
          'Posible Registro Duplicado',
          `Ya registraste al estudiante ${registro} hace ${secsAgo} segundo(s) con exactamente los mismos trámites (${chosenNames}). ¿Deseas guardarlo nuevamente?`,
          () => window.handleSaveRecord({ preventDefault() {}, __confirmedDuplicate: true })
        );
        return;
      }

      const newPayload = {
        fecha,
        mesAnio,
        registro,
        tramites: tramitesObj,
        congelamientoDetalle,
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
        const batch = writeBatch(state.firebase.db);
        batch.set(recordRef, newPayload);
        batch.set(
          getAdvisorLatestAttendanceSummaryRef(registro),
          createAdvisorAttendanceSummaryData(registro, {
            ...newPayload,
            id: recordRef.id
          })
        );
        const writePromise = batch.commit();
        state.studentSummaryCache.delete(registro);
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
          showToast(`Registro ${registro} sincronizado con Firebase.`, 'success');
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
      state.congelamientoFee = '';
      const congCheck = document.getElementById('cong-carta-check');
      if (congCheck) congCheck.checked = false;
      renderProcedureButtons();
      updateCongelamientoSubpanelUI();
      document.getElementById('entry-registro').focus();
    };

    // Atajos de teclado rápidos cuando no se está escribiendo en un input de texto
    document.addEventListener('keydown', (e) => {
      if (!state.currentUser || state.activeTab !== 'registro') return;
      const activeTag = document.activeElement?.tagName;
      if (activeTag === 'INPUT' || activeTag === 'TEXTAREA' || activeTag === 'SELECT') return;

      const key = e.key.toUpperCase();
    const matchedProc = getAllowedProceduresForCurrentUser().find((p) => p.hotkey === key);
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
      const isSupervisor = state.currentUser?.role === 'Supervisor';
      const isAdmin = state.currentUser?.role === 'Administrador';
      const canDeleteRecords = state.currentUser?.role === 'Administrador';

      const displayRecords = state.records.filter(
        (r) => isSupervisor || (isAdmin
          ? r.asesorUid === state.currentUser?.uid
          : r.asesor === advisorName)
      );
      const monthCounts = getStudentMonthCounts(displayRecords);
      const todayRecords = displayRecords.filter((r) => r.fecha === todayISO);

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
                    : 'bg-red-100 text-[#C8102E]'
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
      const advisorMonthRecords = displayRecords.filter((r) => r.mesAnio === currentMonth);
      const goalTarget = isSupervisor || isAdmin ? 2500 : 1250;
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
      const channelFilter = document.getElementById('hist-filter-channel')?.value || 'ALL';
      const onlyRepeats = document.getElementById('hist-filter-repeats')?.checked || false;
      const dateRange = getDateRangeFilter('hist');

      const monthCounts = getStudentMonthCounts(state.records);

      return state.records.filter((r) => {
        if (!isDateInRange(r.fecha || '', dateRange)) return false;
        if (!dateRange.hasRange && monthFilter !== 'ALL' && r.mesAnio !== monthFilter) return false;
        if (advisorFilter !== 'ALL' && r.asesor !== advisorFilter) return false;
        if (channelFilter !== 'ALL' && r.canalAtencion !== channelFilter) return false;
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

    window.resetHistoryPageAndRender = async function () {
      state.historyPage = 1;
      renderHistoryTable();

      const monthFilter = document.getElementById('hist-filter-month')?.value || 'ALL';
      const dateRange = getDateRangeFilter('hist');
      if (!dateRange.valid) return;

      const monthsNeeded = dateRange.hasRange
        ? getMonthsBetweenDates(dateRange.from, dateRange.to)
        : monthFilter === 'ALL'
          ? getRecentMonthsList(12)
          : [monthFilter];

      const loadedNew = await ensureMonthsLoaded(monthsNeeded);
      if (loadedNew) {
        renderHistoryTable();
      }
    };

   window.clearHistoryFilters = function () {
      document.getElementById('hist-search').value = '';
      document.getElementById('hist-date-from').value = '';
      document.getElementById('hist-date-to').value = '';
      document.getElementById('hist-filter-month').value = getTodayISO().slice(0, 7);
      document.getElementById('hist-filter-proc').value = 'ALL';
      document.getElementById('hist-filter-advisor').value =
        isAdvisorRole() ? state.currentUser.name : 'ALL';
      document.getElementById('hist-filter-channel').value = 'ALL';
      document.getElementById('hist-filter-repeats').checked = false;
      state.historyPage = 1;
      renderHistoryTable();
    };

    window.changeHistoryPage = function (dir) {
      const filtered = getFilteredHistoryRecords();
      const pageSize = parseInt(document.getElementById('pagination-size').value, 10) || 10;
      const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));

      if (dir === 'first') state.historyPage = 1;
      else if (dir === 'prev' && state.historyPage > 1) state.historyPage--;
      else if (dir === 'next' && state.historyPage < totalPages) state.historyPage++;
      else if (dir === 'last') state.historyPage = totalPages;

      renderHistoryTable();
    };

    function renderHistoryTable() {
      const filtered = getFilteredHistoryRecords();
      const pageSize = parseInt(document.getElementById('pagination-size').value, 10) || 10;
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
          <td colspan="18" class="py-12 text-center text-slate-500">
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
                    : 'bg-red-100 text-[#C8102E]'
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
    window.selectEditChannel = function (channel) {
      state.editChannelAtencion = channel === 'wati' ? 'wati' : 'plataforma';
      const btnPlat = document.getElementById('edit-channel-btn-plataforma');
      const btnWati = document.getElementById('edit-channel-btn-wati');
      if (!btnPlat || !btnWati) return;

      if (state.editChannelAtencion === 'wati') {
        btnWati.className = 'px-3.5 py-2.5 rounded-xl border-2 border-emerald-600 bg-emerald-50 text-emerald-800 text-xs font-extrabold flex items-center justify-center gap-2 transition shadow-sm';
        btnPlat.className = 'px-3.5 py-2.5 rounded-xl border border-slate-200 bg-slate-50 text-slate-600 hover:border-slate-300 text-xs font-extrabold flex items-center justify-center gap-2 transition';
      } else {
        btnPlat.className = 'px-3.5 py-2.5 rounded-xl border-2 border-[#C8102E] bg-[#FDF2F4] text-[#C8102E] text-xs font-extrabold flex items-center justify-center gap-2 transition shadow-sm';
        btnWati.className = 'px-3.5 py-2.5 rounded-xl border border-slate-200 bg-slate-50 text-slate-600 hover:border-slate-300 text-xs font-extrabold flex items-center justify-center gap-2 transition';
      }
    };

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
      selectEditChannel(rec.canalAtencion === 'wati' ? 'wati' : 'plataforma');

      state.editSelectedProcedures.clear();
      PROCEDURES.forEach((p) => {
        if (rec.tramites?.[p.key]) state.editSelectedProcedures.add(p.key);
      });

      state.editCongelamientoFee = rec.congelamientoDetalle?.tipoArancel || '';
      const editCongCheck = document.getElementById('edit-cong-carta-check');
      if (editCongCheck) editCongCheck.checked = Boolean(rec.congelamientoDetalle?.cartaRecepcionada);

      renderEditProceduresGrid();
      updateEditCongelamientoSubpanelUI();
      document.getElementById('modal-edit').classList.remove('hidden');
    };

    function updateEditCongelamientoSubpanelUI() {
      const subpanel = document.getElementById('edit-congelamiento-subpanel');
      if (!subpanel) return;
      const isSelected = state.editSelectedProcedures.has('congelamiento');
      subpanel.classList.toggle('hidden', !isSelected);
      ['GRATIS', '300', '500'].forEach((fee) => {
        const btn = document.getElementById(`edit-cong-fee-btn-${fee}`);
        if (!btn) return;
        const active = state.editCongelamientoFee === fee;
        btn.className = active
          ? 'py-2 px-2.5 rounded-xl border-2 border-[#C8102E] bg-[#C8102E] text-white text-xs font-extrabold transition'
          : 'py-2 px-2.5 rounded-xl border border-slate-200 bg-white text-slate-700 hover:border-[#C8102E] text-xs font-extrabold transition';
      });
    }

    window.selectEditCongelamientoFee = function (fee) {
      state.editCongelamientoFee = fee;
      updateEditCongelamientoSubpanelUI();
    };

   function renderEditProceduresGrid() {
    const container = document.getElementById('edit-procedures-grid');
    const visibleProcedures = getAllowedProceduresForCurrentUser();
    container.innerHTML = visibleProcedures.map((p) => {
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
    const proc = PROCEDURES.find((p) => p.key === key);
    if (!proc || !canUserAccessProcedure(proc)) return;
    if (state.editSelectedProcedures.has(key)) state.editSelectedProcedures.delete(key);
    else state.editSelectedProcedures.add(key);
    renderEditProceduresGrid();
    updateEditCongelamientoSubpanelUI();
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

      const hasCongelamiento = state.editSelectedProcedures.has('congelamiento');
      if (hasCongelamiento && !['GRATIS', '300', '500'].includes(state.editCongelamientoFee)) {
        showToast('Seleccione si el retiro/congelamiento es Retiro Gratis, 300 Bs o 500 Bs.', 'error');
        return;
      }

      const congelamientoDetalle = hasCongelamiento
        ? {
            cartaRecepcionada: Boolean(document.getElementById('edit-cong-carta-check')?.checked),
            tipoArancel: state.editCongelamientoFee,
            montoBs: state.editCongelamientoFee === '300' ? 300 : state.editCongelamientoFee === '500' ? 500 : 0
          }
        : null;

      const tramites = {};
    PROCEDURES.forEach((p) => {
      tramites[p.key] = canUserAccessProcedure(p) && state.editSelectedProcedures.has(p.key);
    });
      const mesAnio = fecha.slice(0, 7);
      const canalAtencion = state.editChannelAtencion === 'wati' ? 'wati' : 'plataforma';

      if (state.firebase.db && state.firebase.auth?.currentUser && !id.startsWith('local_')) {
        try {
          const docRef = doc(state.firebase.db, 'artifacts', appId, 'public', 'data', 'atenciones', id);
          const updates = canEditStudentCode
            ? { registro, fecha, mesAnio, canalAtencion, tramites, congelamientoDetalle, observaciones }
            : { canalAtencion, tramites, congelamientoDetalle, observaciones };
          await updateDoc(docRef, updates);
        } catch (err) {
          console.error('Error actualizando en Firestore:', err);
          showToast('Error al actualizar en Firestore.', 'error');
          return;
        }
      } else {
        const idx = state.records.findIndex((r) => r.id === id);
        if (idx !== -1) {
          state.records[idx] = { ...state.records[idx], registro, fecha, mesAnio, canalAtencion, tramites, congelamientoDetalle, observaciones };
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
   function getFilteredReportRecords(advisorFilter, channelFilter = state.reportChannelFilter) {
      const selectedMonth = document.getElementById('report-month-select').value || getTodayISO().slice(0, 7);
      const selectedAdvisor = advisorFilter
        || document.getElementById('report-advisor-select').value
        || 'ALL';
      const selectedChannel = channelFilter || 'ALL';
      const dateRange = getDateRangeFilter('report');

      return state.records.filter((record) => {
        if (!isDateInRange(record.fecha || '', dateRange)) return false;
        if (!dateRange.hasRange && record.mesAnio !== selectedMonth) return false;
        if (selectedAdvisor !== 'ALL' && record.asesor !== selectedAdvisor) return false;
        if (selectedChannel !== 'ALL' && record.canalAtencion !== selectedChannel) return false;
        return true;
      });
    }

    const REPORT_SECTION_LABELS = {
      kpis: 'Tarjetas KPI',
      tramites: 'Trámites (Gráfico y Tabla)',
      diario: 'Evolución Diaria',
      asesores: 'Asesores y Horarios',
    congelamientos: 'Retiro / Congelamiento'
  };

    function applyReportSectionsVisibilityUI() {
      Object.entries(REPORT_SECTION_LABELS).forEach(([key, label]) => {
        const visible = state.reportSections[key] !== false;
        const sectionEl = document.getElementById(`report-section-${key}`);
        const btnEl = document.getElementById(`sec-toggle-${key}`);

        if (sectionEl) sectionEl.classList.toggle('hidden', !visible);
        if (btnEl) {
          btnEl.innerHTML = `<span>${visible ? '✓' : '✕'} ${label}</span>`;
          btnEl.className = visible
            ? 'inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold border bg-[#FDF2F4] text-[#C8102E] border-red-200 transition'
            : 'inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold border bg-slate-50 text-slate-400 border-slate-200 hover:text-slate-600 transition';
        }
      });

      // Redimensionar instancias Chart.js al volver a mostrarlas
      requestAnimationFrame(() => {
        if (state.reportSections.tramites) state.charts.procedures?.resize();
        if (state.reportSections.diario) state.charts.daily?.resize();
        if (state.reportSections.asesores) {
          state.charts.advisors?.resize();
          state.charts.hourly?.resize();
        }
      });
    }

    window.toggleReportSection = function (sectionKey) {
      if (!(sectionKey in state.reportSections)) return;
      const activeCount = Object.values(state.reportSections).filter(Boolean).length;
      if (state.reportSections[sectionKey] && activeCount <= 1) {
        showToast('Debe mantener al menos una sección visible en el reporte.', 'info');
        return;
      }
      state.reportSections[sectionKey] = !state.reportSections[sectionKey];
      applyReportSectionsVisibilityUI();
    };

    window.setReportSectionsPreset = function (preset) {
      if (preset === 'CONG_ONLY') {
        state.reportSections = {
          kpis: false,
          tramites: false,
          diario: false,
          asesores: false,
          congelamientos: true
        };
      } else {
        state.reportSections = {
          kpis: true,
          tramites: true,
          diario: true,
          asesores: true,
          congelamientos: true
        };
      }
      applyReportSectionsVisibilityUI();
    };

    window.setReportChannelFilter = function (channel) {
      state.reportChannelFilter = ['wati', 'plataforma'].includes(channel) ? channel : 'ALL';
      renderMonthlyReport();
    };

    function updateReportChannelUI() {
      const active = state.reportChannelFilter || 'ALL';
      const btnAll = document.getElementById('report-channel-btn-ALL');
      const btnPlat = document.getElementById('report-channel-btn-plataforma');
      const btnWati = document.getElementById('report-channel-btn-wati');
      const dotPlat = document.getElementById('report-channel-dot-plataforma');
      const dotWati = document.getElementById('report-channel-dot-wati');
      const statusBadge = document.getElementById('report-channel-status-badge');
      const cardPlat = document.getElementById('kpi-card-platform');
      const cardWati = document.getElementById('kpi-card-wati');
      const badgePlat = document.getElementById('kpi-badge-platform');
      const badgeWati = document.getElementById('kpi-badge-wati');

      const baseBtn = 'inline-flex items-center gap-2 px-3.5 py-2 rounded-xl text-xs font-bold border transition';
      const inactiveBtn = `${baseBtn} bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100`;

      if (btnAll) {
        btnAll.className = active === 'ALL'
          ? `${baseBtn} bg-slate-800 text-white border-slate-800 shadow-sm`
          : inactiveBtn;
      }
      if (btnPlat) {
        btnPlat.className = active === 'plataforma'
          ? `${baseBtn} bg-[#C8102E] text-white border-[#C8102E] shadow-sm`
          : inactiveBtn;
      }
      if (btnWati) {
        btnWati.className = active === 'wati'
          ? `${baseBtn} bg-emerald-600 text-white border-emerald-600 shadow-sm`
          : inactiveBtn;
      }

      if (dotPlat) {
        dotPlat.className = `w-2 h-2 rounded-full ${active === 'plataforma' ? 'bg-white' : 'bg-[#C8102E]'}`;
      }
      if (dotWati) {
        dotWati.className = `w-2 h-2 rounded-full ${active === 'wati' ? 'bg-white' : 'bg-emerald-500'}`;
      }

      if (statusBadge) {
        if (active === 'wati') {
          statusBadge.textContent = 'Mostrando solo WATI';
          statusBadge.className = 'px-3 py-1 rounded-full text-[11px] font-bold bg-emerald-50 text-emerald-700 border border-emerald-200';
        } else if (active === 'plataforma') {
          statusBadge.textContent = 'Mostrando solo Plataforma';
          statusBadge.className = 'px-3 py-1 rounded-full text-[11px] font-bold bg-[#FDF2F4] text-[#C8102E] border border-red-200';
        } else {
          statusBadge.textContent = 'Mostrando todos los canales';
          statusBadge.className = 'px-3 py-1 rounded-full text-[11px] font-bold bg-slate-100 text-slate-600 border border-slate-200';
        }
      }

      if (cardPlat && badgePlat) {
        if (active === 'plataforma') {
          cardPlat.className = 'bg-[#FDF2F4]/40 p-4 rounded-2xl border-2 border-[#C8102E] ring-4 ring-red-100 shadow-md transition-all duration-200';
          badgePlat.classList.remove('hidden');
        } else if (active === 'wati') {
          cardPlat.className = 'bg-white p-4 rounded-2xl border border-slate-200 shadow-sm opacity-50 transition-all duration-200';
          badgePlat.classList.add('hidden');
        } else {
          cardPlat.className = 'bg-white p-4 rounded-2xl border border-red-200 shadow-sm transition-all duration-200';
          badgePlat.classList.add('hidden');
        }
      }

      if (cardWati && badgeWati) {
        if (active === 'wati') {
          cardWati.className = 'bg-emerald-50/40 p-4 rounded-2xl border-2 border-emerald-500 ring-4 ring-emerald-100 shadow-md transition-all duration-200';
          badgeWati.classList.remove('hidden');
        } else if (active === 'plataforma') {
          cardWati.className = 'bg-white p-4 rounded-2xl border border-slate-200 shadow-sm opacity-50 transition-all duration-200';
          badgeWati.classList.add('hidden');
        } else {
          cardWati.className = 'bg-white p-4 rounded-2xl border border-emerald-200 shadow-sm transition-all duration-200';
          badgeWati.classList.add('hidden');
        }
      }
    }

    function renderMonthlyReport() {
      updateReportAdvisorsHeading();
      updateReportChannelUI();
      applyReportSectionsVisibilityUI();
      const selectedMonth = document.getElementById('report-month-select').value || getTodayISO().slice(0, 7);
      const selectedAdvisor = document.getElementById('report-advisor-select').value || 'ALL';
      const selectedChannel = state.reportChannelFilter || 'ALL';
      const dateRange = getDateRangeFilter('report');

      const periodAllRecords = getFilteredReportRecords('ALL', selectedChannel);
      const advisorAllChannelsRecords = getFilteredReportRecords(selectedAdvisor, 'ALL');
      const filteredRecords = getFilteredReportRecords(selectedAdvisor, selectedChannel);

      const totalAttentions = filteredRecords.length;
      const totalChannelBase = advisorAllChannelsRecords.length;
      const watiAttentions = advisorAllChannelsRecords.filter((record) => record.canalAtencion === 'wati').length;
      const platformAttentions = totalChannelBase - watiAttentions;
      const platformPct = totalChannelBase > 0
        ? Math.round((platformAttentions / totalChannelBase) * 100)
        : 0;
      const watiPct = totalChannelBase > 0
        ? Math.round((watiAttentions / totalChannelBase) * 100)
        : 0;

      document.getElementById('kpi-channel-platform-count').textContent =
        platformAttentions.toLocaleString('es-BO');
      document.getElementById('kpi-channel-platform-pct').textContent = `${platformPct}%`;
      document.getElementById('kpi-channel-platform-bar').style.width = `${platformPct}%`;
      document.getElementById('kpi-channel-wati-count').textContent =
        watiAttentions.toLocaleString('es-BO');
      document.getElementById('kpi-channel-wati-pct').textContent = `${watiPct}%`;
      document.getElementById('kpi-channel-wati-bar').style.width = `${watiPct}%`;

      const uniqueAdvisorsInPeriod = new Set(periodAllRecords.map((r) => r.asesor));
      const advisorMultiplier = selectedAdvisor === 'ALL' ? Math.max(1, uniqueAdvisorsInPeriod.size) : 1;
      const minGoal = 1000 * advisorMultiplier;
      const maxGoal = 1500 * advisorMultiplier;
      const midGoal = 1250 * advisorMultiplier;
      const goalPct = dateRange.hasRange
        ? null
        : Math.min(100, Math.round((totalAttentions / midGoal) * 100));

     const isWatiTheme = selectedChannel === 'wati';

      const goalBadgeEl = document.getElementById('kpi-goal-badge');
      goalBadgeEl.textContent = dateRange.hasRange
        ? 'Rango seleccionado'
        : `Meta: ${minGoal.toLocaleString('es-BO')}-${maxGoal.toLocaleString('es-BO')}`;
      goalBadgeEl.className = isWatiTheme
        ? 'px-2 py-0.5 rounded-full text-[10px] font-extrabold bg-emerald-50 text-emerald-700'
        : 'px-2 py-0.5 rounded-full text-[10px] font-extrabold bg-[#FDF2F4] text-[#C8102E]';

      document.getElementById('kpi-total-attentions').textContent = totalAttentions.toLocaleString('es-BO');

      const goalPctEl = document.getElementById('kpi-goal-pct');
      goalPctEl.textContent = goalPct === null ? '—' : `${goalPct}%`;
      goalPctEl.className = isWatiTheme ? 'text-emerald-700 font-bold' : 'text-[#C8102E] font-bold';

      const goalBarEl = document.getElementById('kpi-goal-bar');
      goalBarEl.style.width = `${goalPct || 0}%`;
      goalBarEl.className = isWatiTheme
        ? 'h-full bg-emerald-600 rounded-full transition-all duration-500'
        : 'h-full bg-[#C8102E] rounded-full transition-all duration-500';

      document.getElementById('kpi-total-label').textContent =
        dateRange.hasRange ? 'Total Atenciones del Periodo' : 'Total Atenciones Mes';
      document.getElementById('kpi-goal-caption').textContent =
        dateRange.hasRange ? 'Meta mensual no aplicable' : 'Cumplimiento meta';

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

      const procCounts = PROCEDURES.map((p) => {
        let count = 0;
        filteredRecords.forEach((r) => {
          if (r.tramites?.[p.key]) count++;
        });
        return { ...p, count };
      }).sort((a, b) => b.count - a.count);

      const totalProcMarks = procCounts.reduce((acc, item) => acc + item.count, 0);
      document.getElementById('total-marks-badge').textContent = `${totalProcMarks.toLocaleString('es-BO')} trámites (X)`;

      const topProcDotEl = document.getElementById('kpi-top-procedure-dot');
      if (topProcDotEl) {
        topProcDotEl.className = isWatiTheme
          ? 'w-2.5 h-2.5 rounded-full bg-emerald-500'
          : 'w-2.5 h-2.5 rounded-full bg-[#C8102E]';
      }

      const topProcEl = document.getElementById('kpi-top-procedure');
      topProcEl.className = isWatiTheme
        ? 'text-xl font-black text-emerald-700 truncate mt-1'
        : 'text-xl font-black text-[#C8102E] truncate mt-1';

      const topProc = procCounts[0];
      if (topProc && topProc.count > 0) {
        const sharePct = totalProcMarks > 0 ? ((topProc.count / totalProcMarks) * 100).toFixed(1) : '0.0';
        topProcEl.textContent = topProc.label;
        document.getElementById('kpi-top-procedure-stats').textContent =
          `${topProc.count.toLocaleString('es-BO')} solicitudes (${sharePct}% de trámites)`;
      } else {
        topProcEl.textContent = 'Sin datos';
        document.getElementById('kpi-top-procedure-stats').textContent = '0 solicitudes';
      }

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

      document.getElementById('report-procedures-table').innerHTML = procCounts.map((item, idx) => {
        const pct = totalProcMarks > 0 ? ((item.count / totalProcMarks) * 100).toFixed(1) : '0.0';
        const rankBadgeClass = idx < 3
          ? (isWatiTheme ? 'bg-emerald-600 text-white' : 'bg-[#C8102E] text-white')
          : 'bg-slate-100 text-slate-600';
        const pctTextClass = isWatiTheme ? 'text-emerald-700' : 'text-[#C8102E]';
        return `
          <tr class="hover:bg-slate-50">
            <td class="py-2 px-2.5 font-bold text-slate-800 flex items-center gap-2">
              <span class="w-4 h-4 rounded text-[10px] font-black flex items-center justify-center ${rankBadgeClass}">${idx + 1}</span>
              <span>${item.label}</span>
            </td>
            <td class="py-2 px-2 text-right font-extrabold text-slate-900">${item.count.toLocaleString('es-BO')}</td>
            <td class="py-2 px-2.5 text-right font-bold ${pctTextClass}">${pct}%</td>
          </tr>
        `;
      }).join('');

      const periodLabel = dateRange.hasRange
        ? `Del ${dateRange.from ? formatDateDisplay(dateRange.from) : 'inicio'} al ${dateRange.to ? formatDateDisplay(dateRange.to) : 'presente'}`
        : formatMonthName(selectedMonth);
      const channelSuffix = selectedChannel === 'wati'
        ? ' — Solo WATI'
        : selectedChannel === 'plataforma'
          ? ' — Solo Plataforma'
          : '';
      const channelMetaLabel = selectedChannel === 'wati'
        ? 'Solo WATI'
        : selectedChannel === 'plataforma'
          ? 'Solo Plataforma'
          : 'Todos los canales';

      document.getElementById('report-title').textContent =
        (dateRange.hasRange ? 'Reporte de Atención SAE' : 'Reporte Mensual de Atención SAE') + channelSuffix;
      document.getElementById('print-meta-period').textContent = `Periodo: ${periodLabel}`;
      document.getElementById('print-meta-advisor').textContent =
        `Asesor: ${selectedAdvisor === 'ALL' ? 'Consolidado General SAE' : selectedAdvisor}`;
      const printChannelEl = document.getElementById('print-meta-channel');
      if (printChannelEl) printChannelEl.textContent = `Canal: ${channelMetaLabel}`;
      document.getElementById('print-meta-date').textContent = `Generado: ${formatDateDisplay(getTodayISO())}`;

     renderProceduresChart(procCounts);
      renderDailyTrendChart(selectedMonth, dailyMap, dateRange.hasRange);
      renderAdvisorsChart(periodAllRecords);
      renderHourlyChart(filteredRecords);
      renderCongelamientosReport(filteredRecords);
    }

    function renderCongelamientosReport(recordsList) {
      const congRecords = recordsList.filter((r) => r.tramites?.congelamiento);
      const totalCong = congRecords.length;
      const totalAll = recordsList.length;

      let cartasSi = 0;
      const buckets = {
        GRATIS: { label: 'Retiro Gratis', count: 0, cartas: 0, sinCarta: 0 },
        '300': { label: 'Arancel 300 Bs', count: 0, cartas: 0, sinCarta: 0 },
        '500': { label: 'Arancel 500 Bs', count: 0, cartas: 0, sinCarta: 0 },
        SIN_DETALLE: { label: 'Sin especificar (Histórico)', count: 0, cartas: 0, sinCarta: 0 }
      };

      congRecords.forEach((r) => {
        const det = r.congelamientoDetalle;
        const hasCarta = Boolean(det?.cartaRecepcionada);
        if (hasCarta) cartasSi++;

        const key = ['GRATIS', '300', '500'].includes(det?.tipoArancel) ? det.tipoArancel : 'SIN_DETALLE';
        buckets[key].count++;
        if (hasCarta) buckets[key].cartas++;
        else buckets[key].sinCarta++;
      });

      const shareAllPct = totalAll > 0 ? ((totalCong / totalAll) * 100).toFixed(1) : '0.0';
      const cartasPct = totalCong > 0 ? ((cartasSi / totalCong) * 100).toFixed(1) : '0.0';
      const gratisPct = totalCong > 0 ? ((buckets.GRATIS.count / totalCong) * 100).toFixed(1) : '0.0';
      const pagadosCount = buckets['300'].count + buckets['500'].count;

      const setTxt = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
      setTxt('cong-kpi-total', totalCong.toLocaleString('es-BO'));
      setTxt('cong-kpi-share', `${shareAllPct}% del total de atenciones`);
      setTxt('cong-kpi-cartas', `${cartasSi} / ${totalCong}`);
      setTxt('cong-kpi-cartas-pct', `${cartasPct}% con carta entregada`);
      setTxt('cong-kpi-gratis', buckets.GRATIS.count.toLocaleString('es-BO'));
      setTxt('cong-kpi-gratis-pct', `${gratisPct}% de retiros/congelamientos`);
      setTxt('cong-kpi-pagados', pagadosCount.toLocaleString('es-BO'));
      setTxt('cong-kpi-pagados-desglose', `${buckets['300'].count} de 300 Bs · ${buckets['500'].count} de 500 Bs`);

      const tbody = document.getElementById('cong-report-table-body');
      if (!tbody) return;

      if (totalCong === 0) {
        tbody.innerHTML = `
          <tr>
           <td colspan="5" class="py-6 text-center text-slate-400">No se registraron retiros/congelamientos en el periodo seleccionado.</td>
          </tr>
        `;
        return;
      }

      const order = ['GRATIS', '300', '500', 'SIN_DETALLE'].filter((k) => k !== 'SIN_DETALLE' || buckets[k].count > 0);
      tbody.innerHTML = order.map((k) => {
        const b = buckets[k];
        const pct = totalCong > 0 ? ((b.count / totalCong) * 100).toFixed(1) : '0.0';
        return `
          <tr class="hover:bg-slate-50">
            <td class="py-2.5 px-3 font-bold text-slate-800">${b.label}</td>
            <td class="py-2.5 px-3 text-right font-extrabold text-slate-900">${b.count}</td>
            <td class="py-2.5 px-3 text-right font-semibold text-slate-600">${pct}%</td>
            <td class="py-2.5 px-3 text-right font-bold text-emerald-700">${b.cartas}</td>
            <td class="py-2.5 px-3 text-right font-semibold text-amber-700">${b.sinCarta}</td>
          </tr>
        `;
      }).join('');
    }

    window.exportCongelamientosExcel = function () {
      const congRows = getFilteredReportRecords().filter((r) => r.tramites?.congelamiento);
      if (congRows.length === 0) {
      showToast('No hay retiros/congelamientos para exportar en el filtro actual.', 'error');
      return;
    }

      const thStyle = 'background-color:#C8102E;color:#FFFFFF;font-family:Calibri,sans-serif;font-size:11pt;font-weight:bold;text-align:center;border:1px solid #9E0B22;padding:5px 8px;';
      const tdStyle = 'font-family:Calibri,sans-serif;font-size:11pt;border:1px solid #D9D9D9;padding:4px 8px;';

      const bodyRows = congRows.map((r) => {
        const det = r.congelamientoDetalle;
        const cartaTxt = det?.cartaRecepcionada ? 'SÍ' : 'NO';
        const modalTxt = det?.tipoArancel === 'GRATIS'
          ? 'RETIRO GRATIS'
          : det?.tipoArancel === '300'
            ? '300 BS'
            : det?.tipoArancel === '500'
              ? '500 BS'
              : 'SIN ESPECIFICAR';

        return `
          <tr>
            <td style="${tdStyle}text-align:center;">${escapeHtml(formatDateDisplay(r.fecha))}</td>
            <td style="${tdStyle}text-align:center;font-weight:bold;">${escapeHtml(r.registro)}</td>
            <td style="${tdStyle}text-align:center;">${r.canalAtencion === 'wati' ? 'WATI' : 'PLATAFORMA'}</td>
            <td style="${tdStyle}text-align:center;font-weight:bold;color:${det?.cartaRecepcionada ? '#047857' : '#B45309'};">${cartaTxt}</td>
            <td style="${tdStyle}text-align:center;font-weight:bold;">${modalTxt}</td>
            <td style="${tdStyle}">${escapeHtml(r.observaciones || '')}</td>
            <td style="${tdStyle}text-align:center;">${escapeHtml(r.asesor || '')}</td>
          </tr>
        `;
      }).join('');

      const html = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
        <head><meta charset="UTF-8"></head>
        <body>
          <table style="border-collapse:collapse;">
            <tr>
             <td colspan="7" style="font-family:Calibri,sans-serif;font-size:12pt;font-weight:bold;color:#C8102E;padding:6px 2px;">UTEPSA SAE — CONTROL DE RETIRO / CONGELAMIENTO Y CARTAS RECEPCIONADAS</td>
            </tr>
            <tr>
              <th style="${thStyle}">FECHA</th>
              <th style="${thStyle}">REGISTRO</th>
              <th style="${thStyle}">CANAL</th>
              <th style="${thStyle}">CARTA RECEPCIONADA</th>
              <th style="${thStyle}">MODALIDAD</th>
              <th style="${thStyle}">OBSERVACIONES</th>
              <th style="${thStyle}">ASESOR</th>
            </tr>
            ${bodyRows}
          </table>
        </body>
        </html>
      `;

      const blob = new Blob(['\uFEFF' + html], { type: 'application/vnd.ms-excel;charset=utf-8;' });
      triggerDownload(blob, `Reporte_Retiro_Congelamiento_UTEPSA_SAE_${getTodayISO()}.xls`);
    showToast(`Reporte de Retiro / Congelamiento exportado (${congRows.length} registros).`, 'success');
    };

    window.renderMonthlyReport = async function () {
      renderMonthlyReport();

      const selectedMonth = document.getElementById('report-month-select')?.value || getTodayISO().slice(0, 7);
      const dateRange = getDateRangeFilter('report');
      if (!dateRange.valid) return;

      const monthsNeeded = dateRange.hasRange
        ? getMonthsBetweenDates(dateRange.from, dateRange.to)
        : [selectedMonth];

      const loadedNew = await ensureMonthsLoaded(monthsNeeded);
      if (loadedNew) {
        renderMonthlyReport();
      }
    };

    window.clearReportFilters = function () {
      document.getElementById('report-month-select').value = getTodayISO().slice(0, 7);
      document.getElementById('report-date-from').value = '';
      document.getElementById('report-date-to').value = '';
      document.getElementById('report-advisor-select').value =
        isAdvisorRole() ? state.currentUser.name : 'ALL';
      state.reportChannelFilter = 'ALL';
      renderMonthlyReport();
    };

    function updateReportAdvisorsHeading() {
      const isAdvisor = isAdvisorRole();
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
      const isWati = state.reportChannelFilter === 'wati';

      state.charts.procedures = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: procCounts.map((p) => p.label),
          datasets: [{
            label: 'Solicitudes (X)',
            data: procCounts.map((p) => p.count),
            backgroundColor: procCounts.map((_, i) =>
              isWati
                ? (i === 0 ? '#059669' : i < 3 ? '#10B981' : '#63666A')
                : (i === 0 ? '#C8102E' : i < 3 ? '#E11D48' : '#63666A')
            ),
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
      const isWati = state.reportChannelFilter === 'wati';

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
            borderColor: isWati ? '#059669' : '#C8102E',
            backgroundColor: isWati ? 'rgba(16, 185, 129, 0.14)' : 'rgba(200, 16, 46, 0.12)',
            borderWidth: 2.5,
            fill: true,
            tension: 0.32,
            pointRadius: 3,
            pointHoverRadius: 6,
            pointBackgroundColor: isWati ? '#059669' : '#C8102E'
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
      const isWati = state.reportChannelFilter === 'wati';

      const advMap = new Map();
      periodRecords.forEach((r) => {
        advMap.set(r.asesor, (advMap.get(r.asesor) || 0) + 1);
      });

      const sortedAdvisors = Array.from(advMap.entries()).sort((a, b) => b[1] - a[1]);
      const totalPeriod = periodRecords.length;

      const summaryContainer = document.getElementById('advisors-percentage-summary');
      if (summaryContainer) {
        if (sortedAdvisors.length === 0) {
          summaryContainer.innerHTML = '';
        } else {
          summaryContainer.innerHTML = sortedAdvisors.map(([name, count], idx) => {
            const pct = totalPeriod > 0 ? ((count / totalPeriod) * 100).toFixed(1) : '0.0';
            const badgeStyle = idx === 0
              ? (isWati ? 'bg-emerald-50 text-emerald-800 border-emerald-200' : 'bg-[#FDF2F4] text-[#C8102E] border-red-200')
              : 'bg-slate-50 text-slate-700 border-slate-200';
            return `
              <span class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-[11px] font-bold border ${badgeStyle}">
                <span>${escapeHtml(name)}:</span>
                <span>${count.toLocaleString('es-BO')} (${pct}%)</span>
              </span>
            `;
          }).join('');
        }
      }

      state.charts.advisors = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: sortedAdvisors.map((a) => a[0]),
          datasets: [{
            label: 'Atenciones del Periodo',
            data: sortedAdvisors.map((a) => a[1]),
            backgroundColor: isWati
              ? ['#059669', '#63666A', '#10B981', '#94A3B8']
              : ['#C8102E', '#63666A', '#94A3B8', '#E11D48'],
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
              cornerRadius: 10,
              callbacks: {
                label: (context) => {
                  const val = context.parsed.y || 0;
                  const pct = totalPeriod > 0 ? ((val / totalPeriod) * 100).toFixed(1) : '0.0';
                  return ` Atenciones: ${val.toLocaleString('es-BO')} (${pct}%)`;
                }
              }
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

    function renderHourlyChart(recordsList) {
      const canvas = document.getElementById('chart-hourly');
      if (!canvas) return;
      const ctx = canvas.getContext('2d');
      if (state.charts.hourly) state.charts.hourly.destroy();
      const isWati = state.reportChannelFilter === 'wati';

      const hourFormatter = new Intl.DateTimeFormat('en-GB', {
        timeZone: 'America/La_Paz',
        hour: '2-digit',
        hourCycle: 'h23'
      });

      const hourCounts = new Map();
      for (let h = 7; h <= 22; h++) {
        hourCounts.set(h, 0);
      }

      recordsList.forEach((r) => {
        if (!Number.isFinite(r.createdAt) || r.createdAt <= 0) return;
        const h = parseInt(hourFormatter.format(new Date(r.createdAt)), 10);
        if (!Number.isNaN(h)) {
          hourCounts.set(h, (hourCounts.get(h) || 0) + 1);
        }
      });

      const sortedHours = Array.from(hourCounts.keys()).sort((a, b) => a - b);
      const labels = sortedHours.map((h) => `${String(h).padStart(2, '0')}:00`);
      const values = sortedHours.map((h) => hourCounts.get(h) || 0);

      let peakHourLabel = '-';
      let peakHourVal = 0;
      sortedHours.forEach((h) => {
        const val = hourCounts.get(h) || 0;
        if (val > peakHourVal) {
          peakHourVal = val;
          peakHourLabel = `${String(h).padStart(2, '0')}:00 - ${String(h + 1).padStart(2, '0')}:00`;
        }
      });

      const peakBadge = document.getElementById('kpi-peak-hour-badge');
      if (peakBadge) {
        peakBadge.textContent = peakHourVal > 0
          ? `Hora pico: ${peakHourLabel} (${peakHourVal})`
          : 'Hora pico: -';
      }

      state.charts.hourly = new Chart(ctx, {
        type: 'bar',
        data: {
          labels,
          datasets: [{
            label: 'Atenciones',
            data: values,
            backgroundColor: values.map((v) =>
              v === peakHourVal && v > 0
                ? (isWati ? '#059669' : '#C8102E')
                : (isWati ? 'rgba(16, 185, 129, 0.55)' : 'rgba(200, 16, 46, 0.45)')
            ),
            borderRadius: 6,
            barPercentage: 0.72
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
              ticks: { font: { size: 10, weight: '600' }, color: '#64748B' }
            },
            y: {
              beginAtZero: true,
              grid: { color: '#F1F5F9' },
              ticks: { precision: 0, font: { size: 11 }, color: '#64748B' }
            }
          }
        }
      });
    }

    // =========================================================================
    // 11. EXPORTACIÓN EXCEL UTEPSA (.XLS Y .CSV) E IMPRESIÓN EJECUTIVA
    function getRecordsForExport(source) {
      if (source === 'historial') {
        return getFilteredHistoryRecords();
      }
      return getFilteredReportRecords();
    }

    window.exportToExcelXLS = function (source) {
      const rows = getRecordsForExport(source);
      if (rows.length === 0) {
        showToast('No hay registros para exportar en el filtro actual.', 'error');
        return;
      }

      const monthCounts = getStudentMonthCounts(state.records);
      const selectedFilterAdvisor = source === 'reporte'
        ? document.getElementById('report-advisor-select')?.value
        : document.getElementById('hist-filter-advisor')?.value;
      const advisorTitle = (selectedFilterAdvisor && selectedFilterAdvisor !== 'ALL')
        ? selectedFilterAdvisor
        : (state.currentUser?.name || 'SAE UTEPSA');

      const colWidthsPx = [102, 86, 114, 76, 106, 101, 92, 63, 109, 80, 58, 94, 94, 94, 132, 333, 81];
      const totalTableWidth = colWidthsPx.reduce((acc, w) => acc + w, 0);
      const colGroupHtml = `<colgroup>${colWidthsPx.map((w) => `<col width="${w}" style="width:${w}px;mso-width-source:userset;">`).join('')}</colgroup>`;

      const thBaseStyle = 'background-color:#C8102E;color:#FFFFFF;font-family:Calibri,sans-serif;font-size:11pt;font-weight:bold;text-align:center;vertical-align:middle;white-space:normal;border:none;padding:4px;';

      const tableRowsHtml = rows.map((r) => {
        const isRepeat = (monthCounts.get(`${r.mesAnio}|${r.registro}`) || 0) > 1;
        const regBg = isRepeat ? 'background-color:#F8CBAD;font-weight:bold;' : '';
        const procCols = PROCEDURES.map((p) =>
          `<td style="text-align:center;vertical-align:middle;border:1px solid #D9D9D9;font-family:Calibri,sans-serif;font-size:11pt;">${r.tramites?.[p.key] ? 'X' : ''}</td>`
        ).join('');

        return `
          <tr>
            <td style="text-align:center;vertical-align:middle;border:1px solid #D9D9D9;font-family:Calibri,sans-serif;font-size:11pt;">${escapeHtml(formatDateDisplay(r.fecha))}</td>
            <td style="text-align:center;vertical-align:middle;border:1px solid #D9D9D9;font-family:Calibri,sans-serif;font-size:11pt;${regBg}">${escapeHtml(r.registro)}</td>
            <td style="text-align:center;vertical-align:middle;border:1px solid #D9D9D9;font-family:Calibri,sans-serif;font-size:11pt;">${r.canalAtencion === 'wati' ? 'WATI' : 'PLATAFORMA'}</td>
            ${procCols}
            <td style="vertical-align:middle;white-space:normal;border:1px solid #D9D9D9;font-family:Calibri,sans-serif;font-size:11pt;padding:4px 6px;">${escapeHtml(r.observaciones || '')}</td>
            <td style="text-align:center;vertical-align:middle;border:1px solid #D9D9D9;font-family:Calibri,sans-serif;font-size:11pt;">${escapeHtml(r.asesor || '')}</td>
          </tr>
        `;
      }).join('');

     let executiveSummaryHtml = '';
      if (source === 'reporte') {
        const advisorSummaryMap = new Map();
        rows.forEach((r) => {
          const adv = r.asesor || 'Sin asignar';
          const curr = advisorSummaryMap.get(adv) || { asesor: adv, plataforma: 0, wati: 0, total: 0 };
          if (r.canalAtencion === 'wati') curr.wati += 1;
          else curr.plataforma += 1;
          curr.total += 1;
          advisorSummaryMap.set(adv, curr);
        });

        const summaryList = Array.from(advisorSummaryMap.values()).sort((a, b) => b.total - a.total);
        const totalPlat = summaryList.reduce((acc, s) => acc + s.plataforma, 0);
        const totalWati = summaryList.reduce((acc, s) => acc + s.wati, 0);
        const grandTotal = rows.length;

        const sumThStyle = 'background-color:#1E293B;color:#FFFFFF;font-family:Calibri,sans-serif;font-size:10pt;font-weight:bold;text-align:center;border:1px solid #1E293B;padding:4px 8px;';
        const sumTdStyle = 'font-family:Calibri,sans-serif;font-size:10pt;border:1px solid #D9D9D9;padding:4px 8px;';

        executiveSummaryHtml = `
          <table style="border-collapse:collapse;margin-bottom:14px;">
            <tr>
              <td colspan="5" style="font-family:Calibri,sans-serif;font-size:11pt;font-weight:bold;color:#C8102E;padding:4px 2px;">RESUMEN EJECUTIVO POR ASESOR Y CANAL</td>
            </tr>
            <tr>
              <th style="${sumThStyle}text-align:left;width:240px;">ASESOR SAE</th>
              <th style="${sumThStyle}width:110px;">PLATAFORMA</th>
              <th style="${sumThStyle}width:110px;">WATI</th>
              <th style="${sumThStyle}width:110px;">TOTAL</th>
              <th style="${sumThStyle}width:110px;">% PARTICIPACIÓN</th>
            </tr>
            ${summaryList.map((s) => {
              const pct = grandTotal > 0 ? ((s.total / grandTotal) * 100).toFixed(1) : '0.0';
              return `
                <tr>
                  <td style="${sumTdStyle}font-weight:bold;">${escapeHtml(s.asesor)}</td>
                  <td style="${sumTdStyle}text-align:center;">${s.plataforma}</td>
                  <td style="${sumTdStyle}text-align:center;">${s.wati}</td>
                  <td style="${sumTdStyle}text-align:center;font-weight:bold;">${s.total}</td>
                  <td style="${sumTdStyle}text-align:center;">${pct}%</td>
                </tr>
              `;
            }).join('')}
            <tr style="background-color:#F1F5F9;font-weight:bold;">
              <td style="${sumTdStyle}font-weight:bold;">TOTAL CONSOLIDADO</td>
              <td style="${sumTdStyle}text-align:center;font-weight:bold;">${totalPlat}</td>
              <td style="${sumTdStyle}text-align:center;font-weight:bold;">${totalWati}</td>
              <td style="${sumTdStyle}text-align:center;font-weight:bold;color:#C8102E;">${grandTotal}</td>
              <td style="${sumTdStyle}text-align:center;font-weight:bold;">100.0%</td>
            </tr>
          </table>
          <br/>
        `;
      }

      const htmlExcel = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <head>
          <meta charset="UTF-8">
          <!--[if gte mso 9]>
          <xml>
            <x:ExcelWorkbook>
              <x:ExcelWorksheets>
                <x:ExcelWorksheet>
                  <x:Name>Planilla_Atencion_UTEPSA_SAE</x:Name>
                  <x:WorksheetOptions>
                    <x:PageSetup>
                      <x:Layout x:Orientation="Landscape"/>
                      <x:PageMargins x:Bottom="0.5" x:Left="0.4" x:Right="0.4" x:Top="0.5"/>
                    </x:PageSetup>
                    <x:FitToPage/>
                    <x:Print>
                      <x:ValidPrinterInfo/>
                      <x:FitWidth>1</x:FitWidth>
                      <x:FitHeight>0</x:FitHeight>
                    </x:Print>
                    <x:DoNotDisplayGridlines/>
                  </x:WorksheetOptions>
                </x:ExcelWorksheet>
              </x:ExcelWorksheets>
            </x:ExcelWorkbook>
          </xml>
          <![endif]-->
          <style>
            table { border-collapse: collapse; }
            @page { size: landscape; margin: 0.4in; }
          </style>
        </head>
       <body>
          ${executiveSummaryHtml}
          <table width="${totalTableWidth}" style="width:${totalTableWidth}px;border-collapse:collapse;">
            ${colGroupHtml}
          <tr style="height:34px;">
            <td colspan="4" style="font-family:Calibri,sans-serif;font-size:10pt;font-weight:bold;color:#C8102E;text-align:left;vertical-align:middle;white-space:normal;border:none;border-bottom:2px solid #9E0B22;padding:4px 2px;">UTEPSA - UNIVERSIDAD TECNOLÓGICA PRIVADA DE SANTA CRUZ</td>
            <td colspan="7" style="font-family:Calibri,sans-serif;font-size:10pt;font-weight:bold;color:#C8102E;text-align:center;vertical-align:middle;white-space:normal;border:none;border-bottom:2px solid #9E0B22;padding:4px 2px;">"Educar para emprender y servir"</td>
            <td colspan="6" style="font-family:Calibri,sans-serif;font-size:10pt;font-weight:bold;color:#C8102E;text-align:right;vertical-align:middle;white-space:normal;border:none;border-bottom:2px solid #9E0B22;padding:4px 2px;">${escapeHtml(advisorTitle.toUpperCase())}</td>
          </tr>
          <tr style="height:30px;">
            <th style="${thBaseStyle}width:${colWidthsPx[0]}px;">FECHA</th>
            <th style="${thBaseStyle}width:${colWidthsPx[1]}px;">REGISTRO</th>
            <th style="${thBaseStyle}width:${colWidthsPx[2]}px;">CANAL</th>
            ${PROCEDURES.map((p, idx) => `<th style="${thBaseStyle}width:${colWidthsPx[idx + 3]}px;">${p.label}</th>`).join('')}
            <th style="${thBaseStyle}width:${colWidthsPx[15]}px;">OBSERVACIONES</th>
            <th style="${thBaseStyle}width:${colWidthsPx[16]}px;">ASESOR</th>
          </tr>
            ${tableRowsHtml}
          </table>
        <\/body>
        <\/html>
      `;

      const channelTag = source === 'reporte' && state.reportChannelFilter && state.reportChannelFilter !== 'ALL'
        ? `_${state.reportChannelFilter.toUpperCase()}`
        : '';
      const blob = new Blob(['\uFEFF' + htmlExcel], { type: 'application/vnd.ms-excel;charset=utf-8;' });
      triggerDownload(blob, `Planilla_Atencion_UTEPSA_SAE${channelTag}_${getTodayISO()}.xls`);
      showToast(`Planilla Excel UTEPSA exportada (${rows.length} registros).`, 'success');
    };

    window.exportToCSV = function (source) {
      const rows = getRecordsForExport(source);
      if (rows.length === 0) {
        showToast('No hay registros para exportar en el filtro actual.', 'error');
        return;
      }

      const headers = [
        'FECHA',
        'REGISTRO',
        'CANAL',
        ...PROCEDURES.map((p) => p.label),
        'OBSERVACIONES',
        'ASESOR'
      ];

      const escapeCsv = (val) => `"${String(val ?? '').replace(/"/g, '""')}"`;

      const csvLines = [
        headers.map(escapeCsv).join(','),
        ...rows.map((r) => [
          formatDateDisplay(r.fecha),
          r.registro,
          r.canalAtencion === 'wati' ? 'WATI' : 'PLATAFORMA',
          ...PROCEDURES.map((p) => (r.tramites?.[p.key] ? 'X' : '')),
          r.observaciones || '',
          r.asesor || ''
        ].map(escapeCsv).join(','))
      ];

      const channelTag = source === 'reporte' && state.reportChannelFilter && state.reportChannelFilter !== 'ALL'
        ? `_${state.reportChannelFilter.toUpperCase()}`
        : '';
      const blob = new Blob(['\uFEFF' + csvLines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
      triggerDownload(blob, `Planilla_Atencion_UTEPSA_SAE${channelTag}_${getTodayISO()}.csv`);
      showToast(`Archivo CSV exportado (${rows.length} registros).`, 'success');
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
      const monthsSet = new Set([currentMonth, ...getRecentMonthsList(12)]);
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
      const hadPrevSelection = histMonth.options.length > 1;
      histMonth.innerHTML = '<option value="ALL">Todos los meses</option>' +
        sortedMonths.map((m) => `<option value="${escapeHtml(m)}">${escapeHtml(formatMonthName(m))}</option>`).join('');
      if (hadPrevSelection && (prevHistMonth === 'ALL' || sortedMonths.includes(prevHistMonth))) {
        histMonth.value = prevHistMonth;
      } else {
        histMonth.value = currentMonth;
      }

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
      const currentAdvisor = isAdvisorRole() ? state.currentUser.name : '';
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

    function syncActiveCalendarPeriod() {
      if (!state.currentUser) return;
      const todayISO = getTodayISO();
      const previousDateISO = state.activeDateISO;
      if (!previousDateISO || todayISO === previousDateISO) return;

      const previousMonth = previousDateISO.slice(0, 7);
      const currentMonth = todayISO.slice(0, 7);
      state.activeDateISO = todayISO;

      const entryDateInput = document.getElementById('entry-fecha');
      if (entryDateInput && (state.currentUser.role !== 'Administrador' || entryDateInput.value === previousDateISO)) {
        entryDateInput.value = todayISO;
      }
      document.getElementById('quick-stat-today-date').textContent = formatDateDisplay(todayISO);

      if (currentMonth !== previousMonth) {
        const histMonth = document.getElementById('hist-filter-month');
        const repMonth = document.getElementById('report-month-select');
        const wasHistOnPrevMonth = histMonth?.value === previousMonth;
        const wasRepOnPrevMonth = repMonth?.value === previousMonth;

        state.loadedMonths.clear();
        state.records = state.records.filter((r) => r.mesAnio === currentMonth);

        populateDynamicFilters();
        if (wasHistOnPrevMonth && histMonth) histMonth.value = currentMonth;
        if (wasRepOnPrevMonth && repMonth) repMonth.value = currentMonth;

        if (state.firebase.db && state.firebase.auth?.currentUser) {
          startFirestoreRealtimeSync();
        }
      }

      refreshAllViews();
      const currentCode = document.getElementById('entry-registro')?.value.trim();
      if (currentCode) handleStudentCodeInput(currentCode);
    }

    function startDateRolloverWatcher() {
      if (state.dateWatcherInterval) clearInterval(state.dateWatcherInterval);
      state.activeDateISO = getTodayISO();

      state.dateWatcherInterval = setInterval(syncActiveCalendarPeriod, 30000);

      if (!window.__saeDateWatcherBound) {
        window.__saeDateWatcherBound = true;
        document.addEventListener('visibilitychange', () => {
          if (document.visibilityState === 'visible') syncActiveCalendarPeriod();
        });
        window.addEventListener('focus', syncActiveCalendarPeriod);
      }
    }

    function updateNetworkBannerUI() {
      const banner = document.getElementById('offline-warning-banner');
      if (!banner) return;
      banner.classList.toggle('hidden', navigator.onLine !== false);
    }

    function setupNetworkStatusWatcher() {
      updateNetworkBannerUI();
      if (window.__saeNetworkWatcherBound) return;
      window.__saeNetworkWatcherBound = true;

      window.addEventListener('offline', () => {
        updateNetworkBannerUI();
        showToast('Se perdió la conexión a internet. Modo respaldo local activo.', 'error');
      });

      window.addEventListener('online', () => {
        updateNetworkBannerUI();
        showToast('Conexión restablecida. Sincronizando registros pendientes...', 'success');
      });
    }

    function setupPasswordHoldReveal() {
  const passwordInput = document.getElementById('auth-password');
  const toggleBtn = document.getElementById('toggle-password-btn');
  const eyeShow = document.getElementById('eye-icon-show');
  const eyeHide = document.getElementById('eye-icon-hide');
  if (!passwordInput || !toggleBtn || !eyeShow || !eyeHide) return;

  const revealPassword = () => {
    passwordInput.type = 'text';
    toggleBtn.setAttribute('aria-pressed', 'true');
    eyeShow.classList.add('hidden');
    eyeHide.classList.remove('hidden');
  };

  const hidePassword = () => {
    passwordInput.type = 'password';
    toggleBtn.setAttribute('aria-pressed', 'false');
    eyeShow.classList.remove('hidden');
    eyeHide.classList.add('hidden');
  };

  toggleBtn.addEventListener('pointerdown', (e) => {
    e.preventDefault();
    revealPassword();
  });

  ['pointerup', 'pointerleave', 'pointercancel', 'blur'].forEach((eventName) => {
    toggleBtn.addEventListener(eventName, hidePassword);
  });

  toggleBtn.addEventListener('keydown', (e) => {
    if (e.key === ' ' || e.key === 'Enter') {
      e.preventDefault();
      revealPassword();
    }
  });

  toggleBtn.addEventListener('keyup', (e) => {
    if (e.key === ' ' || e.key === 'Enter') {
      hidePassword();
    }
  });

  window.addEventListener('pointerup', hidePassword);
}

// =========================================================================
// 14. ARRANQUE INICIAL DE LA APLICACIÓN
// =========================================================================
window.addEventListener('DOMContentLoaded', async () => {
  renderProcedureButtons();
  renderQuickObsTags();
  setupPasswordHoldReveal();
  setupNetworkStatusWatcher();

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
