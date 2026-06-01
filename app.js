// Bible Tracker - Logique Applicative (app.js)

// --- GESTION DE L'ÉTAT LOCAL (STATE) ---
let state = {
  bibleChapters: [],    // Chargé depuis Excel ou local: [{ id, book, bookId, chapter, ref, testament }]
  profiles: [],         // [{ id, username, avatarColor, readChapters, dailyGoal, notes, historyLog, todaysPlan, todaysPlanDate }]
  currentProfileId: null,
  
  // Variables simplifiées pointant sur le profil actif (pour compatibilité)
  readChapters: [],
  dailyGoal: 3,
  notes: [],
  historyLog: {},

  activeTab: 'dashboard-tab',
  currentTheme: 'dark',
  
  // Pagination pour les livres
  currentPage: 1,
  itemsPerPage: 18 // Livres par page en format cubes compacts
};

// Variables globales de contrôle
let readingChart = null;
let currentDrawerBookId = null;

// Association d'icônes symboliques pour chaque livre de la Bible (design épuré)
const BOOK_ICONS = {
  // Ancien Testament
  genese: 'fa-seedling',        // Création / Commencement
  exode: 'fa-water',           // Traversée de la Mer Rouge
  levitique: 'fa-fire-burner',  // Lois de sacrifice / Altar
  nombres: 'fa-list-ol',        // Dénombrement du peuple
  deuteronome: 'fa-book-open-reader', // Seconde loi donnée
  josue: 'fa-mountain',         // Conquête de Canaan / Murs de Jéricho
  juges: 'fa-gavel',            // Juges d'Israël
  ruth: 'fa-wheat-awn',         // Glanage de Ruth
  '1samuel': 'fa-crown',        // Royauté de David
  '2samuel': 'fa-crown',
  '1rois': 'fa-crown',
  '2rois': 'fa-crown',
  '1chroniques': 'fa-scroll',   // Chroniques historiques
  '2chroniques': 'fa-scroll',
  esdras: 'fa-scroll',          // Reconstruction du temple
  nehemie: 'fa-trowel-bricks',  // Reconstruction de la muraille
  esther: 'fa-gem',             // Reine Esther / Sceptre d'or
  job: 'fa-wind',               // Souffle de l'épreuve
  psaumes: 'fa-music',          // Louange / Harpe
  proverbes: 'fa-brain',        // Sagesse de Salomon
  ecclesiaste: 'fa-brain',
  cantique: 'fa-heart',         // Amour spirituel
  esaie: 'fa-bullhorn',         // Prophétie / Messie
  jeremie: 'fa-bullhorn',
  lamentations: 'fa-droplet',   // Pleurs de Jérémie
  ezechiel: 'fa-eye',           // Visions prophétiques
  daniel: 'fa-hourglass-half',  // Rêves et époques / Fosse aux lions
  osee: 'fa-bullhorn',
  joel: 'fa-bullhorn',
  amos: 'fa-bullhorn',
  abdias: 'fa-bullhorn',
  jonas: 'fa-fish',             // Jonas et le grand poisson
  michee: 'fa-bullhorn',
  nahum: 'fa-bullhorn',
  habacuc: 'fa-bullhorn',
  sophonie: 'fa-bullhorn',
  aggee: 'fa-bullhorn',
  zacharie: 'fa-bullhorn',
  malachie: 'fa-bullhorn',
  // Nouveau Testament
  matthieu: 'fa-cross',         // Évangile de Jésus
  marc: 'fa-cross',
  luc: 'fa-cross',
  jean: 'fa-cross',
  actes: 'fa-fire',             // Pentecôte / Saint-Esprit
  romains: 'fa-envelope-open-text', // Épîtres de Paul
  '1corinthiens': 'fa-envelope-open-text',
  '2corinthiens': 'fa-envelope-open-text',
  galates: 'fa-envelope-open-text',
  ephesiens: 'fa-envelope-open-text',
  philippiens: 'fa-envelope-open-text',
  colossiens: 'fa-envelope-open-text',
  '1thessaloniciens': 'fa-envelope-open-text',
  '2thessaloniciens': 'fa-envelope-open-text',
  '1timothee': 'fa-envelope-open-text',
  '2timothee': 'fa-envelope-open-text',
  tite: 'fa-envelope-open-text',
  philemon: 'fa-envelope-open-text',
  hebreux: 'fa-envelope-open-text',
  jacques: 'fa-envelope-open-text',
  '1pierre': 'fa-envelope-open-text',
  '2pierre': 'fa-envelope-open-text',
  '1jean': 'fa-envelope-open-text',
  '2jean': 'fa-envelope-open-text',
  '3jean': 'fa-envelope-open-text',
  jude: 'fa-envelope-open-text',
  apocalypse: 'fa-meteor'       // Révélation de la fin des temps
};

// Livres de l'Ancien Testament
const OT_BOOKS = [
  "Genèse", "Exode", "Lévitique", "Nombres", "Deutéronome", "Josué", "Juges", "Ruth",
  "1 Samuel", "2 Samuel", "1 Rois", "2 Rois", "1 Chroniques", "2 Chroniques", "Esdras",
  "Néhémie", "Esther", "Job", "Psaumes", "Proverbes", "Ecclésiaste", "Cantique des Cantiques",
  "Ésaïe", "Jérémie", "Lamentations", "Ézéchiel", "Daniel", "Osée", "Joël", "Amos",
  "Abdias", "Jonas", "Michée", "Nahum", "Habacuc", "Sophonie", "Aggée", "Zacharie", "Malachie"
];

// --- INITIALISATION DE L'APPLICATION ---
document.addEventListener('DOMContentLoaded', async () => {
  await initBibleData();
  loadLocalStorage();
  initTheme();
  initTabNavigation();
  initGoalControls();
  initProfilesUI();
  initNotesControls();
  initDrawerControls();
  
  // Vérifier si on doit afficher la modale obligatoire (bienvenue)
  checkWelcomeProfileState();

  if (state.currentProfileId) {
    populateBookSelect();
    renderReadingList();
    renderNotesList();
    updateDashboard();
  }
  
  // Ecouteurs sur les filtres de la liste de lecture
  document.querySelectorAll('.filter-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
      e.currentTarget.classList.add('active');
      state.currentPage = 1;
      renderReadingList();
    });
  });

  // Écouteur sur la recherche de livre
  document.getElementById('book-search-input').addEventListener('input', () => {
    state.currentPage = 1;
    renderReadingList();
  });
});

// --- CHARGEMENT DYNAMIQUE DU FICHIER EXCEL ---
async function initBibleData() {
  try {
    const response = await fetch('/liste_chapitres_bible.xlsx');
    if (!response.ok) throw new Error("Fichier introuvable.");
    
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet);
    
    if (rows.length === 0) throw new Error("Feuille Excel vide.");
    
    state.bibleChapters = rows.map(row => {
      const bookName = row['Livre'] || '';
      const chapter = parseInt(row['Chapitre'], 10) || 1;
      const ref = row['Référence'] || `${bookName} ${chapter}`;
      const bookId = slugify(bookName);
      
      return {
        id: `${bookId}-${chapter}`,
        book: bookName,
        bookId: bookId,
        chapter: chapter,
        ref: ref,
        testament: getTestamentByBookName(bookName)
      };
    });
    console.log("Bible chargée depuis Excel :", state.bibleChapters.length, "chapitres.");
  } catch (error) {
    console.warn("Échec du chargement de l'Excel. Utilisation des données locales de secours.", error);
    
    // Secours : Générer à partir de bibleData.js
    state.bibleChapters = [];
    BIBLE_DATA.forEach(book => {
      for (let c = 1; c <= book.chapters; c++) {
        state.bibleChapters.push({
          id: `${book.id}-${c}`,
          book: book.name,
          bookId: book.id,
          chapter: c,
          ref: `${book.name} ${c}`,
          testament: book.testament
        });
      }
    });
  }
}

// Helpers
function getTestamentByBookName(bookName) {
  return OT_BOOKS.includes(bookName) ? 'Ancien' : 'Nouveau';
}

function slugify(text) {
  return text.toString().toLowerCase().trim()
    .replace(/\s+/g, '')
    .replace(/[éèêë]/g, 'e')
    .replace(/[àâä]/g, 'a')
    .replace(/[îï]/g, 'i')
    .replace(/[ôö]/g, 'o')
    .replace(/[ûü]/g, 'u')
    .replace(/[ç]/g, 'c')
    .replace(/[^a-z0-9]/g, '');
}

// --- GESTION DES PROFILS (MULTI-UTILISATEURS) ---
function loadLocalStorage() {
  const savedProfiles = localStorage.getItem('bible_tracker_profiles');
  const savedCurrentProfileId = localStorage.getItem('bible_tracker_current_profile_id');
  
  // Détecter et nettoyer l'ancien état de test s'il existe
  if (savedProfiles && (savedProfiles.includes("prof-1") || savedProfiles.includes("prof-2"))) {
    localStorage.removeItem('bible_tracker_profiles');
    localStorage.removeItem('bible_tracker_current_profile_id');
    localStorage.removeItem('bible_tracker_read_chapters');
    localStorage.removeItem('bible_tracker_history');
    window.location.reload();
    return;
  }
  
  if (savedProfiles) {
    state.profiles = JSON.parse(savedProfiles);
  } else {
    state.profiles = [];
  }

  if (state.profiles.length === 0) {
    state.currentProfileId = null;
    localStorage.removeItem('bible_tracker_current_profile_id');
  } else {
    if (savedCurrentProfileId && state.profiles.some(p => p.id === savedCurrentProfileId)) {
      state.currentProfileId = savedCurrentProfileId;
    } else {
      state.currentProfileId = state.profiles[0].id;
      localStorage.setItem('bible_tracker_current_profile_id', state.currentProfileId);
    }
    loadActiveProfileState();
  }
}

function loadActiveProfileState() {
  const activeProfile = state.profiles.find(p => p.id === state.currentProfileId);
  if (!activeProfile) return;

  state.readChapters = activeProfile.readChapters || [];
  state.dailyGoal = activeProfile.dailyGoal || 3;
  state.notes = activeProfile.notes || [];
  state.historyLog = activeProfile.historyLog || {};
  
  document.getElementById('daily-goal-input').value = state.dailyGoal;
  
  updateActiveProfileHeader();
}

function saveActiveProfileState() {
  const activeProfile = state.profiles.find(p => p.id === state.currentProfileId);
  if (activeProfile) {
    activeProfile.readChapters = state.readChapters;
    activeProfile.dailyGoal = state.dailyGoal;
    activeProfile.notes = state.notes;
    activeProfile.historyLog = state.historyLog;
  }
  saveProfilesToStorage();
}

function saveProfilesToStorage() {
  localStorage.setItem('bible_tracker_profiles', JSON.stringify(state.profiles));
}

function updateActiveProfileHeader() {
  const activeProfile = state.profiles.find(p => p.id === state.currentProfileId);
  if (!activeProfile) return;

  const headerAvatar = document.getElementById('header-avatar');
  const headerUsername = document.getElementById('header-username');
  
  headerAvatar.textContent = activeProfile.username.charAt(0);
  headerAvatar.style.backgroundColor = activeProfile.avatarColor;
  headerUsername.textContent = activeProfile.username;
}

function initProfilesUI() {
  const activeProfileBtn = document.getElementById('active-profile-btn');
  const dropdownList = document.getElementById('profile-dropdown-list');
  const profileModal = document.getElementById('profile-modal');
  const cancelModalBtn = document.getElementById('close-profile-modal-btn');
  const saveProfileBtn = document.getElementById('save-profile-btn');
  const newProfileUsername = document.getElementById('new-profile-username');
  
  // Sélection de couleur dans la modale
  let selectedColor = '#6366f1';
  document.querySelectorAll('#modal-color-picker .color-dot').forEach(dot => {
    dot.addEventListener('click', (e) => {
      document.querySelectorAll('#modal-color-picker .color-dot').forEach(d => d.classList.remove('active'));
      e.currentTarget.classList.add('active');
      selectedColor = e.currentTarget.getAttribute('data-color');
    });
  });

  // Toggle Dropdown
  activeProfileBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    if (state.profiles.length > 0) {
      dropdownList.classList.toggle('active');
      populateProfileDropdown();
    }
  });

  // Fermer le dropdown en cliquant à côté
  document.addEventListener('click', () => {
    dropdownList.classList.remove('active');
  });

  // Fermer la modale (uniquement si des profils existent)
  cancelModalBtn.addEventListener('click', () => {
    if (state.profiles.length > 0) {
      profileModal.classList.remove('active');
    }
  });

  // Empêcher la fermeture en cliquant à côté en mode Bienvenue
  profileModal.addEventListener('click', (e) => {
    if (e.target === profileModal && state.profiles.length > 0) {
      profileModal.classList.remove('active');
    }
  });

  // Sauvegarder le nouveau profil
  saveProfileBtn.addEventListener('click', () => {
    const username = newProfileUsername.value.trim();
    if (!username) {
      alert("Veuillez saisir un pseudo.");
      return;
    }

    const isFirstProfile = state.profiles.length === 0;

    const newProfile = {
      id: 'prof-' + Date.now(),
      username: username,
      avatarColor: selectedColor,
      readChapters: [],
      dailyGoal: 3,
      notes: isFirstProfile ? [
        {
          id: 'note-welcome',
          title: 'Bienvenue dans votre journal !',
          bookId: '',
          content: 'Ceci est votre espace de notes personnel. Vous pouvez y inscrire vos réflexions de lecture quotidiennes, vos versets préférés ou vos prières. Bon voyage à travers les Écritures !',
          date: getFormattedDate(new Date())
        }
      ] : [],
      historyLog: {},
      todaysPlan: [],
      todaysPlanDate: ''
    };

    state.profiles.push(newProfile);
    saveProfilesToStorage();
    
    // Fermer la modale
    profileModal.classList.remove('active');
    newProfileUsername.value = '';
    
    // Mettre à jour l'état du bouton d'annulation de la modale
    checkWelcomeProfileState();
    
    // Switch sur ce nouveau profil
    switchProfile(newProfile.id);
  });
}

function populateProfileDropdown() {
  const dropdownList = document.getElementById('profile-dropdown-list');
  dropdownList.innerHTML = '';

  state.profiles.forEach(p => {
    const item = document.createElement('button');
    item.className = 'profile-dropdown-item';
    if (p.id === state.currentProfileId) {
      item.classList.add('active-item');
    }

    item.innerHTML = `
      <span class="avatar-circle" style="background-color: ${p.avatarColor}">${p.username.charAt(0)}</span>
      <span>${p.username}</span>
    `;

    item.addEventListener('click', () => {
      switchProfile(p.id);
    });

    dropdownList.appendChild(item);
  });

  // Option d'ajout de profil
  const addBtn = document.createElement('button');
  addBtn.className = 'profile-dropdown-item add-profile-btn';
  addBtn.innerHTML = `
    <i class="fa-solid fa-user-plus"></i>
    <span>Créer un profil</span>
  `;

  addBtn.addEventListener('click', () => {
    document.getElementById('profile-modal').classList.add('active');
  });

  dropdownList.appendChild(addBtn);
}

function switchProfile(profileId) {
  saveActiveProfileState();
  state.currentProfileId = profileId;
  localStorage.setItem('bible_tracker_current_profile_id', profileId);
  loadActiveProfileState();

  updateDashboard();
  renderReadingList();
  renderNotesList();
  resetNoteForm();
}

// --- THÈME CLAIR/SOMBRE ---
function initTheme() {
  const savedTheme = localStorage.getItem('bible_tracker_theme') || 'dark';
  state.currentTheme = savedTheme;
  document.documentElement.setAttribute('data-theme', savedTheme);
  
  const themeIcon = document.getElementById('theme-icon');
  themeIcon.className = savedTheme === 'light' ? 'fa-solid fa-sun' : 'fa-solid fa-moon';

  document.getElementById('theme-toggle-btn').addEventListener('click', () => {
    const newTheme = state.currentTheme === 'dark' ? 'light' : 'dark';
    state.currentTheme = newTheme;
    document.documentElement.setAttribute('data-theme', newTheme);
    localStorage.setItem('bible_tracker_theme', newTheme);
    themeIcon.className = newTheme === 'light' ? 'fa-solid fa-sun' : 'fa-solid fa-moon';
    
    if (readingChart) {
      updateChartTheme();
    }
  });
}

// --- NAVIGATION ---
function initTabNavigation() {
  const tabs = document.querySelectorAll('.tab-btn');
  tabs.forEach(tab => {
    tab.addEventListener('click', (e) => {
      const targetTab = e.currentTarget.getAttribute('data-tab');
      
      tabs.forEach(t => t.classList.remove('active'));
      e.currentTarget.classList.add('active');
      
      document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
      });
      document.getElementById(targetTab).classList.add('active');
      
      state.activeTab = targetTab;
      
      if (targetTab === 'dashboard-tab') {
        updateDashboard();
      } else if (targetTab === 'reading-list-tab') {
        renderReadingList();
      }
    });
  });
}

// --- RÉGLAGE DE L'OBJECTIF QUOTIDIEN ---
function initGoalControls() {
  const input = document.getElementById('daily-goal-input');
  const minusBtn = document.getElementById('goal-minus-btn');
  const plusBtn = document.getElementById('goal-plus-btn');

  const updateGoal = (newVal) => {
    if (newVal < 1) newVal = 1;
    if (newVal > 150) newVal = 150;
    state.dailyGoal = newVal;
    input.value = newVal;
    
    // Forcer la régénération du plan du jour
    const activeProfile = state.profiles.find(p => p.id === state.currentProfileId);
    if (activeProfile) {
      activeProfile.todaysPlan = [];
      activeProfile.todaysPlanDate = '';
    }
    
    saveActiveProfileState();
    renderReadingList();
    updateDashboard();
  };

  minusBtn.addEventListener('click', () => updateGoal(state.dailyGoal - 1));
  plusBtn.addEventListener('click', () => updateGoal(state.dailyGoal + 1));
  input.addEventListener('change', (e) => updateGoal(parseInt(e.target.value, 10) || 3));
}

// --- POPULER LES LIVRES DANS LE JOURNAL ---
function populateBookSelect() {
  if (!state.currentProfileId) return;
  const select = document.getElementById('note-book-select');
  select.innerHTML = '<option value="">Général / Aucun</option>';
  
  const uniqueBooks = [];
  const seenBooks = new Set();
  
  state.bibleChapters.forEach(ch => {
    if (!seenBooks.has(ch.bookId)) {
      seenBooks.add(ch.bookId);
      uniqueBooks.push({ id: ch.bookId, name: ch.book });
    }
  });

  uniqueBooks.forEach(book => {
    const option = document.createElement('option');
    option.value = book.id;
    option.textContent = book.name;
    select.appendChild(option);
  });
}

// --- PLAN DE LECTURE DU JOUR FIXE (SANS GLISSEMENT DYNAMIQUE) ---
function getTodaysReadingPlan() {
  const activeProfile = state.profiles.find(p => p.id === state.currentProfileId);
  if (!activeProfile) return [];

  const todayKey = formatDateKey(new Date());

  // Si le plan a déjà été généré aujourd'hui pour la taille souhaitée, on le fige
  if (activeProfile.todaysPlanDate === todayKey && activeProfile.todaysPlan && activeProfile.todaysPlan.length > 0) {
    if (activeProfile.todaysPlan.length === state.dailyGoal) {
      return activeProfile.todaysPlan;
    }
  }

  // Sinon, on le génère à partir des premiers chapitres non lus
  const plan = [];
  let count = 0;
  
  for (const ch of state.bibleChapters) {
    if (!state.readChapters.includes(ch.id)) {
      plan.push(ch.id);
      count++;
      if (count >= state.dailyGoal) {
        break;
      }
    }
  }
  
  activeProfile.todaysPlan = plan;
  activeProfile.todaysPlanDate = todayKey;
  saveProfilesToStorage();

  return plan;
}

// --- TIROIR LATÉRAL DES CHAPITRES (DRAWER) ---
function initDrawerControls() {
  const closeBtn = document.getElementById('close-drawer-btn');
  const drawerOverlay = document.getElementById('chapters-drawer');

  closeBtn.addEventListener('click', () => {
    drawerOverlay.classList.remove('active');
    currentDrawerBookId = null;
  });

  // Fermer en cliquant sur l'overlay
  drawerOverlay.addEventListener('click', (e) => {
    if (e.target === drawerOverlay) {
      drawerOverlay.classList.remove('active');
      currentDrawerBookId = null;
    }
  });
}

function openBookChaptersDrawer(bookId) {
  const bookChapters = state.bibleChapters.filter(ch => ch.bookId === bookId);
  if (bookChapters.length === 0) return;

  const bookName = bookChapters[0].book;
  currentDrawerBookId = bookId;

  document.getElementById('drawer-book-title').textContent = bookName;
  document.getElementById('chapters-drawer').classList.add('active');

  renderDrawerChaptersGrid(bookId, bookChapters);
}

function renderDrawerChaptersGrid(bookId, chapters) {
  const listContainer = document.getElementById('drawer-chapters-list');
  listContainer.innerHTML = '';

  const todaysPlan = getTodaysReadingPlan();
  const readChaptersInBook = chapters.filter(ch => state.readChapters.includes(ch.id)).length;
  document.getElementById('drawer-book-progress').textContent = `${readChaptersInBook} / ${chapters.length} chapitres lus`;

  chapters.forEach(ch => {
    const isRead = state.readChapters.includes(ch.id);
    const isHighlighted = todaysPlan.includes(ch.id);

    const btn = document.createElement('button');
    btn.className = 'chapter-btn';
    btn.textContent = ch.chapter;
    btn.setAttribute('aria-label', `Chapitre ${ch.chapter}`);

    if (isRead) btn.classList.add('checked');
    if (isHighlighted && !isRead) btn.classList.add('highlighted');

    btn.addEventListener('click', () => {
      toggleChapterState(ch.id, btn);
      // Mettre à jour les indicateurs du drawer
      const updatedRead = chapters.filter(c => state.readChapters.includes(c.id)).length;
      document.getElementById('drawer-book-progress').textContent = `${updatedRead} / ${chapters.length} chapitres lus`;
    });

    listContainer.appendChild(btn);
  });
}

function toggleChapterState(chapterId, btnElement) {
  const index = state.readChapters.indexOf(chapterId);
  const todayKey = formatDateKey(new Date());

  if (index === -1) {
    state.readChapters.push(chapterId);
    btnElement.classList.remove('highlighted');
    btnElement.classList.add('checked');
    state.historyLog[todayKey] = (state.historyLog[todayKey] || 0) + 1;
  } else {
    state.readChapters.splice(index, 1);
    btnElement.classList.remove('checked');
    const todaysPlan = getTodaysReadingPlan();
    if (todaysPlan.includes(chapterId)) {
      btnElement.classList.add('highlighted');
    }
    if (state.historyLog[todayKey] && state.historyLog[todayKey] > 0) {
      state.historyLog[todayKey]--;
    }
  }

  saveActiveProfileState();

  // Actualiser la grille des cubes et de la surbrillance sans fermer le tiroir
  renderReadingListCubesOnly();
  renderTodayCubes();
  
  if (state.activeTab === 'dashboard-tab') {
    updateDashboard();
  }
}

// --- RENDU DE LA LISTE DE LECTURE (GRILLE DE CUBES ET LECTURE DU JOUR) ---
function renderReadingList() {
  if (!state.currentProfileId) return;
  renderTodayCubes();
  renderReadingListCubesOnly();
}

function renderTodayCubes() {
  const section = document.getElementById('today-readings-section');
  const gridContainer = document.getElementById('today-cubes-list');
  const progressRatioSpan = document.getElementById('today-progress-ratio');

  gridContainer.innerHTML = '';
  const todaysPlan = getTodaysReadingPlan();

  if (todaysPlan.length === 0) {
    section.style.display = 'none';
    return;
  }

  section.style.display = 'block';

  let readCount = 0;

  todaysPlan.forEach(chapterId => {
    const ch = state.bibleChapters.find(c => c.id === chapterId);
    if (!ch) return;

    const isRead = state.readChapters.includes(chapterId);
    if (isRead) readCount++;

    const cube = document.createElement('div');
    cube.className = 'today-chapter-cube';
    if (isRead) cube.classList.add('checked');

    cube.innerHTML = `
      <span class="book-name">${ch.book}</span>
      <span class="chapter-num">${ch.chapter}</span>
      <div class="read-check-icon">
        <i class="fa-solid fa-check"></i>
      </div>
    `;

    // Clic pour cocher/décocher directement le cube du jour
    cube.addEventListener('click', (e) => {
      e.stopPropagation();
      toggleTodayChapter(chapterId, cube);
    });

    gridContainer.appendChild(cube);
  });

  progressRatioSpan.textContent = `${readCount} / ${todaysPlan.length} chapitres`;
}

function toggleTodayChapter(chapterId, cubeElement) {
  const index = state.readChapters.indexOf(chapterId);
  const todayKey = formatDateKey(new Date());

  if (index === -1) {
    state.readChapters.push(chapterId);
    cubeElement.classList.add('checked');
    state.historyLog[todayKey] = (state.historyLog[todayKey] || 0) + 1;
  } else {
    state.readChapters.splice(index, 1);
    cubeElement.classList.remove('checked');
    if (state.historyLog[todayKey] && state.historyLog[todayKey] > 0) {
      state.historyLog[todayKey]--;
    }
  }

  saveActiveProfileState();
  renderTodayCubes();
  renderReadingListCubesOnly();

  // Si le tiroir est actuellement ouvert sur le livre de ce chapitre, on met à jour le tiroir
  if (currentDrawerBookId) {
    const ch = state.bibleChapters.find(c => c.id === chapterId);
    if (ch && ch.bookId === currentDrawerBookId) {
      const bookChapters = state.bibleChapters.filter(c => c.bookId === currentDrawerBookId);
      renderDrawerChaptersGrid(currentDrawerBookId, bookChapters);
    }
  }

  if (state.activeTab === 'dashboard-tab') {
    updateDashboard();
  }
}

function renderReadingListCubesOnly() {
  const container = document.getElementById('books-list-container');
  const paginationContainer = document.getElementById('pagination-container');
  const searchQuery = document.getElementById('book-search-input').value.toLowerCase().trim();
  const activeFilter = document.querySelector('.filter-btn.active').getAttribute('data-filter');

  container.innerHTML = '';
  paginationContainer.innerHTML = '';

  const todaysPlan = getTodaysReadingPlan();

  // 1. Grouper les chapitres par livre pour filtrer et afficher les livres en tant que cubes
  const booksMap = {};
  state.bibleChapters.forEach(ch => {
    if (!booksMap[ch.bookId]) {
      booksMap[ch.bookId] = {
        id: ch.bookId,
        name: ch.book,
        testament: ch.testament,
        chapters: []
      };
    }
    booksMap[ch.bookId].chapters.push(ch);
  });

  const allBooks = Object.values(booksMap);

  // 2. Filtrer les livres
  const filteredBooks = allBooks.filter(book => {
    const matchesSearch = book.name.toLowerCase().includes(searchQuery);
    if (!matchesSearch) return false;

    const readCount = book.chapters.filter(ch => state.readChapters.includes(ch.id)).length;
    const isCompleted = readCount === book.chapters.length;

    if (activeFilter === 'read') return isCompleted;
    if (activeFilter === 'unread') return !isCompleted;
    if (activeFilter === 'today') {
      return book.chapters.some(ch => todaysPlan.includes(ch.id));
    }
    return true;
  });

  if (filteredBooks.length === 0) {
    container.innerHTML = `
      <div class="card empty-state" style="grid-column: 1 / -1; width: 100%;">
        <i class="fa-solid fa-magnifying-glass"></i>
        <h3>Aucun livre trouvé</h3>
        <p>Ajustez vos filtres ou la barre de recherche.</p>
      </div>
    `;
    return;
  }

  // 3. Paginer les livres (sauf filtre "Aujourd'hui")
  let displayBooks = filteredBooks;
  const isFilterToday = activeFilter === 'today';

  if (!isFilterToday) {
    const totalItems = filteredBooks.length;
    const totalPages = Math.ceil(totalItems / state.itemsPerPage);

    if (state.currentPage > totalPages) {
      state.currentPage = 1;
    }

    const startIndex = (state.currentPage - 1) * state.itemsPerPage;
    const endIndex = startIndex + state.itemsPerPage;
    displayBooks = filteredBooks.slice(startIndex, endIndex);

    if (totalPages > 1) {
      paginationContainer.innerHTML = `
        <button class="pagination-btn" id="prev-page-btn" ${state.currentPage === 1 ? 'disabled' : ''}>
          <i class="fa-solid fa-chevron-left"></i>
        </button>
        <span class="pagination-info">Page ${state.currentPage} sur ${totalPages} (${totalItems} livres)</span>
        <button class="pagination-btn" id="next-page-btn" ${state.currentPage === totalPages ? 'disabled' : ''}>
          <i class="fa-solid fa-chevron-right"></i>
        </button>
      `;

      document.getElementById('prev-page-btn').addEventListener('click', () => {
        state.currentPage--;
        renderReadingListCubesOnly();
        container.scrollIntoView({ behavior: 'smooth' });
      });

      document.getElementById('next-page-btn').addEventListener('click', () => {
        state.currentPage++;
        renderReadingListCubesOnly();
        container.scrollIntoView({ behavior: 'smooth' });
      });
    }
  }

  // 4. Rendre les cubes de livres
  displayBooks.forEach(book => {
    const readCount = book.chapters.filter(ch => state.readChapters.includes(ch.id)).length;
    const isCompleted = readCount === book.chapters.length;
    const hasToday = book.chapters.some(ch => todaysPlan.includes(ch.id));
    const theme = getBookTheme(book.id);

    const cube = document.createElement('div');
    cube.className = 'book-cube';
    cube.style.borderColor = theme.border;
    
    if (isCompleted) {
      cube.classList.add('completed');
      cube.style.background = 'rgba(16, 185, 129, 0.02)';
      cube.style.borderColor = 'rgba(16, 185, 129, 0.3)';
    } else if (hasToday) {
      cube.classList.add('has-highlighted');
      cube.style.borderColor = 'var(--border-highlight)';
    }

    const iconClass = BOOK_ICONS[book.id] || 'fa-book';

    cube.innerHTML = `
      <div class="book-cube-icon-wrapper" style="background-color: ${theme.bg}; color: ${theme.color}; border-color: ${theme.border}">
        <i class="fa-solid ${iconClass}"></i>
      </div>
      <span class="book-cube-name" style="${hasToday ? 'color: var(--color-accent);' : ''}">${book.name}</span>
      <span class="book-cube-progress" style="color: ${theme.color}">${readCount} / ${book.chapters.length}</span>
    `;

    // Ouvrir le tiroir des chapitres au clic
    cube.addEventListener('click', () => {
      openBookChaptersDrawer(book.id);
    });

    container.appendChild(cube);
  });
}

// --- TABLEAU DE BORD (COMPACT & COMPARATIF MEMBRES) ---
function updateDashboard() {
  if (!state.currentProfileId) return;
  if (state.bibleChapters.length === 0) return;

  // Enregistrer l'état du profil actif d'abord pour avoir des données à jour
  saveActiveProfileState();

  const totalChapters = state.bibleChapters.length;
  const readCount = state.readChapters.length;
  const percentage = Math.round((readCount / totalChapters) * 100);
  
  document.getElementById('stat-chapters-ratio').textContent = `${readCount} / ${totalChapters}`;
  document.getElementById('progress-percentage-val').textContent = `${percentage}%`;
  
  // Regrouper les chapitres par livre pour les calculs de livres complétés
  const chaptersByBook = {};
  state.bibleChapters.forEach(ch => {
    if (!chaptersByBook[ch.bookId]) {
      chaptersByBook[ch.bookId] = {
        name: ch.book,
        testament: ch.testament,
        chapters: []
      };
    }
    chaptersByBook[ch.bookId].chapters.push(ch.id);
  });

  let completedBooks = 0;
  let completedOT = 0;
  let completedNT = 0;
  let otReadChapters = 0;
  let ntReadChapters = 0;
  let otTotalChapters = 0;
  let ntTotalChapters = 0;

  Object.keys(chaptersByBook).forEach(bookId => {
    const bookData = chaptersByBook[bookId];
    const readInBook = bookData.chapters.filter(id => state.readChapters.includes(id)).length;
    const isCompleted = readInBook === bookData.chapters.length;

    if (bookData.testament === 'Ancien') {
      otTotalChapters += bookData.chapters.length;
      otReadChapters += readInBook;
      if (isCompleted) completedOT++;
    } else {
      ntTotalChapters += bookData.chapters.length;
      ntReadChapters += readInBook;
      if (isCompleted) completedNT++;
    }

    if (isCompleted) completedBooks++;
  });

  const totalBooks = Object.keys(chaptersByBook).length;
  const totalOTBooks = Object.values(chaptersByBook).filter(b => b.testament === 'Ancien').length;
  const totalNTBooks = Object.values(chaptersByBook).filter(b => b.testament === 'Nouveau').length;

  document.getElementById('stat-books-ratio').textContent = `${completedBooks} / ${totalBooks}`;
  document.getElementById('stat-ot-ratio').textContent = `${completedOT} / ${totalOTBooks}`;
  document.getElementById('stat-nt-ratio').textContent = `${completedNT} / ${totalNTBooks}`;

  // Résumé textuel
  const summaryTxt = document.getElementById('progress-summary-txt');
  if (readCount === 0) {
    summaryTxt.textContent = "Aucun chapitre lu. Ouvrez l'onglet des lectures pour commencer !";
  } else if (readCount === totalChapters) {
    summaryTxt.textContent = "Félicitations ! Vous avez complété toute la liste de lecture ! 🎉";
  } else {
    const dailyGoalDays = Math.ceil((totalChapters - readCount) / state.dailyGoal);
    summaryTxt.textContent = `En lisant ${state.dailyGoal} chapitres/jour, vous terminerez dans environ ${dailyGoalDays} jours.`;
  }

  // Mettre à jour le cercle de progression
  const circle = document.getElementById('circular-progress-bar');
  const circumference = 2 * Math.PI * 70;
  circle.style.strokeDasharray = circumference;
  const offset = circumference - (percentage / 100) * circumference;
  circle.style.strokeDashoffset = offset;

  // Calculs par testament
  const otPct = otTotalChapters > 0 ? Math.round((otReadChapters / otTotalChapters) * 100) : 0;
  const ntPct = ntTotalChapters > 0 ? Math.round((ntReadChapters / ntTotalChapters) * 100) : 0;

  document.getElementById('ot-pct-val').textContent = `${otPct}%`;
  document.getElementById('ot-progress-fill').style.width = `${otPct}%`;
  document.getElementById('ot-chapters-data').textContent = `${otReadChapters} / ${otTotalChapters}`;
  document.getElementById('ot-books-remaining').textContent = `${totalOTBooks - completedOT} livres restants`;

  document.getElementById('nt-pct-val').textContent = `${ntPct}%`;
  document.getElementById('nt-progress-fill').style.width = `${ntPct}%`;
  document.getElementById('nt-chapters-data').textContent = `${ntReadChapters} / ${ntTotalChapters}`;
  document.getElementById('nt-books-remaining').textContent = `${totalNTBooks - completedNT} livres restants`;

  // Mettre à jour le tableau comparatif des membres (Compagnons)
  renderCompanionsList();

  renderOrUpdateChart();
}

function renderCompanionsList() {
  const container = document.getElementById('companions-list-container');
  container.innerHTML = '';

  const totalChapters = state.bibleChapters.length;
  if (totalChapters === 0) return;

  state.profiles.forEach(p => {
    // Calculer le pourcentage de progression pour ce profil
    const count = p.readChapters ? p.readChapters.length : 0;
    const pct = Math.round((count / totalChapters) * 100);

    const card = document.createElement('div');
    card.className = 'companion-card';
    if (p.id === state.currentProfileId) {
      card.style.borderColor = 'var(--color-primary)';
      card.style.background = 'rgba(99, 102, 241, 0.05)';
    }

    card.innerHTML = `
      <div class="companion-avatar" style="background-color: ${p.avatarColor}">
        ${p.username.charAt(0)}
      </div>
      <div class="companion-info">
        <div class="companion-name">
          ${p.username} ${p.id === state.currentProfileId ? '<span style="font-size:0.75rem; color:var(--color-primary); font-weight:600;">(Vous)</span>' : ''}
        </div>
        <div class="companion-progress">
          ${count} / ${totalChapters} chapitres (${pct}%)
        </div>
        <div class="companion-progress-bar">
          <div class="companion-progress-fill" style="width: ${pct}%; background-color: ${p.avatarColor};"></div>
        </div>
      </div>
    `;

    // Clic pour basculer sur ce profil directement depuis la liste
    card.addEventListener('click', () => {
      switchProfile(p.id);
    });

    container.appendChild(card);
  });
}

// --- LOGIQUE DES GRAPHIQUES (CHART.JS) ---
function renderOrUpdateChart() {
  const chartCanvas = document.getElementById('readingHistoryChart');
  if (!chartCanvas) return;
  
  const ctx = chartCanvas.getContext('2d');
  
  const labels = [];
  const data = [];
  const dateKeys = [];
  
  for (let i = 6; i >= 0; i--) {
    const d = new Date();
    d.setDate(d.getDate() - i);
    dateKeys.push(formatDateKey(d));
    labels.push(d.toLocaleDateString('fr-FR', { weekday: 'short', day: 'numeric' }));
  }

  dateKeys.forEach(key => {
    data.push(state.historyLog[key] || 0);
  });

  const primaryColor = getComputedStyle(document.documentElement).getPropertyValue('--color-primary').trim() || '#6366f1';
  const textColor = getComputedStyle(document.documentElement).getPropertyValue('--text-secondary').trim() || '#9ca3af';
  const gridColor = getComputedStyle(document.documentElement).getPropertyValue('--border-color').trim() || 'rgba(255,255,255,0.08)';

  if (readingChart) {
    readingChart.data.labels = labels;
    readingChart.data.datasets[0].data = data;
    readingChart.data.datasets[0].borderColor = primaryColor;
    readingChart.data.datasets[0].backgroundColor = createChartGradient(ctx, primaryColor);
    readingChart.options.scales.x.ticks.color = textColor;
    readingChart.options.scales.y.ticks.color = textColor;
    readingChart.options.scales.x.grid.color = gridColor;
    readingChart.options.scales.y.grid.color = gridColor;
    readingChart.update();
  } else {
    readingChart = new Chart(ctx, {
      type: 'line',
      data: {
        labels: labels,
        datasets: [{
          label: 'Chapitres lus',
          data: data,
          borderColor: primaryColor,
          borderWidth: 3,
          pointBackgroundColor: primaryColor,
          pointHoverRadius: 7,
          pointHoverBackgroundColor: '#ffffff',
          tension: 0.4,
          fill: true,
          backgroundColor: createChartGradient(ctx, primaryColor)
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            mode: 'index',
            intersect: false,
            backgroundColor: 'rgba(15, 23, 42, 0.9)',
            titleFont: { family: 'Plus Jakarta Sans', weight: 'bold' },
            bodyFont: { family: 'Plus Jakarta Sans' },
            padding: 12,
            cornerRadius: 10
          }
        },
        scales: {
          x: {
            grid: { color: gridColor, drawBorder: false },
            ticks: { color: textColor, font: { family: 'Plus Jakarta Sans', size: 11 } }
          },
          y: {
            grid: { color: gridColor, drawBorder: false },
            ticks: { 
              color: textColor, 
              font: { family: 'Plus Jakarta Sans', size: 11 },
              stepSize: 1,
              precision: 0
            },
            min: 0
          }
        }
      }
    });
  }
}

function createChartGradient(ctx, color) {
  const gradient = ctx.createLinearGradient(0, 0, 0, 200);
  gradient.addColorStop(0, color + '44');
  gradient.addColorStop(1, color + '00');
  return gradient;
}

function updateChartTheme() {
  const primaryColor = getComputedStyle(document.documentElement).getPropertyValue('--color-primary').trim() || '#6366f1';
  const textColor = getComputedStyle(document.documentElement).getPropertyValue('--text-secondary').trim() || '#9ca3af';
  const gridColor = getComputedStyle(document.documentElement).getPropertyValue('--border-color').trim() || 'rgba(255,255,255,0.08)';
  const ctx = document.getElementById('readingHistoryChart').getContext('2d');

  readingChart.data.datasets[0].borderColor = primaryColor;
  readingChart.data.datasets[0].pointBackgroundColor = primaryColor;
  readingChart.data.datasets[0].backgroundColor = createChartGradient(ctx, primaryColor);
  readingChart.options.scales.x.ticks.color = textColor;
  readingChart.options.scales.y.ticks.color = textColor;
  readingChart.options.scales.x.grid.color = gridColor;
  readingChart.options.scales.y.grid.color = gridColor;
  readingChart.update();
}

// --- GESTION DES NOTES ---
function initNotesControls() {
  const newNoteBtn = document.getElementById('new-note-btn');
  const saveNoteBtn = document.getElementById('save-note-btn');
  const deleteNoteBtn = document.getElementById('delete-note-btn');
  const cancelNoteBtn = document.getElementById('cancel-note-btn');
  const searchInput = document.getElementById('notes-search-input');

  newNoteBtn.addEventListener('click', () => {
    resetNoteForm();
  });

  saveNoteBtn.addEventListener('click', () => {
    const title = document.getElementById('note-title-input').value.trim();
    const bookId = document.getElementById('note-book-select').value;
    const content = document.getElementById('note-content-textarea').value.trim();

    if (!title) {
      alert("Veuillez saisir un titre pour votre note.");
      return;
    }

    if (state.currentNoteId) {
      const note = state.notes.find(n => n.id === state.currentNoteId);
      if (note) {
        note.title = title;
        note.bookId = bookId;
        note.content = content;
        note.date = getFormattedDate(new Date());
      }
    } else {
      const newNote = {
        id: 'note-' + Date.now(),
        title: title,
        bookId: bookId,
        content: content,
        date: getFormattedDate(new Date())
      };
      state.notes.unshift(newNote);
      state.currentNoteId = newNote.id;
    }

    saveActiveProfileState();
    renderNotesList();
    displayNoteInForm(state.currentNoteId);
    
    saveNoteBtn.innerHTML = '<i class="fa-solid fa-circle-check"></i> Enregistré !';
    setTimeout(() => {
      saveNoteBtn.innerHTML = '<i class="fa-solid fa-floppy-disk"></i> Enregistrer la Note';
    }, 1500);
  });

  deleteNoteBtn.addEventListener('click', () => {
    if (!state.currentNoteId) return;
    
    if (confirm("Êtes-vous sûr de vouloir supprimer cette note ?")) {
      state.notes = state.notes.filter(n => n.id !== state.currentNoteId);
      saveActiveProfileState();
      renderNotesList();
      resetNoteForm();
    }
  });

  cancelNoteBtn.addEventListener('click', () => {
    if (state.currentNoteId) {
      displayNoteInForm(state.currentNoteId);
    } else {
      resetNoteForm();
    }
  });

  searchInput.addEventListener('input', () => {
    renderNotesList();
  });

  if (state.notes.length > 0) {
    displayNoteInForm(state.notes[0].id);
  }
}

function renderNotesList() {
  if (!state.currentProfileId) return;
  const listContainer = document.getElementById('notes-list-items');
  const searchQuery = document.getElementById('notes-search-input').value.toLowerCase().trim();
  
  listContainer.innerHTML = '';
  
  const filteredNotes = state.notes.filter(note => {
    const bookName = getBookNameById(note.bookId).toLowerCase();
    return (
      note.title.toLowerCase().includes(searchQuery) ||
      note.content.toLowerCase().includes(searchQuery) ||
      bookName.includes(searchQuery)
    );
  });

  if (filteredNotes.length === 0) {
    listContainer.innerHTML = `
      <div style="text-align: center; color: var(--text-muted); padding: 2rem 0; font-size: 0.85rem;">
        Aucune note correspondante.
      </div>
    `;
    return;
  }

  filteredNotes.forEach(note => {
    const item = document.createElement('button');
    item.className = 'note-item-card';
    if (note.id === state.currentNoteId) {
      item.classList.add('active');
    }

    const bookName = getBookNameById(note.bookId);

    item.innerHTML = `
      <div class="note-item-header">
        <span class="note-item-date">${note.date}</span>
        ${note.bookId ? `<span class="note-item-tag">${bookName}</span>` : ''}
      </div>
      <h4 class="note-item-title">${note.title}</h4>
      <p class="note-item-preview">${note.content}</p>
    `;

    item.addEventListener('click', () => {
      displayNoteInForm(note.id);
    });

    listContainer.appendChild(item);
  });
}

function displayNoteInForm(noteId) {
  state.currentNoteId = noteId;
  const note = state.notes.find(n => n.id === noteId);
  if (!note) return;

  document.getElementById('editor-action-title').textContent = "Modifier la Note";
  document.getElementById('note-title-input').value = note.title;
  document.getElementById('note-book-select').value = note.bookId;
  document.getElementById('note-content-textarea').value = note.content;
  
  document.getElementById('delete-note-btn').style.display = 'inline-flex';
  
  renderNotesList();
}

function resetNoteForm() {
  state.currentNoteId = null;
  document.getElementById('editor-action-title').textContent = "Écrire une Note";
  document.getElementById('note-title-input').value = '';
  document.getElementById('note-book-select').value = '';
  document.getElementById('note-content-textarea').value = '';
  
  document.getElementById('delete-note-btn').style.display = 'none';

  document.querySelectorAll('.note-item-card').forEach(card => {
    card.classList.remove('active');
  });
}

// --- FONCTIONS UTILITAIRES ---
function getBookNameById(bookId) {
  if (!bookId) return "";
  const ch = state.bibleChapters.find(c => c.bookId === bookId);
  return ch ? ch.book : "";
}

function formatDateKey(date) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

function getFormattedDate(date) {
  const months = [
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin", 
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
  ];
  const day = date.getDate();
  const monthName = months[date.getMonth()];
  const year = date.getFullYear();
  return `${day} ${monthName} ${year}`;
}

function getBookTheme(bookId) {
  const PENTATEUCH = ["genese", "exode", "levitique", "nombres", "deuteronome"];
  const HISTORICAL = ["josue", "juges", "ruth", "1samuel", "2samuel", "1rois", "2rois", "1chroniques", "2chroniques", "esdras", "nehemie", "esther"];
  const POETRY = ["job", "psaumes", "proverbes", "ecclesiaste", "cantique"];
  const MAJOR_PROPHETS = ["esaie", "jeremie", "lamentations", "ezechiel", "daniel"];
  const GOSPELS = ["matthieu", "marc", "luc", "jean", "actes"];
  const MINOR_PROPHETS = ["osee", "joel", "amos", "abdias", "jonas", "michee", "nahum", "habacuc", "sophonie", "aggee", "zacharie", "malachie"];

  if (PENTATEUCH.includes(bookId)) {
    return { color: '#10b981', border: 'rgba(16, 185, 129, 0.25)', bg: 'rgba(16, 185, 129, 0.1)' }; // Green
  }
  if (HISTORICAL.includes(bookId)) {
    return { color: '#f59e0b', border: 'rgba(245, 158, 11, 0.25)', bg: 'rgba(245, 158, 11, 0.1)' }; // Amber
  }
  if (POETRY.includes(bookId)) {
    return { color: '#ec4899', border: 'rgba(236, 72, 153, 0.25)', bg: 'rgba(236, 72, 153, 0.1)' }; // Rose
  }
  if (MAJOR_PROPHETS.includes(bookId)) {
    return { color: '#06b6d4', border: 'rgba(6, 182, 212, 0.25)', bg: 'rgba(6, 182, 212, 0.1)' }; // Cyan
  }
  if (MINOR_PROPHETS.includes(bookId)) {
    return { color: '#f97316', border: 'rgba(249, 115, 22, 0.25)', bg: 'rgba(249, 115, 22, 0.1)' }; // Orange
  }
  if (GOSPELS.includes(bookId)) {
    return { color: '#6366f1', border: 'rgba(99, 102, 241, 0.25)', bg: 'rgba(99, 102, 241, 0.1)' }; // Indigo
  }
  if (bookId === "apocalypse") {
    return { color: '#ef4444', border: 'rgba(239, 68, 68, 0.25)', bg: 'rgba(239, 68, 68, 0.1)' }; // Red
  }
  return { color: '#14b8a6', border: 'rgba(20, 184, 166, 0.25)', bg: 'rgba(20, 184, 166, 0.1)' }; // Teal
}

function checkWelcomeProfileState() {
  const profileModal = document.getElementById('profile-modal');
  const cancelModalBtn = document.getElementById('close-profile-modal-btn');
  const modalTitle = document.querySelector('#profile-modal .card-title');

  if (state.profiles.length === 0) {
    cancelModalBtn.style.display = 'none';
    modalTitle.innerHTML = '<i class="fa-solid fa-user-astronaut"></i> Créez votre premier Profil';
    profileModal.classList.add('active');
  } else {
    cancelModalBtn.style.display = 'inline-flex';
    modalTitle.innerHTML = '<i class="fa-solid fa-user-plus" style="color: var(--color-primary)"></i> Nouveau Profil';
  }
}
