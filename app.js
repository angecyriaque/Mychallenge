// Bible Tracker - Logique Applicative (app.js)

// --- GESTION DE L'ÉTAT LOCAL (STATE) ---
let state = {
  bibleChapters: [], // Chargé depuis Excel ou local: [{ id, book, bookId, chapter, ref, testament }]
  readChapters: [],  // Tableau de chaînes 'bookId-chapterNumber' (ex: 'genese-1')
  dailyGoal: 3,      // Objectif de lecture par jour
  notes: [],         // Tableau d'objets { id, title, bookId, content, date }
  historyLog: {},    // { 'YYYY-MM-DD': nombre_de_chapitres_lus_ce_jour }
  currentNoteId: null, // ID de la note en cours d'édition (null si nouvelle note)
  activeTab: 'dashboard-tab',
  currentTheme: 'dark',
  
  // Pagination pour la To-Do list
  currentPage: 1,
  itemsPerPage: 50
};

// Éléments DOM Globaux
let readingChart = null;

// Livres de l'Ancien Testament pour trier
const OT_BOOKS = [
  "Genèse", "Exode", "Lévitique", "Nombres", "Deutéronome", "Josué", "Juges", "Ruth",
  "1 Samuel", "2 Samuel", "1 Rois", "2 Rois", "1 Chroniques", "2 Chroniques", "Esdras",
  "Néhémie", "Esther", "Job", "Psaumes", "Proverbes", "Ecclésiaste", "Cantique des Cantiques",
  "Ésaïe", "Jérémie", "Lamentations", "Ézéchiel", "Daniel", "Osée", "Joël", "Amos",
  "Abdias", "Jonas", "Michée", "Nahum", "Habacuc", "Sophonie", "Aggée", "Zacharie", "Malachie"
];

// --- INITIALISATION DE L'APPLICATION ---
document.addEventListener('DOMContentLoaded', async () => {
  loadLocalStorage();
  initTheme();
  initTabNavigation();
  initGoalControls();
  
  // Charger les données de la Bible (Excel ou secours)
  await initBibleData();
  
  populateBookSelect();
  renderReadingList();
  renderNotesList();
  updateDashboard();
  initNotesControls();
  
  // Ecouteurs sur les filtres de la liste de lecture
  document.querySelectorAll('.filter-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
      e.currentTarget.classList.add('active');
      state.currentPage = 1; // Reset pagination
      renderReadingList();
    });
  });

  // Écouteur sur la recherche de livre
  document.getElementById('book-search-input').addEventListener('input', () => {
    state.currentPage = 1; // Reset pagination
    renderReadingList();
  });
});

// --- CHARGEMENT DYNAMIQUE DU FICHIER EXCEL ---
async function initBibleData() {
  try {
    const response = await fetch('/liste_chapitres_bible.xlsx');
    if (!response.ok) throw new Error("Fichier introuvable.");
    
    const arrayBuffer = await response.buffer ? await response.buffer() : await response.arrayBuffer();
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

// Helper pour déterminer le testament
function getTestamentByBookName(bookName) {
  return OT_BOOKS.includes(bookName) ? 'Ancien' : 'Nouveau';
}

// Helper pour formater l'identifiant (slugify)
function slugify(text) {
  return text.toString().toLowerCase().trim()
    .replace(/\s+/g, '')           // Enlever les espaces
    .replace(/[éèêë]/g, 'e')       // Remplacer accents
    .replace(/[àâä]/g, 'a')
    .replace(/[îï]/g, 'i')
    .replace(/[ôö]/g, 'o')
    .replace(/[ûü]/g, 'u')
    .replace(/[ç]/g, 'c')
    .replace(/[^a-z0-9]/g, '');    // Retirer non-alphanumérique
}

// --- PERSISTANCE LOCAL STORAGE ---
function loadLocalStorage() {
  const savedReadChapters = localStorage.getItem('bible_tracker_read_chapters');
  if (savedReadChapters) {
    state.readChapters = JSON.parse(savedReadChapters);
  }

  const savedDailyGoal = localStorage.getItem('bible_tracker_daily_goal');
  if (savedDailyGoal) {
    state.dailyGoal = parseInt(savedDailyGoal, 10);
    document.getElementById('daily-goal-input').value = state.dailyGoal;
  }

  const savedNotes = localStorage.getItem('bible_tracker_notes');
  if (savedNotes) {
    state.notes = JSON.parse(savedNotes);
  } else {
    state.notes = [
      {
        id: 'example-1',
        title: 'Premières pensées sur la Création',
        bookId: 'genese',
        content: 'La Genèse pose les bases de notre foi. Le chapitre 1 révèle l\'ordre, la beauté et la puissance de la parole de Dieu. Tout ce qu\'Il crée est qualifié de "bon", puis "très bon" après la création de l\'homme. Cela montre la valeur infinie que nous avons à Ses yeux.',
        date: getFormattedDate(new Date(Date.now() - 24 * 60 * 60 * 1000 * 2))
      },
      {
        id: 'example-2',
        title: 'La Grâce de l\'Évangile de Jean',
        bookId: 'jean',
        content: 'Jean 1:1-14 est un prologue magnifique. "Le Verbe s\'est fait chair." La divinité de Jésus combinée à Son humanité est le cœur de la foi chrétienne. Il apporte la grâce et la vérité.',
        date: getFormattedDate(new Date())
      }
    ];
    saveNotesToStorage();
  }

  const savedHistoryLog = localStorage.getItem('bible_tracker_history');
  if (savedHistoryLog) {
    state.historyLog = JSON.parse(savedHistoryLog);
  } else {
    if (state.readChapters.length > 0) {
      seedHistoryLog();
    } else {
      state.historyLog = {};
    }
    saveHistoryToStorage();
  }
}

function saveReadChaptersToStorage() {
  localStorage.setItem('bible_tracker_read_chapters', JSON.stringify(state.readChapters));
}

function saveNotesToStorage() {
  localStorage.setItem('bible_tracker_notes', JSON.stringify(state.notes));
}

function saveHistoryToStorage() {
  localStorage.setItem('bible_tracker_history', JSON.stringify(state.historyLog));
}

function seedHistoryLog() {
  const dates = [];
  for (let i = 6; i >= 0; i--) {
    const d = new Date();
    d.setDate(d.getDate() - i);
    dates.push(formatDateKey(d));
  }
  
  let remaining = state.readChapters.length;
  dates.forEach((dateKey, index) => {
    if (index === dates.length - 1) {
      state.historyLog[dateKey] = remaining;
    } else {
      const chunk = Math.min(remaining, Math.floor(Math.random() * 3));
      state.historyLog[dateKey] = chunk;
      remaining -= chunk;
    }
  });
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

// --- NAVIGATION PAR ONGLETS ---
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

// --- OBJECTIF QUOTIDIEN ---
function initGoalControls() {
  const input = document.getElementById('daily-goal-input');
  const minusBtn = document.getElementById('goal-minus-btn');
  const plusBtn = document.getElementById('goal-plus-btn');

  const updateGoal = (newVal) => {
    if (newVal < 1) newVal = 1;
    if (newVal > 150) newVal = 150;
    state.dailyGoal = newVal;
    input.value = newVal;
    localStorage.setItem('bible_tracker_daily_goal', newVal);
    renderReadingList();
    updateDashboard();
  };

  minusBtn.addEventListener('click', () => updateGoal(state.dailyGoal - 1));
  plusBtn.addEventListener('click', () => updateGoal(state.dailyGoal + 1));
  input.addEventListener('change', (e) => updateGoal(parseInt(e.target.value, 10) || 3));
}

// --- POPULER LES LIVRES DANS LE JOURNAL ---
function populateBookSelect() {
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

// --- OBTENIR LE PLAN DE LECTURE DU JOUR ---
function getTodaysReadingPlan() {
  const plan = [];
  let count = 0;
  
  for (const ch of state.bibleChapters) {
    if (!state.readChapters.includes(ch.id)) {
      plan.push(ch.id);
      count++;
      if (count >= state.dailyGoal) {
        return plan;
      }
    }
  }
  return plan;
}

// --- RENDU DE LA LISTE DE LECTURE (CHECKLIST / TO-DO) ---
function renderReadingList() {
  const container = document.getElementById('books-list-container');
  const paginationContainer = document.getElementById('pagination-container');
  const searchQuery = document.getElementById('book-search-input').value.toLowerCase().trim();
  const activeFilter = document.querySelector('.filter-btn.active').getAttribute('data-filter');
  
  container.innerHTML = '';
  paginationContainer.innerHTML = '';
  
  const todaysPlan = getTodaysReadingPlan();
  
  // 1. Filtrer les chapitres
  const filteredChapters = state.bibleChapters.filter(ch => {
    const matchesSearch = ch.ref.toLowerCase().includes(searchQuery) || ch.book.toLowerCase().includes(searchQuery);
    if (!matchesSearch) return false;
    
    const isRead = state.readChapters.includes(ch.id);
    if (activeFilter === 'read') return isRead;
    if (activeFilter === 'unread') return !isRead;
    if (activeFilter === 'today') return todaysPlan.includes(ch.id);
    
    return true;
  });

  if (filteredChapters.length === 0) {
    container.innerHTML = `
      <div class="card empty-state" style="grid-column: 1 / -1; width: 100%;">
        <i class="fa-solid fa-magnifying-glass"></i>
        <h3>Aucun chapitre trouvé</h3>
        <p>Ajustez vos filtres ou la barre de recherche pour explorer la Bible.</p>
      </div>
    `;
    return;
  }

  // 2. Pagination (sauf pour le filtre "Aujourd'hui")
  let displayChapters = filteredChapters;
  const isFilterToday = activeFilter === 'today';
  
  if (!isFilterToday) {
    const totalItems = filteredChapters.length;
    const totalPages = Math.ceil(totalItems / state.itemsPerPage);
    
    if (state.currentPage > totalPages) {
      state.currentPage = 1;
    }
    
    const startIndex = (state.currentPage - 1) * state.itemsPerPage;
    const endIndex = startIndex + state.itemsPerPage;
    displayChapters = filteredChapters.slice(startIndex, endIndex);
    
    if (totalPages > 1) {
      paginationContainer.innerHTML = `
        <button class="pagination-btn" id="prev-page-btn" ${state.currentPage === 1 ? 'disabled' : ''}>
          <i class="fa-solid fa-chevron-left"></i>
        </button>
        <span class="pagination-info">Page ${state.currentPage} sur ${totalPages} (${totalItems} chapitres)</span>
        <button class="pagination-btn" id="next-page-btn" ${state.currentPage === totalPages ? 'disabled' : ''}>
          <i class="fa-solid fa-chevron-right"></i>
        </button>
      `;
      
      document.getElementById('prev-page-btn').addEventListener('click', () => {
        state.currentPage--;
        renderReadingList();
        container.scrollIntoView({ behavior: 'smooth' });
      });
      
      document.getElementById('next-page-btn').addEventListener('click', () => {
        state.currentPage++;
        renderReadingList();
        container.scrollIntoView({ behavior: 'smooth' });
      });
    }
  }

  // 3. Rendu des lignes To-Do
  displayChapters.forEach(ch => {
    const isRead = state.readChapters.includes(ch.id);
    const isHighlighted = todaysPlan.includes(ch.id);
    
    const todoItem = document.createElement('div');
    todoItem.className = 'todo-item';
    if (isRead) todoItem.classList.add('checked');
    if (isHighlighted && !isRead) todoItem.classList.add('highlighted');
    
    const badgeClass = ch.testament === 'Ancien' ? 'ot' : 'nt';
    const badgeText = ch.testament === 'Ancien' ? 'AT' : 'NT';
    
    todoItem.innerHTML = `
      <div class="todo-item-left">
        <button class="todo-check-btn" aria-label="Marquer comme lu">
          <i class="fa-solid fa-check"></i>
        </button>
        <span class="todo-ref">${ch.ref}</span>
      </div>
      <span class="todo-badge ${badgeClass}">${badgeText}</span>
    `;
    
    todoItem.addEventListener('click', () => {
      toggleChapterTodo(ch.id, todoItem);
    });
    
    container.appendChild(todoItem);
  });
}

// Action cocher / décocher To-Do item
function toggleChapterTodo(chapterId, todoItemElement) {
  const index = state.readChapters.indexOf(chapterId);
  const todayKey = formatDateKey(new Date());
  
  if (index === -1) {
    state.readChapters.push(chapterId);
    todoItemElement.classList.remove('highlighted');
    todoItemElement.classList.add('checked');
    state.historyLog[todayKey] = (state.historyLog[todayKey] || 0) + 1;
  } else {
    state.readChapters.splice(index, 1);
    todoItemElement.classList.remove('checked');
    const todaysPlan = getTodaysReadingPlan();
    if (todaysPlan.includes(chapterId)) {
      todoItemElement.classList.add('highlighted');
    }
    if (state.historyLog[todayKey] && state.historyLog[todayKey] > 0) {
      state.historyLog[todayKey]--;
    }
  }

  saveReadChaptersToStorage();
  saveHistoryToStorage();
  
  // Petit rafraîchissement différé pour recalculer l'ordre et la pagination
  setTimeout(() => {
    renderReadingList();
  }, 250);

  if (state.activeTab === 'dashboard-tab') {
    updateDashboard();
  }
}

// --- TABLEAU DE BORD & ANALYSE ---
function updateDashboard() {
  if (state.bibleChapters.length === 0) return;

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

  renderOrUpdateChart();
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

    saveNotesToStorage();
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
      saveNotesToStorage();
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
