const STORAGE_KEYS = {
  users: "training_hub_users_v2",
  progress: "training_hub_progress_v2",
  excelReports: "training_hub_excel_reports_v1",
  theme: "training_hub_theme",
  adminSession: "training_hub_admin_session",
  traineeSessionUserId: "training_hub_trainee_session_user",
  videoFallback: "training_hub_videos_fallback_v1",
};

// Change admin credentials here.
const ADMIN_CREDENTIALS = {
  username: "vcharan71-sys",
  passcode: "AutoMynd@2026",
};

const DB_NAME = "training_hub_db";
const DB_VERSION = 1;
const VIDEO_STORE = "videos";

const state = {
  users: loadUsers(),
  progress: loadProgress(),
  excelReports: loadExcelReports(),
  videos: [],
  activeView: "training",
  searchQuery: "",
  folderFilter: "all",
  activeExcelReportId: "",
  isAdminUnlocked: localStorage.getItem(STORAGE_KEYS.adminSession) === "1",
  activeTraineeId: localStorage.getItem(STORAGE_KEYS.traineeSessionUserId) || "",
  analyticsUserId: "",
  analyticsSearchQuery: "",
  runtimeIssues: [],
  videoStorageMode: "indexeddb",
  transientVideos: [],
  quizEditorVideoId: "",
};

const els = {
  trainingTabBtn: document.getElementById("trainingTabBtn"),
  adminTabBtn: document.getElementById("adminTabBtn"),
  trainingView: document.getElementById("trainingView"),
  adminView: document.getElementById("adminView"),
  runtimeStatus: document.getElementById("runtimeStatus"),
  adminLogoutBtn: document.getElementById("adminLogoutBtn"),
  traineeLogoutBtn: document.getElementById("traineeLogoutBtn"),
  themeToggle: document.getElementById("themeToggle"),

  traineeGate: document.getElementById("traineeGate"),
  traineeLoginForm: document.getElementById("traineeLoginForm"),
  traineeEmail: document.getElementById("traineeEmail"),
  traineePassword: document.getElementById("traineePassword"),
  traineeGateMessage: document.getElementById("traineeGateMessage"),
  traineeContent: document.getElementById("traineeContent"),
  activeTraineeDisplay: document.getElementById("activeTraineeDisplay"),
  folderFilterSelect: document.getElementById("folderFilterSelect"),
  passwordChangePanel: document.getElementById("passwordChangePanel"),
  passwordChangeForm: document.getElementById("passwordChangeForm"),
  newPassword: document.getElementById("newPassword"),
  confirmPassword: document.getElementById("confirmPassword"),
  passwordChangeMessage: document.getElementById("passwordChangeMessage"),
  trainingWorkspace: document.getElementById("trainingWorkspace"),

  videoSearchInput: document.getElementById("videoSearchInput"),
  videoCardGrid: document.getElementById("videoCardGrid"),
  trainingPlayer: document.getElementById("trainingPlayer"),
  playerTitle: document.getElementById("playerTitle"),
  videoMeta: document.getElementById("videoMeta"),
  completionCheckbox: document.getElementById("completionCheckbox"),
  openQuizBtn: document.getElementById("openQuizBtn"),
  completionNote: document.getElementById("completionNote"),
  videoStatus: document.getElementById("videoStatus"),

  adminGate: document.getElementById("adminGate"),
  adminLoginForm: document.getElementById("adminLoginForm"),
  adminUsername: document.getElementById("adminUsername"),
  adminPasscode: document.getElementById("adminPasscode"),
  adminGateMessage: document.getElementById("adminGateMessage"),
  adminContent: document.getElementById("adminContent"),

  joineeForm: document.getElementById("joineeForm"),
  joineeName: document.getElementById("joineeName"),
  joineeEmail: document.getElementById("joineeEmail"),
  joineePassword: document.getElementById("joineePassword"),

  videoForm: document.getElementById("videoForm"),
  videoTitle: document.getElementById("videoTitle"),
  videoFolderSelect: document.getElementById("videoFolderSelect"),
  videoFolderNewWrap: document.getElementById("videoFolderNewWrap"),
  videoFolderNew: document.getElementById("videoFolderNew"),
  videoDescription: document.getElementById("videoDescription"),
  videoFile: document.getElementById("videoFile"),

  joineeList: document.getElementById("joineeList"),
  videoList: document.getElementById("videoList"),
  progressTableBody: document.getElementById("progressTableBody"),
  analyticsSummary: document.getElementById("analyticsSummary"),
  analyticsSearchInput: document.getElementById("analyticsSearchInput"),
  analyticsCards: document.getElementById("analyticsCards"),
  excelUploadForm: document.getElementById("excelUploadForm"),
  excelFile: document.getElementById("excelFile"),
  excelReportSelect: document.getElementById("excelReportSelect"),
  excelSheetSelect: document.getElementById("excelSheetSelect"),
  excelCategorySelect: document.getElementById("excelCategorySelect"),
  excelValueSelect: document.getElementById("excelValueSelect"),
  excelStatus: document.getElementById("excelStatus"),
  excelEmptyState: document.getElementById("excelEmptyState"),
  excelResults: document.getElementById("excelResults"),
  excelReportMeta: document.getElementById("excelReportMeta"),
  excelChartTitle: document.getElementById("excelChartTitle"),
  excelChartHeadline: document.getElementById("excelChartHeadline"),
  excelChartDescription: document.getElementById("excelChartDescription"),
  excelPieChart: document.getElementById("excelPieChart"),
  excelLegend: document.getElementById("excelLegend"),
  excelTableBody: document.getElementById("excelTableBody"),
  analyticsModal: document.getElementById("analyticsModal"),
  analyticsModalBackdrop: document.getElementById("analyticsModalBackdrop"),
  analyticsModalCard: document.getElementById("analyticsModalCard"),
  analyticsModalTitle: document.getElementById("analyticsModalTitle"),
  analyticsModalBody: document.getElementById("analyticsModalBody"),
  analyticsModalClose: document.getElementById("analyticsModalClose"),
  quizModal: document.getElementById("quizModal"),
  quizModalBackdrop: document.getElementById("quizModalBackdrop"),
  quizModalCard: document.getElementById("quizModalCard"),
  quizModalTitle: document.getElementById("quizModalTitle"),
  quizModalClose: document.getElementById("quizModalClose"),
  quizQuestionText: document.getElementById("quizQuestionText"),
  quizAnswerForm: document.getElementById("quizAnswerForm"),
  quizFeedback: document.getElementById("quizFeedback"),
  quizEditorModal: document.getElementById("quizEditorModal"),
  quizEditorBackdrop: document.getElementById("quizEditorBackdrop"),
  quizEditorCard: document.getElementById("quizEditorCard"),
  quizEditorClose: document.getElementById("quizEditorClose"),
  quizEditorForm: document.getElementById("quizEditorForm"),
  quizEditorQuestion: document.getElementById("quizEditorQuestion"),
  quizEditorOption1: document.getElementById("quizEditorOption1"),
  quizEditorOption2: document.getElementById("quizEditorOption2"),
  quizEditorOption3: document.getElementById("quizEditorOption3"),
  quizEditorOption4: document.getElementById("quizEditorOption4"),
  quizEditorCorrect: document.getElementById("quizEditorCorrect"),
  deleteQuizBtn: document.getElementById("deleteQuizBtn"),
};

const dbPromise = openDatabase().catch((error) => {
  reportRuntimeIssue(
    "warning",
    `Persistent video storage is unavailable in this browser. Video uploads will work only for this tab until you refresh. (${error.message || "Storage error"})`
  );
  state.videoStorageMode = "memory";
  return null;
});

init().catch((error) => {
  console.error(error);
  reportRuntimeIssue(
    "error",
    `The portal loaded with limited functionality because startup hit an error. ${error.message || "Unknown startup error."}`
  );
  renderAll();
});

async function init() {
  applySavedTheme();
  validateTraineeSession();
  syncAnalyticsModalStyles();
  syncModalStyles(els.quizModal, els.quizModalBackdrop, els.quizModalCard, els.quizModalClose);
  syncModalStyles(els.quizEditorModal, els.quizEditorBackdrop, els.quizEditorCard, els.quizEditorClose);

  wireHeaderAndViews();
  wireTraineeAuth();
  wirePasswordChange();
  wireTraining();
  wireAdmin();
  wirePlayerTracking();

  if (!state.activeExcelReportId && state.excelReports.length > 0) {
    state.activeExcelReportId = state.excelReports[0].id;
  }

  state.videos = await getAllVideos();
  renderAll();
}

function reportRuntimeIssue(level, message) {
  if (state.runtimeIssues.some((issue) => issue.message === message)) {
    updateRuntimeStatus();
    return;
  }

  state.runtimeIssues.push({ level, message });
  updateRuntimeStatus();
}

function updateRuntimeStatus() {
  if (!state.runtimeIssues.length) {
    els.runtimeStatus.className = "status-banner hidden";
    els.runtimeStatus.textContent = "";
    return;
  }

  const hasError = state.runtimeIssues.some((issue) => issue.level === "error");
  els.runtimeStatus.className = `status-banner ${hasError ? "error" : "warning"}`;
  els.runtimeStatus.textContent = state.runtimeIssues.map((issue) => issue.message).join(" ");
}

function validateTraineeSession() {
  if (!state.activeTraineeId) {
    return;
  }

  const activeUser = state.users.find((user) => user.id === state.activeTraineeId);
  if (!activeUser || activeUser.status === "disabled") {
    state.activeTraineeId = "";
    localStorage.removeItem(STORAGE_KEYS.traineeSessionUserId);
  }
}

function wireHeaderAndViews() {
  els.themeToggle.addEventListener("click", () => {
    document.body.classList.toggle("dark");
    const isDark = document.body.classList.contains("dark");
    localStorage.setItem(STORAGE_KEYS.theme, isDark ? "dark" : "light");
  });

  els.trainingTabBtn.addEventListener("click", () => switchView("training"));
  els.adminTabBtn.addEventListener("click", () => switchView("admin"));

  els.adminLogoutBtn.addEventListener("click", () => {
    state.isAdminUnlocked = false;
    localStorage.removeItem(STORAGE_KEYS.adminSession);
    updateAdminVisibility();
    updateNavigationVisibility();
    switchView("training");
  });

  els.traineeLogoutBtn.addEventListener("click", () => {
    state.activeTraineeId = "";
    localStorage.removeItem(STORAGE_KEYS.traineeSessionUserId);
    els.passwordChangeMessage.textContent = "";
    clearPlayer();
    renderTraineeVisibility();
    renderVideoCards();
    renderPlayerStatus();
    updateNavigationVisibility();
    switchView("training");
  });

  updateAdminVisibility();
  renderTraineeVisibility();
  updateNavigationVisibility();
  switchView("training");
}

function wireTraineeAuth() {
  els.traineeLoginForm.addEventListener("submit", (event) => {
    event.preventDefault();

    const email = els.traineeEmail.value.trim().toLowerCase();
    const password = els.traineePassword.value;

    const matchedUser = state.users.find((user) => user.email === email && user.password === password);
    if (!matchedUser) {
      els.traineeGateMessage.textContent = "Invalid trainee credentials. Please try again.";
      return;
    }

    if (matchedUser.status === "disabled") {
      els.traineeGateMessage.textContent = "This employee account is disabled in the GitHub testing build.";
      return;
    }

    state.isAdminUnlocked = false;
    localStorage.removeItem(STORAGE_KEYS.adminSession);
    state.activeTraineeId = matchedUser.id;
    localStorage.setItem(STORAGE_KEYS.traineeSessionUserId, matchedUser.id);
    els.passwordChangeMessage.textContent = "";
    els.traineeGateMessage.textContent = matchedUser.mustChangePassword
      ? "Temporary password accepted. Please create a new password to continue."
      : "";
    els.traineeLoginForm.reset();

    renderTraineeVisibility();
    updateAdminVisibility();
    updateNavigationVisibility();
    renderVideoCards();

    if (!currentVideoId() && state.videos.length > 0) {
      selectVideo(state.videos[0].id);
    } else {
      refreshPlayerPanel();
    }

    switchView("training");
  });
}

function wirePasswordChange() {
  els.passwordChangeForm.addEventListener("submit", (event) => {
    event.preventDefault();

    const activeUser = getActiveTrainee();
    if (!activeUser) {
      return;
    }

    const nextPassword = els.newPassword.value.trim();
    const confirmPassword = els.confirmPassword.value.trim();

    if (nextPassword.length < 8) {
      els.passwordChangeMessage.textContent = "Use at least 8 characters for the new password.";
      return;
    }

    if (nextPassword !== confirmPassword) {
      els.passwordChangeMessage.textContent = "The passwords do not match.";
      return;
    }

    activeUser.password = nextPassword;
    activeUser.mustChangePassword = false;
    saveUsers(state.users);

    els.passwordChangeMessage.textContent = "Password updated. You can now use the training modules below.";
    els.passwordChangeForm.reset();
    renderTraineeVisibility();
    renderVideoCards();
    refreshPlayerPanel();
  });
}

function switchView(viewName) {
  if (viewName === "admin" && getActiveTrainee()) {
    viewName = "training";
  }

  state.activeView = viewName;
  const training = viewName === "training";

  els.trainingTabBtn.classList.toggle("active", training);
  els.adminTabBtn.classList.toggle("active", !training);
  els.trainingView.classList.toggle("active", training);
  els.adminView.classList.toggle("active", !training);

  if (!training) {
    updateAdminVisibility();
  }
}

function updateAdminVisibility() {
  els.adminGate.classList.toggle("hidden", state.isAdminUnlocked);
  els.adminContent.classList.toggle("hidden", !state.isAdminUnlocked);
  els.adminLogoutBtn.classList.toggle("hidden", !state.isAdminUnlocked);
}

function updateNavigationVisibility() {
  const traineeLoggedIn = Boolean(getActiveTrainee());

  els.adminTabBtn.classList.toggle("hidden", traineeLoggedIn);
  els.trainingTabBtn.classList.toggle("hidden", state.isAdminUnlocked && !traineeLoggedIn);
}

function renderTraineeVisibility() {
  const activeUser = getActiveTrainee();
  const isLoggedIn = Boolean(activeUser);
  const mustChangePassword = Boolean(activeUser?.mustChangePassword);

  els.traineeGate.classList.toggle("hidden", isLoggedIn);
  els.traineeContent.classList.toggle("hidden", !isLoggedIn);
  els.traineeLogoutBtn.classList.toggle("hidden", !isLoggedIn);
  els.passwordChangePanel.classList.toggle("hidden", !mustChangePassword);
  els.trainingWorkspace.classList.toggle("hidden", mustChangePassword);
  if (!mustChangePassword) {
    els.passwordChangeMessage.textContent = "";
  }
  els.activeTraineeDisplay.value = activeUser
    ? `${activeUser.name} (${activeUser.email})${activeUser.status === "disabled" ? " - disabled" : ""}`
    : "";
}

function wireTraining() {
  els.openQuizBtn.addEventListener("click", () => {
    openQuizModal(currentVideoId());
  });

  els.folderFilterSelect.addEventListener("change", () => {
    state.folderFilter = els.folderFilterSelect.value;
    renderVideoCards();
  });

  els.videoSearchInput.addEventListener("input", () => {
    state.searchQuery = els.videoSearchInput.value.trim().toLowerCase();
    renderVideoCards();
  });

  els.completionCheckbox.addEventListener("change", () => {
    const userId = state.activeTraineeId;
    const videoId = currentVideoId();
    if (!userId || !videoId) {
      return;
    }

    const key = progressKey(userId, videoId);
    const entry = state.progress[key] || createEmptyProgressEntry();

    if (els.completionCheckbox.checked && !isEligibleForManualCompletion(entry)) {
      els.completionCheckbox.checked = false;
      return;
    }

    entry.manualCompleted = els.completionCheckbox.checked;
    entry.completedAt = entry.manualCompleted ? new Date().toISOString() : null;
    state.progress[key] = entry;
    saveProgress(state.progress);

    renderVideoCards();
    renderPlayerStatus();
    renderProgressTable();
    renderAnalyticsPanel();
  });
}

function wireAdmin() {
  els.adminLoginForm.addEventListener("submit", (event) => {
    event.preventDefault();

    const username = els.adminUsername.value.trim();
    const passcode = els.adminPasscode.value;

    if (username !== ADMIN_CREDENTIALS.username || passcode !== ADMIN_CREDENTIALS.passcode) {
      els.adminGateMessage.textContent = "Invalid credentials. Please try again.";
      return;
    }

    state.isAdminUnlocked = true;
    localStorage.setItem(STORAGE_KEYS.adminSession, "1");
    state.activeTraineeId = "";
    localStorage.removeItem(STORAGE_KEYS.traineeSessionUserId);
    els.adminGateMessage.textContent = "";
    els.adminLoginForm.reset();
    els.passwordChangeMessage.textContent = "";

    updateAdminVisibility();
    renderTraineeVisibility();
    updateNavigationVisibility();
    renderAdminData();
    switchView("admin");
  });

  els.joineeForm.addEventListener("submit", (event) => {
    event.preventDefault();

    const name = els.joineeName.value.trim();
    const email = els.joineeEmail.value.trim().toLowerCase();
    const password = els.joineePassword.value.trim();

    if (!name || !email || !password) {
      return;
    }

    if (state.users.some((user) => user.email === email)) {
      alert("A joinee with this email already exists.");
      return;
    }

    state.users.push({
      id: crypto.randomUUID(),
      name,
      email,
      password,
      status: "active",
      mustChangePassword: true,
      joinedAt: new Date().toISOString(),
    });

    saveUsers(state.users);
    els.joineeForm.reset();
    renderAll();
  });

  els.videoFolderSelect.addEventListener("change", updateVideoFolderInputVisibility);

  els.videoForm.addEventListener("submit", async (event) => {
    event.preventDefault();

    const title = els.videoTitle.value.trim();
    const folder =
      els.videoFolderSelect.value === "__new__"
        ? els.videoFolderNew.value.trim()
        : els.videoFolderSelect.value.trim();
    const description = els.videoDescription.value.trim();
    const file = els.videoFile.files?.[0];
    if (!title || !folder || !file) {
      return;
    }

    await saveVideo({
      id: crypto.randomUUID(),
      title,
      folder,
      description,
      fileName: file.name,
      mimeType: file.type,
      size: file.size,
      createdAt: new Date().toISOString(),
      blob: file,
    });

    state.videos = await getAllVideos();
    els.videoForm.reset();
    updateVideoFolderOptions();
    updateVideoFolderInputVisibility();
    renderAll();
  });

  els.analyticsSearchInput.addEventListener("input", () => {
    state.analyticsSearchQuery = els.analyticsSearchInput.value.trim().toLowerCase();
    renderAnalyticsPanel();
  });

  els.excelReportSelect.addEventListener("change", () => {
    state.activeExcelReportId = els.excelReportSelect.value;
    const report = getActiveExcelReport();
    if (report && (!report.activeSheetName || !report.sheets.some((sheet) => sheet.name === report.activeSheetName))) {
      report.activeSheetName = report.sheets[0]?.name || "";
      saveExcelReports(state.excelReports);
    }
    renderExcelAnalytics();
  });

  els.excelSheetSelect.addEventListener("change", () => {
    const report = getActiveExcelReport();
    if (!report) {
      return;
    }

    report.activeSheetName = els.excelSheetSelect.value;
    const sheet = getActiveExcelSheet(report);
    if (sheet) {
      applySuggestedSheetConfig(sheet);
    }
    saveExcelReports(state.excelReports);
    renderExcelAnalytics();
  });

  els.excelCategorySelect.addEventListener("change", () => {
    const sheet = getActiveExcelSheet();
    if (!sheet) {
      return;
    }

    sheet.config = sheet.config || {};
    sheet.config.categoryColumn = els.excelCategorySelect.value;
    saveExcelReports(state.excelReports);
    renderExcelAnalytics();
  });

  els.excelValueSelect.addEventListener("change", () => {
    const sheet = getActiveExcelSheet();
    if (!sheet) {
      return;
    }

    sheet.config = sheet.config || {};
    sheet.config.valueColumn = els.excelValueSelect.value === "__row_count__" ? "" : els.excelValueSelect.value;
    saveExcelReports(state.excelReports);
    renderExcelAnalytics();
  });

  els.excelUploadForm.addEventListener("submit", async (event) => {
    event.preventDefault();

    const file = els.excelFile.files?.[0];
    if (file) {
      await importExcelReport(file);
      return;
    }

    const sheet = getActiveExcelSheet();
    if (!sheet) {
      els.excelStatus.textContent = "Choose a report or upload a file first.";
      return;
    }

    sheet.config = {
      categoryColumn: els.excelCategorySelect.value,
      valueColumn: els.excelValueSelect.value === "__row_count__" ? "" : els.excelValueSelect.value,
    };
    saveExcelReports(state.excelReports);
    renderExcelAnalytics();
  });

  els.analyticsModalClose.addEventListener("click", closeAnalyticsModal);
  els.analyticsModal.addEventListener("click", (event) => {
    if (event.target.dataset.closeModal === "true") {
      closeAnalyticsModal();
    }
  });

  els.quizModalClose.addEventListener("click", closeQuizModal);
  els.quizModal.addEventListener("click", (event) => {
    if (event.target.dataset.closeQuiz === "true") {
      closeQuizModal();
    }
  });

  els.quizEditorClose.addEventListener("click", closeQuizEditorModal);
  els.quizEditorModal.addEventListener("click", (event) => {
    if (event.target.dataset.closeQuizEditor === "true") {
      closeQuizEditorModal();
    }
  });

  els.quizAnswerForm.addEventListener("submit", (event) => {
    event.preventDefault();
    submitQuizAnswer();
  });

  els.quizEditorForm.addEventListener("submit", async (event) => {
    event.preventDefault();

    const video = state.videos.find((item) => item.id === state.quizEditorVideoId);
    if (!video) {
      return;
    }

    video.quiz = {
      question: els.quizEditorQuestion.value.trim(),
      options: [
        els.quizEditorOption1.value.trim(),
        els.quizEditorOption2.value.trim(),
        els.quizEditorOption3.value.trim(),
        els.quizEditorOption4.value.trim(),
      ],
      correctIndex: Number(els.quizEditorCorrect.value),
    };

    await persistVideos();
    state.videos = await getAllVideos();
    closeQuizEditorModal();
    renderAll();
  });

  els.deleteQuizBtn.addEventListener("click", async () => {
    const video = state.videos.find((item) => item.id === state.quizEditorVideoId);
    if (!video) {
      return;
    }

    delete video.quiz;
    await persistVideos();
    state.videos = await getAllVideos();
    closeQuizEditorModal();
    renderAll();
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !els.analyticsModal.classList.contains("hidden")) {
      closeAnalyticsModal();
    }
    if (event.key === "Escape" && !els.quizModal.classList.contains("hidden")) {
      closeQuizModal();
    }
    if (event.key === "Escape" && !els.quizEditorModal.classList.contains("hidden")) {
      closeQuizEditorModal();
    }
  });
}

function syncAnalyticsModalStyles() {
  syncModalStyles(els.analyticsModal, els.analyticsModalBackdrop, els.analyticsModalCard, els.analyticsModalClose);
}

function syncModalStyles(modal, backdrop, card, closeButton) {
  Object.assign(modal.style, {
    position: "fixed",
    inset: "0",
    zIndex: "60",
    display: "grid",
    placeItems: "center",
    padding: "1rem"
  });

  Object.assign(backdrop.style, {
    position: "absolute",
    inset: "0",
    background: "rgba(15, 22, 34, 0.58)"
  });

  Object.assign(card.style, {
    position: "relative",
    width: "min(860px, calc(100vw - 2rem))",
    maxHeight: "calc(100vh - 3rem)",
    overflow: "auto",
    borderRadius: "22px",
    background: "#fff",
    boxShadow: "0 30px 80px rgba(0, 0, 0, 0.28)"
  });

  Object.assign(closeButton.style, {
    border: "0",
    background: "transparent",
    color: "#6b7280",
    fontSize: "2.2rem",
    lineHeight: "1",
    padding: "0",
    cursor: "pointer"
  });
}

function wirePlayerTracking() {
  els.trainingPlayer.addEventListener("timeupdate", () => {
    const userId = state.activeTraineeId;
    const videoId = currentVideoId();
    const player = els.trainingPlayer;

    if (!userId || !videoId || !Number.isFinite(player.duration) || player.duration <= 0) {
      return;
    }

    const key = progressKey(userId, videoId);
    const entry = state.progress[key] || createEmptyProgressEntry();

    entry.watchedSeconds = Math.max(entry.watchedSeconds, player.currentTime);
    entry.duration = player.duration;
    entry.lastWatchedAt = new Date().toISOString();

    state.progress[key] = entry;
    saveProgress(state.progress);

    updateCompletionControls(entry);
    renderPlayerStatus();
  });

  els.trainingPlayer.addEventListener("ended", () => {
    const userId = state.activeTraineeId;
    const videoId = currentVideoId();
    const player = els.trainingPlayer;

    if (!userId || !videoId || !Number.isFinite(player.duration) || player.duration <= 0) {
      return;
    }

    const key = progressKey(userId, videoId);
    const entry = state.progress[key] || createEmptyProgressEntry();

    entry.watchedSeconds = Math.max(entry.watchedSeconds, player.duration);
    entry.duration = player.duration;
    entry.viewCount = (entry.viewCount || 0) + 1;
    entry.lastWatchedAt = new Date().toISOString();

    state.progress[key] = entry;
    saveProgress(state.progress);

    updateCompletionControls(entry);
    renderVideoCards();
    renderPlayerStatus();
    renderProgressTable();
    renderAnalyticsPanel();
    openQuizModal(videoId, true);
  });

  els.trainingPlayer.addEventListener("loadedmetadata", () => {
    restorePlaybackPosition();
    renderPlayerStatus();
  });
}

function renderAll() {
  renderTraineeVisibility();
  renderFolderFilterOptions();
  updateVideoFolderOptions();
  updateVideoFolderInputVisibility();
  renderVideoCards();
  renderAdminData();

  if (state.activeTraineeId && !currentVideoId() && state.videos.length > 0) {
    selectVideo(state.videos[0].id);
  } else {
    refreshPlayerPanel();
  }
}

function renderVideoCards() {
  if (!state.activeTraineeId) {
    els.videoCardGrid.innerHTML = "";
    return;
  }

  els.videoCardGrid.className = "folder-groups";
  Object.assign(els.videoCardGrid.style, {
    display: "grid",
    gridTemplateColumns: "1fr",
    gap: "1.4rem",
    width: "100%"
  });

  if (state.videos.length === 0) {
    els.videoCardGrid.innerHTML = '<article class="panel">No training videos uploaded yet.</article>';
    return;
  }

  const selectedVideoId = currentVideoId();

  const filtered = state.videos.filter((video) => {
    if (state.folderFilter !== "all" && getVideoFolder(video) !== state.folderFilter) {
      return false;
    }

    if (!state.searchQuery) {
      return true;
    }
    const text = `${video.title} ${video.description || ""} ${getVideoFolder(video)}`.toLowerCase();
    return text.includes(state.searchQuery);
  });

  if (filtered.length === 0) {
    els.videoCardGrid.innerHTML = '<article class="panel">No modules match this search.</article>';
    return;
  }

  const grouped = filtered.reduce((accumulator, video) => {
    const folder = getVideoFolder(video);
    accumulator[folder] = accumulator[folder] || [];
    accumulator[folder].push(video);
    return accumulator;
  }, {});

  els.videoCardGrid.innerHTML = Object.entries(grouped)
    .map(([folder, videos]) => {
      const cards = videos
        .map((video, index) => {
          const entry = state.progress[progressKey(state.activeTraineeId, video.id)] || createEmptyProgressEntry();
          const completedMark = entry.manualCompleted ? '<span class="done-badge">✓</span>' : "";
          const activeStyle = selectedVideoId === video.id ? ' style="outline: 2px solid #1496da;"' : "";
          return `
            <article class="video-card"${activeStyle}>
              <div class="card-thumb">
                ▶
                ${completedMark}
              </div>
              <div class="card-body">
                <span class="folder-chip">${escapeHtml(folder)}</span>
                <p class="card-title">${escapeHtml(video.title)}</p>
                <div class="card-meta">Module ${index + 1} • ${formatFileSize(video.size)} • Watched ${entry.viewCount || 0}x</div>
                <button class="watch-btn" data-video-id="${video.id}" type="button">Watch Now</button>
              </div>
            </article>
          `;
        })
        .join("");

      return `
        <section class="folder-section" style="display:block;width:100%;grid-column:1 / -1;">
          <div class="folder-heading">
            <h3>${escapeHtml(folder)}</h3>
            <span class="folder-count">${videos.length} video${videos.length === 1 ? "" : "s"}</span>
          </div>
          <div class="folder-video-grid" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:1.15rem;width:100%;">
            ${cards}
          </div>
        </section>
      `;
    })
    .join("");

  els.videoCardGrid.querySelectorAll("[data-video-id]").forEach((btn) => {
    btn.addEventListener("click", () => {
      selectVideo(btn.dataset.videoId);
    });
  });
}

function selectVideo(videoId) {
  if (!state.activeTraineeId) {
    return;
  }

  const video = state.videos.find((item) => item.id === videoId);
  if (!video) {
    return;
  }

  const oldSrc = els.trainingPlayer.getAttribute("src");
  if (oldSrc && oldSrc.startsWith("blob:")) {
    URL.revokeObjectURL(oldSrc);
  }

  els.trainingPlayer.src = URL.createObjectURL(video.blob);
  els.trainingPlayer.dataset.videoId = video.id;
  els.trainingPlayer.load();

  refreshPlayerPanel();
  renderVideoCards();
}

function refreshPlayerPanel() {
  if (!state.activeTraineeId) {
    clearPlayerMessages();
    return;
  }

  const videoId = currentVideoId();
  const video = state.videos.find((item) => item.id === videoId);

  if (!video) {
    clearPlayerMessages();
    return;
  }

  els.playerTitle.textContent = video.title;
  els.videoMeta.textContent = video.description || "No description provided.";

  const entry = state.progress[progressKey(state.activeTraineeId, video.id)] || createEmptyProgressEntry();
  updateCompletionControls(entry);
  renderPlayerStatus();
}

function clearPlayerMessages() {
  els.playerTitle.textContent = "Select a module";
  els.videoMeta.textContent = state.activeTraineeId
    ? "Choose a card to start watching."
    : "Login as trainee to access your modules.";
  els.completionCheckbox.checked = false;
  els.completionCheckbox.disabled = true;
  els.completionNote.textContent = "Watch until near the end to enable completion checkbox.";
  els.videoStatus.textContent = "Progress will appear here once playback starts.";
}

function clearPlayer() {
  const oldSrc = els.trainingPlayer.getAttribute("src");
  if (oldSrc && oldSrc.startsWith("blob:")) {
    URL.revokeObjectURL(oldSrc);
  }
  els.trainingPlayer.removeAttribute("src");
  els.trainingPlayer.removeAttribute("data-video-id");
  els.trainingPlayer.load();
  closeQuizModal();
}

function updateCompletionControls(entry) {
  const video = state.videos.find((item) => item.id === currentVideoId());
  const enabled = isEligibleForManualCompletion(entry, video);
  els.completionCheckbox.disabled = !enabled;
  els.completionCheckbox.checked = Boolean(entry.manualCompleted);
  els.openQuizBtn.classList.toggle("hidden", !video?.quiz);
  els.completionNote.textContent = enabled
    ? "You can now check completion for this module."
    : video?.quiz
      ? "Watch the video and pass the quiz to unlock completion."
      : "Watch at least 90% of the video to enable completion checkbox.";
}

function renderPlayerStatus() {
  const userId = state.activeTraineeId;
  const videoId = currentVideoId();

  if (!userId || !videoId) {
    els.videoStatus.textContent = "Select a module to begin.";
    return;
  }

  const entry = state.progress[progressKey(userId, videoId)] || createEmptyProgressEntry();
  const duration = entry.duration || els.trainingPlayer.duration || 0;
  const ratio = duration > 0 ? Math.min(100, (entry.watchedSeconds / duration) * 100) : 0;

  if (entry.manualCompleted) {
    const completionDate = entry.completedAt ? new Date(entry.completedAt).toLocaleString() : "just now";
    els.videoStatus.textContent = `Completed by trainee. Progress ${ratio.toFixed(0)}%. Last completion: ${completionDate}.`;
    return;
  }

  if (videoHasQuiz(videoId) && !entry.quizPassed) {
    els.videoStatus.textContent = `In progress ${ratio.toFixed(0)}%. Finish the video and pass the quiz to complete it.`;
    return;
  }

  els.videoStatus.textContent = `In progress ${ratio.toFixed(0)}%. Full watches: ${entry.viewCount || 0}.`;
}

function restorePlaybackPosition() {
  const userId = state.activeTraineeId;
  const videoId = currentVideoId();
  const player = els.trainingPlayer;

  if (!userId || !videoId || !Number.isFinite(player.duration) || player.duration <= 0) {
    return;
  }

  const entry = state.progress[progressKey(userId, videoId)];
  if (!entry) {
    return;
  }

  const maxSeek = Math.max(player.duration - 1, 0);
  player.currentTime = Math.min(entry.watchedSeconds || 0, maxSeek);
}

function renderAdminData() {
  renderJoineeList();
  renderVideoLibrary();
  renderProgressTable();
  renderAnalyticsPanel();
  renderExcelAnalytics();
}

function renderJoineeList() {
  if (state.users.length === 0) {
    els.joineeList.innerHTML = "<li>No joinees added yet.</li>";
    return;
  }

  els.joineeList.innerHTML = state.users
    .map(
      (user) =>
        `
          <li>
            <article class="employee-card">
              <div class="employee-card-header">
                <div>
                  <strong>${escapeHtml(user.name)}</strong><br />
                  <span>${escapeHtml(user.email)}</span>
                </div>
                <div class="employee-card-actions">
                  <span class="pill ${user.status === "active" ? "pill-active" : "pill-disabled"}">${escapeHtml(user.status)}</span>
                  ${
                    user.mustChangePassword
                      ? '<span class="pill pill-temp">Temporary password in use</span>'
                      : ""
                  }
                </div>
              </div>
              <div>Testing password: <strong>${escapeHtml(user.password)}</strong></div>
              <div class="employee-inline-fields">
                <input data-reset-password="${user.id}" type="text" placeholder="New temporary password" minlength="8" />
                <button class="secondary-btn" data-reset-user="${user.id}" type="button">Reset Password</button>
                <button class="${user.status === "active" ? "danger-btn" : "secondary-btn"}" data-toggle-user="${user.id}" type="button">
                  ${user.status === "active" ? "Disable Access" : "Enable Access"}
                </button>
              </div>
            </article>
          </li>
        `
    )
    .join("");

  els.joineeList.querySelectorAll("[data-toggle-user]").forEach((button) => {
    button.addEventListener("click", () => {
      const user = state.users.find((item) => item.id === button.dataset.toggleUser);
      if (!user) {
        return;
      }

      user.status = user.status === "active" ? "disabled" : "active";

      if (user.status === "disabled" && state.activeTraineeId === user.id) {
        state.activeTraineeId = "";
        localStorage.removeItem(STORAGE_KEYS.traineeSessionUserId);
        clearPlayer();
      }

      saveUsers(state.users);
      updateNavigationVisibility();
      renderAll();
    });
  });

  els.joineeList.querySelectorAll("[data-reset-user]").forEach((button) => {
    button.addEventListener("click", () => {
      const userId = button.dataset.resetUser;
      const passwordInput = els.joineeList.querySelector(`[data-reset-password="${userId}"]`);
      const user = state.users.find((item) => item.id === userId);

      if (!user || !passwordInput) {
        return;
      }

      const nextPassword = passwordInput.value.trim();
      if (nextPassword.length < 8) {
        alert("Use at least 8 characters for the temporary password.");
        return;
      }

      user.password = nextPassword;
      user.mustChangePassword = true;
      user.status = "active";
      saveUsers(state.users);

      if (state.activeTraineeId === user.id) {
        state.activeTraineeId = "";
        localStorage.removeItem(STORAGE_KEYS.traineeSessionUserId);
        clearPlayer();
      }

      updateNavigationVisibility();
      renderAll();
      alert(`Temporary password reset for ${user.name}.`);
    });
  });
}

function renderVideoLibrary() {
  if (state.videos.length === 0) {
    els.videoList.innerHTML = "<li>No training videos uploaded yet.</li>";
    return;
  }

  const folders = [...new Set(state.videos.map((video) => getVideoFolder(video)))].sort((left, right) =>
    left.localeCompare(right)
  );

  els.videoList.innerHTML = state.videos
    .map(
      (video, index) =>
        `
          <li>
            <article class="video-library-item">
              <div>
                <strong>Module ${index + 1}: ${escapeHtml(video.title)}</strong><br />
                <span>Folder: ${escapeHtml(getVideoFolder(video))}</span><br />
                <span>${escapeHtml(video.fileName)} • ${formatFileSize(video.size)}${state.videoStorageMode === "memory" ? " • session only" : ""}</span>
              </div>
              <div class="video-library-actions">
                <select data-folder-select="${video.id}">
                  <option value="__new__">Create new folder</option>
                  ${folders
                    .map(
                      (folder) =>
                        `<option value="${escapeHtml(folder)}" ${getVideoFolder(video) === folder ? "selected" : ""}>${escapeHtml(folder)}</option>`
                    )
                    .join("")}
                </select>
                <input data-folder-input="${video.id}" type="text" placeholder="New folder name" class="hidden" />
                <button class="secondary-btn" data-update-folder="${video.id}" type="button">Update Folder</button>
                <button class="danger-btn" data-delete-video="${video.id}" type="button">Delete Video</button>
              </div>
            </article>
          </li>
        `
    )
    .join("");

  els.videoList.querySelectorAll("[data-folder-select]").forEach((select) => {
    select.addEventListener("change", () => {
      const folderInput = els.videoList.querySelector(`[data-folder-input="${select.dataset.folderSelect}"]`);
      if (!folderInput) {
        return;
      }

      folderInput.classList.toggle("hidden", select.value !== "__new__");
    });
  });

  els.videoList.querySelectorAll("[data-update-folder]").forEach((button) => {
    button.addEventListener("click", async () => {
      const videoId = button.dataset.updateFolder;
      const select = els.videoList.querySelector(`[data-folder-select="${videoId}"]`);
      const input = els.videoList.querySelector(`[data-folder-input="${videoId}"]`);
      const video = state.videos.find((item) => item.id === videoId);

      if (!select || !input || !video) {
        return;
      }

      const nextFolder = select.value === "__new__" ? input.value.trim() : select.value.trim();

      if (!nextFolder) {
        alert("Enter a folder name before updating.");
        return;
      }

      video.folder = nextFolder;
      await persistVideos();
      state.videos = await getAllVideos();
      renderAll();
    });
  });

  els.videoList.querySelectorAll("[data-delete-video]").forEach((button) => {
    button.addEventListener("click", async () => {
      const videoId = button.dataset.deleteVideo;
      const video = state.videos.find((item) => item.id === videoId);

      if (!video) {
        return;
      }

      const confirmed = window.confirm(`Delete "${video.title}" from this testing portal?`);
      if (!confirmed) {
        return;
      }

      await deleteVideo(videoId);
      state.videos = await getAllVideos();

      Object.keys(state.progress).forEach((key) => {
        if (key.endsWith(`::${videoId}`)) {
          delete state.progress[key];
        }
      });
      saveProgress(state.progress);

      if (currentVideoId() === videoId) {
        clearPlayer();
      }

      renderAll();
    });
  });
}

function renderProgressTable() {
  if (state.users.length === 0) {
    els.progressTableBody.innerHTML = "<tr><td colspan='6'>No joinees available.</td></tr>";
    return;
  }

  const totalVideos = state.videos.length;

  els.progressTableBody.innerHTML = state.users
    .map((user) => {
      let completedCount = 0;
      let totalViews = 0;
      let latestCompletion = null;

      state.videos.forEach((video) => {
        const entry = state.progress[progressKey(user.id, video.id)];
        if (!entry) {
          return;
        }

        totalViews += entry.viewCount || 0;
        if (entry.manualCompleted) {
          completedCount += 1;
          if (entry.completedAt && (!latestCompletion || entry.completedAt > latestCompletion)) {
            latestCompletion = entry.completedAt;
          }
        }
      });

      return `
        <tr>
          <td>${escapeHtml(user.name)}</td>
          <td>${escapeHtml(user.email)}</td>
          <td>${escapeHtml(user.status)}</td>
          <td>${completedCount}/${totalVideos}</td>
          <td>${totalViews}</td>
          <td>${latestCompletion ? new Date(latestCompletion).toLocaleString() : "-"}</td>
        </tr>
      `;
    })
    .join("");
}

function renderAnalyticsPanel() {
  if (state.users.length === 0) {
    els.analyticsSummary.innerHTML = "";
    els.analyticsCards.innerHTML = '<div class="analytics-empty">No analytics available yet. Add an employee first.</div>';
    return;
  }

  const analytics = state.users.map(getEmployeeAnalyticsData);
  const totalEmployees = analytics.length;
  const totalCompletions = analytics.reduce((sum, item) => sum + item.completed, 0);
  const totalVideos = state.videos.length;
  const averageCompletion = totalEmployees > 0 ? Math.round(analytics.reduce((sum, item) => sum + item.rate, 0) / totalEmployees) : 0;

  els.analyticsSummary.innerHTML = `
    <article class="analytics-stat blue">
      <p class="metric-label">Total Employees</p>
      <p class="metric-value">${totalEmployees}</p>
    </article>
    <article class="analytics-stat cyan">
      <p class="metric-label">Total Completions</p>
      <p class="metric-value">${totalCompletions}</p>
    </article>
    <article class="analytics-stat orange">
      <p class="metric-label">Average Completion</p>
      <p class="metric-value">${averageCompletion}%</p>
    </article>
    <article class="analytics-stat green">
      <p class="metric-label">Total Videos</p>
      <p class="metric-value">${totalVideos}</p>
    </article>
  `;

  const filtered = analytics.filter((item) => {
    if (!state.analyticsSearchQuery) {
      return true;
    }

    const haystack = `${item.user.name} ${item.user.email}`.toLowerCase();
    return haystack.includes(state.analyticsSearchQuery);
  });

  if (filtered.length === 0) {
    els.analyticsCards.innerHTML = '<div class="analytics-empty">No employees match that search.</div>';
    return;
  }

  els.analyticsCards.innerHTML = filtered
    .map((item) => {
      const latestCompletion = item.completions[0];
      return `
        <article class="analytics-employee-card">
          <div class="analytics-employee-head">
            <div>
              <h3 class="analytics-employee-name">${escapeHtml(item.user.name)}</h3>
              <p class="analytics-employee-email">${escapeHtml(item.user.email)}</p>
            </div>
            <div class="analytics-percent">${item.rate}%</div>
          </div>

          <div class="analytics-progress" aria-hidden="true">
            <div class="analytics-progress-bar" style="width: ${item.rate}%"></div>
          </div>

          <p class="analytics-copy">${item.completed} of ${item.totalModules} videos completed</p>

          <div class="analytics-recent">
            <h3>Recent Completions</h3>
            ${
              latestCompletion
                ? renderCompletionItem(latestCompletion)
                : '<p class="analytics-copy">No completed modules yet.</p>'
            }
          </div>

          <button class="analytics-details-btn" data-analytics-user="${item.user.id}" type="button">View Full Details</button>
        </article>
      `;
    })
    .join("");

  els.analyticsCards.querySelectorAll("[data-analytics-user]").forEach((button) => {
    button.addEventListener("click", () => {
      openAnalyticsModal(button.dataset.analyticsUser);
    });
  });
}

function renderExcelAnalytics() {
  const reports = [...state.excelReports].sort((left, right) => new Date(right.uploadedAt) - new Date(left.uploadedAt));
  const activeReport = getActiveExcelReport(reports);

  if (!state.activeExcelReportId && reports.length > 0) {
    state.activeExcelReportId = reports[0].id;
  }

  els.excelReportSelect.innerHTML = reports.length
    ? reports
        .map(
          (report) =>
            `<option value="${report.id}" ${report.id === state.activeExcelReportId ? "selected" : ""}>${escapeHtml(report.fileName)}</option>`
        )
        .join("")
    : '<option value="">No uploaded reports yet</option>';

  if (!activeReport) {
    els.excelSheetSelect.innerHTML = '<option value="">Select a report first</option>';
    els.excelCategorySelect.innerHTML = '<option value="">Select a sheet first</option>';
    els.excelValueSelect.innerHTML = '<option value="__row_count__">Using row count</option>';
    els.excelSheetSelect.disabled = true;
    els.excelCategorySelect.disabled = true;
    els.excelValueSelect.disabled = true;
    els.excelEmptyState.classList.remove("hidden");
    els.excelResults.classList.add("hidden");
    els.excelReportMeta.classList.add("hidden");
    els.excelTableBody.innerHTML = "";
    els.excelLegend.innerHTML = "";
    els.excelPieChart.style.background =
      "radial-gradient(circle at center, rgba(255, 255, 255, 0.2) 0 36%, transparent 37%), conic-gradient(#d7dee9 0deg 360deg)";
    els.excelStatus.textContent = "Upload an Excel sheet to start generating charts.";
    return;
  }

  if (!activeReport.activeSheetName || !activeReport.sheets.some((sheet) => sheet.name === activeReport.activeSheetName)) {
    activeReport.activeSheetName = activeReport.sheets[0]?.name || "";
    saveExcelReports(state.excelReports);
  }

  els.excelReportMeta.classList.remove("hidden");
  els.excelReportMeta.innerHTML = `
    <strong>${escapeHtml(activeReport.fileName)}</strong>
    <div>${activeReport.sheets.length} sheet${activeReport.sheets.length === 1 ? "" : "s"} • Uploaded ${new Date(activeReport.uploadedAt).toLocaleString()}</div>
  `;

  els.excelSheetSelect.disabled = false;
  els.excelSheetSelect.innerHTML = activeReport.sheets
    .map(
      (sheet) =>
        `<option value="${escapeHtml(sheet.name)}" ${sheet.name === activeReport.activeSheetName ? "selected" : ""}>${escapeHtml(sheet.name)}</option>`
    )
    .join("");

  const activeSheet = getActiveExcelSheet(activeReport);
  if (!activeSheet) {
    els.excelStatus.textContent = "This report does not contain a usable sheet yet.";
    return;
  }

  applySuggestedSheetConfig(activeSheet);
  saveExcelReports(state.excelReports);

  els.excelCategorySelect.disabled = false;
  els.excelCategorySelect.innerHTML = activeSheet.headers
    .map(
      (header) =>
        `<option value="${escapeHtml(header)}" ${header === activeSheet.config.categoryColumn ? "selected" : ""}>${escapeHtml(header)}</option>`
    )
    .join("");

  const valueOptions = [`<option value="__row_count__" ${!activeSheet.config.valueColumn ? "selected" : ""}>Use row count</option>`].concat(
    activeSheet.numericHeaders.map(
      (header) =>
        `<option value="${escapeHtml(header)}" ${header === activeSheet.config.valueColumn ? "selected" : ""}>${escapeHtml(header)}</option>`
    )
  );
  els.excelValueSelect.disabled = false;
  els.excelValueSelect.innerHTML = valueOptions.join("");

  const summary = summarizeExcelSheet(activeSheet);
  updateExcelStatusMessage(activeSheet, summary);

  if (!summary.length) {
    els.excelEmptyState.classList.remove("hidden");
    els.excelResults.classList.add("hidden");
    els.excelEmptyState.textContent =
      "This sheet could not produce a pie chart. Make sure it has at least one category column and either a numeric column or repeated categories.";
    return;
  }

  els.excelEmptyState.classList.add("hidden");
  els.excelResults.classList.remove("hidden");
  renderExcelSummary(summary, activeSheet);
}

function renderExcelSummary(summary, sheet) {
  const measurementLabel = sheet.config.valueColumn
    ? `Sum of ${sheet.config.valueColumn}`
    : "Row count";
  const total = summary.reduce((sum, item) => sum + item.value, 0);
  const topItem = summary[0];
  const chartStops = [];
  let currentStop = 0;

  summary.forEach((item) => {
    const nextStop = currentStop + item.percent * 3.6;
    chartStops.push(`${item.color} ${currentStop}deg ${nextStop}deg`);
    currentStop = nextStop;
  });

  els.excelChartTitle.textContent = `Pie Chart • ${measurementLabel}`;
  els.excelChartHeadline.textContent = `${topItem.label} leads with ${formatSummaryValue(topItem.value)} (${topItem.percent.toFixed(1)}%)`;
  els.excelChartDescription.textContent = `${summary.length} segments created from "${sheet.name}" using "${sheet.config.categoryColumn}". Total ${measurementLabel.toLowerCase()}: ${formatSummaryValue(total)}.`;
  els.excelPieChart.style.background =
    `radial-gradient(circle at center, rgba(255, 255, 255, 0.2) 0 36%, transparent 37%), conic-gradient(${chartStops.join(", ")})`;

  els.excelLegend.innerHTML = summary
    .map(
      (item) => `
        <div class="excel-legend-item">
          <div class="excel-legend-copy">
            <span class="excel-legend-swatch" style="background:${item.color}"></span>
            <span class="excel-legend-label">${escapeHtml(item.label)}</span>
          </div>
          <div class="excel-legend-value">${formatSummaryValue(item.value)} • ${item.percent.toFixed(1)}%</div>
        </div>
      `
    )
    .join("");

  els.excelTableBody.innerHTML = summary
    .map(
      (item) => `
        <tr>
          <td>${escapeHtml(item.label)}</td>
          <td>${formatSummaryValue(item.value)}</td>
          <td>${item.percent.toFixed(1)}%</td>
        </tr>
      `
    )
    .join("");
}

function getActiveExcelReport(reportList = state.excelReports) {
  return reportList.find((report) => report.id === state.activeExcelReportId) || reportList[0] || null;
}

function getActiveExcelSheet(report = getActiveExcelReport()) {
  if (!report) {
    return null;
  }

  return report.sheets.find((sheet) => sheet.name === report.activeSheetName) || report.sheets[0] || null;
}

function applySuggestedSheetConfig(sheet) {
  sheet.config = sheet.config || {};

  if (!sheet.config.categoryColumn || !sheet.headers.includes(sheet.config.categoryColumn)) {
    sheet.config.categoryColumn = suggestCategoryColumn(sheet);
  }

  if (sheet.config.valueColumn && !sheet.numericHeaders.includes(sheet.config.valueColumn)) {
    sheet.config.valueColumn = "";
  }

  if (sheet.config.valueColumn === null || typeof sheet.config.valueColumn === "undefined") {
    sheet.config.valueColumn = suggestValueColumn(sheet);
  }
}

async function importExcelReport(file) {
  if (!window.XLSX) {
    els.excelStatus.textContent = "Excel parser failed to load. Refresh the page and try again.";
    return;
  }

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheets = workbook.SheetNames.map((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: "", raw: false });
    return buildExcelSheet(sheetName, rawRows);
  }).filter((sheet) => sheet.headers.length > 0 && sheet.rows.length > 0);

  if (!sheets.length) {
    els.excelStatus.textContent = "No usable tabular data was found in that file.";
    return;
  }

  const report = {
    id: crypto.randomUUID(),
    fileName: file.name,
    uploadedAt: new Date().toISOString(),
    activeSheetName: sheets[0].name,
    sheets,
  };

  state.excelReports = [report, ...state.excelReports.filter((item) => item.fileName !== file.name)];
  state.activeExcelReportId = report.id;
  saveExcelReports(state.excelReports);
  els.excelUploadForm.reset();
  renderExcelAnalytics();
}

function buildExcelSheet(name, rows) {
  const normalizedRows = rows
    .map((row) => {
      const nextRow = {};
      Object.entries(row).forEach(([key, value]) => {
        const normalizedKey = String(key || "").trim();
        if (!normalizedKey) {
          return;
        }
        nextRow[normalizedKey] = typeof value === "string" ? value.trim() : value;
      });
      return nextRow;
    })
    .filter((row) => Object.keys(row).length > 0);

  const headers = [...new Set(normalizedRows.flatMap((row) => Object.keys(row)))];
  const numericHeaders = headers.filter((header) => {
    const numericCount = normalizedRows.reduce((count, row) => {
      return Number.isFinite(parseNumericValue(row[header])) ? count + 1 : count;
    }, 0);
    return numericCount > 0;
  });

  const sheet = {
    name,
    headers,
    numericHeaders,
    rows: normalizedRows,
    config: {
      categoryColumn: "",
      valueColumn: null,
    },
  };

  applySuggestedSheetConfig(sheet);
  return sheet;
}

function summarizeExcelSheet(sheet) {
  const categoryColumn = sheet.config?.categoryColumn;
  if (!categoryColumn) {
    return [];
  }

  const bucket = new Map();
  for (const row of sheet.rows) {
    const rawLabel = row[categoryColumn];
    const label = String(rawLabel || "").trim() || "Blank";
    const value = sheet.config.valueColumn ? parseNumericValue(row[sheet.config.valueColumn]) : 1;
    if (!Number.isFinite(value)) {
      continue;
    }
    bucket.set(label, (bucket.get(label) || 0) + value);
  }

  const sortedEntries = [...bucket.entries()]
    .map(([label, value]) => ({ label, value }))
    .filter((item) => item.value > 0)
    .sort((left, right) => right.value - left.value);

  const limitedEntries = sortedEntries.slice(0, 7);
  const remainingTotal = sortedEntries.slice(7).reduce((sum, item) => sum + item.value, 0);
  if (remainingTotal > 0) {
    limitedEntries.push({ label: "Other", value: remainingTotal });
  }

  const entries = limitedEntries.map((item, index) => ({ ...item, color: chartColorAt(index) }));

  const total = entries.reduce((sum, item) => sum + item.value, 0);
  return total > 0 ? entries.map((item) => ({ ...item, percent: (item.value / total) * 100 })) : [];
}

function suggestCategoryColumn(sheet) {
  const scored = sheet.headers
    .map((header) => {
      const values = sheet.rows.map((row) => String(row[header] || "").trim()).filter(Boolean);
      const uniqueCount = new Set(values).size;
      const textCount = values.filter((value) => !Number.isFinite(parseNumericValue(value))).length;
      const headerName = header.toLowerCase();
      const priorityBoost = /(category|department|team|region|status|type|segment|group|name)/.test(headerName) ? 3 : 0;
      return {
        header,
        score: priorityBoost + textCount + (values.length > uniqueCount ? 2 : 0),
      };
    })
    .sort((left, right) => right.score - left.score);

  return scored[0]?.header || sheet.headers[0] || "";
}

function suggestValueColumn(sheet) {
  const scored = sheet.numericHeaders
    .map((header) => {
      const headerName = header.toLowerCase();
      const priorityBoost = /(amount|sales|revenue|count|total|score|value|qty|quantity|price)/.test(headerName) ? 3 : 0;
      return {
        header,
        score: priorityBoost,
      };
    })
    .sort((left, right) => right.score - left.score);

  return scored[0]?.header || "";
}

function updateExcelStatusMessage(sheet, summary = summarizeExcelSheet(sheet)) {
  const valueMode = sheet.config.valueColumn
    ? `summing "${sheet.config.valueColumn}"`
    : "counting rows";

  els.excelStatus.textContent = summary.length
    ? `Showing ${summary.length} pie segments from "${sheet.name}" using "${sheet.config.categoryColumn}" and ${valueMode}.`
    : `Ready to analyze "${sheet.name}". Pick a category column and optionally a numeric value column.`;
}

function parseNumericValue(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : NaN;
  }

  const normalized = String(value || "")
    .replaceAll(",", "")
    .replace(/[^\d.-]/g, "")
    .trim();
  if (!normalized) {
    return NaN;
  }

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : NaN;
}

function chartColorAt(index) {
  const palette = ["#0f88d6", "#13b5d8", "#34c759", "#f59e0b", "#ef4444", "#8b5cf6", "#14b8a6", "#f97316"];
  return palette[index % palette.length];
}

function formatSummaryValue(value) {
  if (!Number.isFinite(value)) {
    return "0";
  }

  const hasDecimals = Math.abs(value % 1) > 0.001;
  return new Intl.NumberFormat(undefined, {
    maximumFractionDigits: hasDecimals ? 2 : 0,
  }).format(value);
}

function getEmployeeAnalyticsData(user) {
  const totalModules = state.videos.length;
  const completions = state.videos
    .map((video) => {
      const entry = state.progress[progressKey(user.id, video.id)] || createEmptyProgressEntry();
      return {
        video,
        entry,
      };
    })
    .filter((item) => item.entry.manualCompleted)
    .sort((a, b) => new Date(b.entry.completedAt || 0) - new Date(a.entry.completedAt || 0));

  const completed = completions.length;
  const rate = totalModules > 0 ? Math.round((completed / totalModules) * 100) : 0;

  return {
    user,
    totalModules,
    completed,
    rate,
    completions,
  };
}

function renderCompletionItem(item) {
  const title = item.video.fileName || item.video.title;
  return `
    <div class="completion-item">
      <span class="completion-dot">✓</span>
      <div>
        <strong>${escapeHtml(title)}</strong>
        <div class="completion-meta">${escapeHtml(getVideoFolder(item.video))}</div>
        <div class="completion-meta">${escapeHtml(item.video.title || "")}</div>
        <div class="completion-meta">${item.entry.completedAt ? `Completed: ${new Date(item.entry.completedAt).toLocaleString()}` : ""}</div>
      </div>
    </div>
  `;
}

function renderFolderFilterOptions() {
  const folders = [...new Set(state.videos.map((video) => getVideoFolder(video)))].sort((left, right) =>
    left.localeCompare(right)
  );

  if (state.folderFilter !== "all" && !folders.includes(state.folderFilter)) {
    state.folderFilter = "all";
  }

  els.folderFilterSelect.innerHTML = ['<option value="all">All folders</option>']
    .concat(folders.map((folder) => `<option value="${escapeHtml(folder)}">${escapeHtml(folder)}</option>`))
    .join("");

  els.folderFilterSelect.value = state.folderFilter;
}

function updateVideoFolderOptions() {
  const folders = [...new Set(state.videos.map((video) => getVideoFolder(video)))].sort((left, right) =>
    left.localeCompare(right)
  );

  const currentValue = els.videoFolderSelect.value || "__new__";

  els.videoFolderSelect.innerHTML = ['<option value="__new__">Create new folder</option>']
    .concat(folders.map((folder) => `<option value="${escapeHtml(folder)}">${escapeHtml(folder)}</option>`))
    .join("");

  if (currentValue !== "__new__" && folders.includes(currentValue)) {
    els.videoFolderSelect.value = currentValue;
  } else {
    els.videoFolderSelect.value = "__new__";
  }
}

function updateVideoFolderInputVisibility() {
  const isCreatingNew = els.videoFolderSelect.value === "__new__";
  els.videoFolderNewWrap.classList.toggle("hidden", !isCreatingNew);
  els.videoFolderNew.required = isCreatingNew;
  if (!isCreatingNew) {
    els.videoFolderNew.value = "";
  }
}

function getVideoFolder(video) {
  return (video.folder || "General").trim() || "General";
}

function openAnalyticsModal(userId) {
  const user = state.users.find((item) => item.id === userId);
  if (!user) {
    return;
  }

  const analytics = getEmployeeAnalyticsData(user);
  els.analyticsModalTitle.textContent = `${user.name} - Completion History`;
  els.analyticsModalBody.innerHTML = analytics.completions.length
    ? analytics.completions
        .map(
          (item) => `
            <article class="modal-completion-card">
              ${renderCompletionItem(item)}
            </article>
          `
        )
        .join("")
    : '<div class="analytics-empty">No completion history yet for this employee.</div>';

  els.analyticsModal.classList.remove("hidden");
  document.body.style.overflow = "hidden";
}

function closeAnalyticsModal() {
  els.analyticsModal.classList.add("hidden");
  document.body.style.overflow = "";
}

function getActiveTrainee() {
  return state.users.find((user) => user.id === state.activeTraineeId) || null;
}

function currentVideoId() {
  return els.trainingPlayer.dataset.videoId || "";
}

function isEligibleForManualCompletion(entry) {
  const playerDuration = Number.isFinite(els.trainingPlayer.duration) ? els.trainingPlayer.duration : 0;
  const duration = Math.max(entry.duration || 0, playerDuration);
  if (duration <= 0) {
    return false;
  }
  const ratio = (entry.watchedSeconds || 0) / duration;
  return ratio >= 0.9 || (entry.viewCount || 0) > 0;
}

function createEmptyProgressEntry() {
  return {
    watchedSeconds: 0,
    duration: 0,
    viewCount: 0,
    manualCompleted: false,
    completedAt: null,
    lastWatchedAt: null,
  };
}

function progressKey(userId, videoId) {
  return `${userId}::${videoId}`;
}

function applySavedTheme() {
  if (localStorage.getItem(STORAGE_KEYS.theme) === "dark") {
    document.body.classList.add("dark");
  }
}

function loadUsers() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.users);
    const users = raw ? JSON.parse(raw) : [];
    return users.map((user) => ({
      id: user.id || crypto.randomUUID(),
      name: user.name || "Unnamed",
      email: (user.email || "").toLowerCase(),
      password: user.password || "Trainee@123",
      status: user.status === "disabled" ? "disabled" : "active",
      mustChangePassword: Boolean(user.mustChangePassword),
      joinedAt: user.joinedAt || new Date().toISOString(),
    }));
  } catch {
    return [];
  }
}

function saveUsers(users) {
  localStorage.setItem(STORAGE_KEYS.users, JSON.stringify(users));
}

function loadProgress() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.progress);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function saveProgress(progress) {
  localStorage.setItem(STORAGE_KEYS.progress, JSON.stringify(progress));
}

function loadExcelReports() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.excelReports);
    const reports = raw ? JSON.parse(raw) : [];
    return reports.map((report) => ({
      ...report,
      sheets: (report.sheets || []).map((sheet) => ({
        ...sheet,
        headers: Array.isArray(sheet.headers) ? sheet.headers : [],
        numericHeaders: Array.isArray(sheet.numericHeaders) ? sheet.numericHeaders : [],
        rows: Array.isArray(sheet.rows) ? sheet.rows : [],
        config: {
          categoryColumn: sheet.config?.categoryColumn || "",
          valueColumn: sheet.config?.valueColumn ?? null,
        },
      })),
    }));
  } catch {
    return [];
  }
}

function saveExcelReports(reports) {
  localStorage.setItem(STORAGE_KEYS.excelReports, JSON.stringify(reports));
}

function openDatabase() {
  return new Promise((resolve, reject) => {
    if (!("indexedDB" in window)) {
      reject(new Error("IndexedDB is not supported."));
      return;
    }

    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(VIDEO_STORE)) {
        db.createObjectStore(VIDEO_STORE, { keyPath: "id" });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("IndexedDB failed to open."));
  });
}

async function saveVideo(video) {
  const db = await dbPromise;

  if (!db) {
    state.videoStorageMode = "memory";
    state.transientVideos = [video, ...state.transientVideos.filter((item) => item.id !== video.id)];
    reportRuntimeIssue(
      "warning",
      "This browser is using session-only video storage. Uploaded videos will disappear after refresh."
    );
    return;
  }

  await runTransaction(db, VIDEO_STORE, "readwrite", (store) => store.put(video));
}

async function deleteVideo(videoId) {
  const db = await dbPromise;

  if (!db) {
    state.transientVideos = state.transientVideos.filter((video) => video.id !== videoId);
    return;
  }

  await runTransaction(db, VIDEO_STORE, "readwrite", (store) => store.delete(videoId));
}

async function getAllVideos() {
  const db = await dbPromise;

  if (!db) {
    return state.transientVideos
      .map((video) => ({ ...video, folder: getVideoFolder(video) }))
      .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  }

  const videos = await runTransaction(db, VIDEO_STORE, "readonly", (store) => store.getAll());
  return videos
    .map((video) => ({ ...video, folder: getVideoFolder(video) }))
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
}

async function persistVideos() {
  const db = await dbPromise;

  if (!db) {
    state.transientVideos = [...state.videos];
    return;
  }

  await Promise.all(state.videos.map((video) => runTransaction(db, VIDEO_STORE, "readwrite", (store) => store.put(video))));
}

function runTransaction(db, storeName, mode, operation) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, mode);
    const store = tx.objectStore(storeName);
    const request = operation(store);

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    tx.onerror = () => reject(tx.error);
  });
}

function formatFileSize(sizeInBytes) {
  const mb = sizeInBytes / (1024 * 1024);
  return `${mb.toFixed(1)} MB`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
