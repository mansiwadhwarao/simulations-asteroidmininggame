/* Day 5: Asteroid Mining Rush & Resource Optimization - JavaScript Logic */

document.addEventListener("DOMContentLoaded", () => {
  
  // Safe wrapper for localStorage to prevent security exceptions in sandboxed environments
  const safeStorage = {
    getItem(key) {
      try {
        return localStorage.getItem(key);
      } catch (e) {
        console.warn("localStorage.getItem blocked:", e);
        return null;
      }
    },
    setItem(key, val) {
      try {
        localStorage.setItem(key, val);
      } catch (e) {
        console.warn("localStorage.setItem blocked:", e);
      }
    },
    removeItem(key) {
      try {
        localStorage.removeItem(key);
      } catch (e) {
        console.warn("localStorage.removeItem blocked:", e);
      }
    }
  };

  // ==========================================================================
  // 1. DATA MODEL & APPLICATION STATE
  // ==========================================================================
  
  // Model parameters (matching the Excel sheet cells)
  const excelModel = {
    startingMoney: 0.0,       // Cell E7
    initialTimer: 300,        // Cell E6 (seconds)
    travelCost: 100000.0,     // Cell E11
    
    excavatorCost: 5000.0,    // Cell E9
    fuelCost: 100.0,          // Cell E10
    excavationTimePerL: 10,   // Constant 10s of run time per L
    
    coinDropChance: 0.20,     // Cell D26
    pearlDropChance: 0.20,    // Cell D37
    crystalDropChance: 0.40,  // Cell D57
    
    drillCost: 50000.0,       // Cell D39
    gasCost: 200.0,           // Cell D40
    
    stormProtectionCost: 10000.0, // Cell D41
    drillBoreCost: 1000000.0, // Cell D59
    tyliumCost: 100000.0,     // Cell D60
    novaStormProtectionCost: 1000000.0, // Cell D61
    earthReturnCost: 50000000.0, // Cell D63
    
    // Titan Ores (Default values matching row 21-25)
    titanOres: [
      { id: "iron", name: "Iron Ore", cellProb: "D21", cellVal: "E21", cellAppr: "F21", cellEV: "H21", probability: 0.40, baseValue: 50.0, currentPrice: 50.0, appreciationRate: 0.0, icon: "⬜", color: "#f8fafc" },
      { id: "copper", name: "Copper Ore", cellProb: "D22", cellVal: "E22", cellAppr: "F22", cellEV: "H22", probability: 0.40, baseValue: 50.0, currentPrice: 50.0, appreciationRate: 0.0, icon: "🟫", color: "#f39c12" },
      { id: "gold", name: "Gold Ore", cellProb: "D23", cellVal: "E23", cellAppr: "F23", cellEV: "H23", probability: 0.10, baseValue: 200.0, currentPrice: 200.0, appreciationRate: 0.10, icon: "🟨", color: "#ffd700" },
      { id: "amethyst", name: "Amethyst Ore", cellProb: "D24", cellVal: "E24", cellAppr: "F24", cellEV: "H24", probability: 0.06, baseValue: 300.0, currentPrice: 300.0, appreciationRate: 0.20, icon: "🟪", color: "#d033ff" },
      { id: "ruby", name: "Ruby Ore", cellProb: "D25", cellVal: "E25", cellAppr: "F25", cellEV: "H25", probability: 0.04, baseValue: 500.0, currentPrice: 500.0, appreciationRate: 0.30, icon: "🟥", color: "#ff2d55" }
    ],
    
    // Void Ores (Default values matching row 32-36)
    voidOres: [
      { id: "bauxite", name: "Bauxite", cellProb: "D32", cellVal: "E32", cellAppr: "F32", cellEV: "H32", probability: 0.40, baseValue: 200.0, currentPrice: 200.0, appreciationRate: 0.0, icon: "⬜", color: "#b0c4de" },
      { id: "magnetite", name: "Magnetite", cellProb: "D33", cellVal: "E33", cellAppr: "F33", cellEV: "H33", probability: 0.40, baseValue: 200.0, currentPrice: 200.0, appreciationRate: 0.0, icon: "🟫", color: "#8b4513" },
      { id: "pyrite", name: "Pyrite", cellProb: "D34", cellVal: "E34", cellAppr: "F34", cellEV: "H34", probability: 0.10, baseValue: 1000.0, currentPrice: 1000.0, appreciationRate: 0.10, icon: "🟨", color: "#e5a93b" },
      { id: "diamond", name: "Diamond", cellProb: "D35", cellVal: "E35", cellAppr: "F35", cellEV: "H35", probability: 0.06, baseValue: 1200.0, currentPrice: 1200.0, appreciationRate: 0.20, icon: "💎", color: "#a0e6ff" },
      { id: "emerald", name: "Emerald", cellProb: "D36", cellVal: "E36", cellAppr: "F36", cellEV: "H36", probability: 0.04, baseValue: 5000.0, currentPrice: 5000.0, appreciationRate: 0.30, icon: "🟩", color: "#2ecc71" }
    ],

    // Nova Ores (Default values matching row 52-56)
    novaOres: [
      { id: "cryptonite", name: "Cryptonite", cellProb: "D52", cellVal: "E52", cellAppr: "F52", cellEV: "H52", probability: 0.40, baseValue: 10000.0, currentPrice: 10000.0, appreciationRate: 0.0, icon: "🟢", color: "#00e676" },
      { id: "elerium", name: "Elerium", cellProb: "D53", cellVal: "E53", cellAppr: "F53", cellEV: "H53", probability: 0.40, baseValue: 10000.0, currentPrice: 10000.0, appreciationRate: 0.0, icon: "🔵", color: "#00b0ff" },
      { id: "obsidian", name: "Obsidian", cellProb: "D54", cellVal: "E54", cellAppr: "F54", cellEV: "H54", probability: 0.10, baseValue: 100000.0, currentPrice: 100000.0, appreciationRate: 0.10, icon: "⚫", color: "#e2e8f0" },
      { id: "warpStone", name: "Warp-Stone", cellProb: "D55", cellVal: "E55", cellAppr: "F55", cellEV: "H55", probability: 0.06, baseValue: 500000.0, currentPrice: 500000.0, appreciationRate: 0.20, icon: "🌀", color: "#d033ff" },
      { id: "infinityGem", name: "Infinity Gem", cellProb: "D56", cellVal: "E56", cellAppr: "F56", cellEV: "H56", probability: 0.04, baseValue: 1000000.0, currentPrice: 1000000.0, appreciationRate: 0.30, icon: "💎", color: "#ff2d55" }
    ],
    
    titanGear: {
      shovel: { cellCost: "K10", costCoins: 2 },
      loosener: { cellCost: "K11", costCoins: 3 },
      detector: { cellCost: "K12", costCoins: 10 },
      time: { cellCost: "K13", costCoins: 1 }
    },

    voidGear: {
      pickaxe: { cellCost: "K32", costPearls: 2 },
      dynamite: { cellCost: "K33", costPearls: 25 },
      time: { cellCost: "K34", costPearls: 2 }
    },

    novaGear: {
      sledgehammer: { cellCost: "K50", costCrystals: 5 },
      nuke: { cellCost: "K51", costCrystals: 30 },
      stabilizer: { cellCost: "K53", costCrystals: 10 },
      water: { cellCost: "K54", costCrystals: 5 }
    },
    
    ores: [], // Active ores array pointer (assigned to titanOres initially)
    gear: {}, // Active gear pointer (assigned to titanGear initially)
    grossExpectedValue: 98.0,
    voidExpectedValue: 532.0,
    novaExpectedValue: 88000.0
  };
  
  // Set initial pointers
  excelModel.ores = excelModel.titanOres;
  excelModel.gear = excelModel.titanGear;

  // Keep a backup of the original model for resets
  const backupModel = JSON.parse(JSON.stringify(excelModel));

  // Active Game State
  const gameState = {
    isActive: false,
    isPaused: false,
    timer: 300, // seconds remaining
    walletCash: 0.0,
    walletCoins: 0,
    walletPearls: 0,
    walletCrystals: 0,
    digsCompleted: 0,
    coinsFound: 0,
    pearlsFound: 0,
    crystalsFound: 0,
    
    currentAsteroid: "TITAN", // "TITAN", "VOID", or "NOVA"
    
    inventory: {
      iron: 0,
      copper: 0,
      gold: 0,
      amethyst: 0,
      ruby: 0,
      
      bauxite: 0,
      magnetite: 0,
      pyrite: 0,
      diamond: 0,
      emerald: 0,

      cryptonite: 0,
      elerium: 0,
      obsidian: 0,
      warpStone: 0,
      infinityGem: 0
    },

    // Gear ownership (Titan)
    hasTitaniumShovel: false,
    hasDirtLoosener: false,
    hasMetalDetector: false,
    hasExcavator: false,
    fuelLitres: 0.0,
    
    // Gear ownership (Void)
    hasShinyPickaxe: false,
    hasDrillMachine: false,
    gasUnits: 0.0,

    // Gear ownership (Nova)
    hasSledgehammer: false,
    hasCosmicWater: false,
    hasResonanceStabilizer: false,
    hasDrillBore: false,
    tyliumUnits: 0.0,
    
    // Storm & Dynamite & Warhead State
    hasStormProtection: false,
    stormShieldActive: false,
    dynamitePurchasedCount: 0,
    nukePurchasedCount: 0,
    elapsedSeconds: 0,
    scheduledStormSecond: null,
    stormWarningStarted: false,
    stormActive: false,

    // Earthquake state (Nova)
    elapsedEarthquakeSeconds: 0,
    scheduledEarthquakeSecond: null,
    earthquakeWarningStarted: false,
    earthquakeActive: false,

    // Automation state
    excavatorRunning: false,
    excavatorProgress: 0, // ms elapsed since last dig
    
    // Live ore prices appreciation state
    appreciationTimer: 10, // countdown to next appreciation tick
    appreciationTicks: 0,  // number of 10s intervals elapsed
    
    // Live stats tallies for drop rate convergence
    tallies: {
      iron: 0,
      copper: 0,
      gold: 0,
      amethyst: 0,
      ruby: 0,
      coins: 0,
      
      bauxite: 0,
      magnetite: 0,
      pyrite: 0,
      diamond: 0,
      emerald: 0,
      pearls: 0,

      cryptonite: 0,
      elerium: 0,
      obsidian: 0,
      warpStone: 0,
      infinityGem: 0,
      crystals: 0
    },

    // Historical records for canvas chart
    history: {
      timestamps: [],
      netWorths: []
    }
  };

  // Audio configuration
  let soundEnabled = true;
  let audioCtx = null;

  // Active double clicked or selected cell in Excel Sandbox
  let selectedExcelCell = null;

  // Digging state (Manual)
  let isManualDigging = false;
  let manualDigTimeoutId = null;
  let manualDigDuration = 3000; // ms to complete a dig

  // Game tick interval
  let gameIntervalId = null;

  // Cache DOM nodes
  const DOM = {
    // Tabs
    tabs: document.querySelectorAll(".tab-btn"),
    tabContents: document.querySelectorAll(".tab-content"),
    
    // Telemetry & Info
    gameTimer: document.getElementById("game-timer"),
    gameNetWorth: document.getElementById("game-net-worth"),
    telemetryShovel: document.getElementById("telemetry-shovel"),
    telemetryLoosener: document.getElementById("telemetry-loosener"),
    telemetryDetector: document.getElementById("telemetry-detector"),
    telemetryDigs: document.getElementById("telemetry-digs"),
    
    // Wallet
    walletCash: document.getElementById("wallet-cash"),
    walletCoins: document.getElementById("wallet-coins"),
    walletPearls: document.getElementById("wallet-pearls"),
    walletCrystals: document.getElementById("wallet-crystals"),
    coinPillContainer: document.getElementById("coin-pill-container"),
    pearlPillContainer: document.getElementById("pearl-pill-container"),
    crystalPillContainer: document.getElementById("crystal-pill-container"),
    
    // Digging Actions
    btnManualDig: document.getElementById("btn-manual-dig"),
    btnToggleExcavator: document.getElementById("btn-toggle-excavator"),
    fuelDisplaySubtext: document.getElementById("fuel-display-subtext"),
    digProgressBar: document.getElementById("dig-progress-bar"),
    digProgressText: document.getElementById("dig-progress-text"),
    asteroidTarget: document.getElementById("asteroid-target"),
    asteroidWrapper: document.getElementById("asteroid-wrapper"),
    digFlashText: document.getElementById("dig-flash-text"),
    particleEmitter: document.getElementById("particle-emitter"),
    
    // Teleportation
    asteroidName: document.getElementById("asteroid-name"),
    asteroidDistance: document.getElementById("asteroid-distance"),
    btnTeleportVoid: document.getElementById("btn-teleport-void"),
    btnConfirmFlag: document.getElementById("btn-confirm-flag"),
    flagModal: document.getElementById("flag-modal"),
    flagTextInput: document.getElementById("flag-text-input"),
    flagsTableBody: document.getElementById("flags-table-body"),
    teleportDescText: document.getElementById("teleport-desc-text"),
    
    // Nova / Earth / Destination selection / Earthquake overlays
    seismicPanel: document.getElementById("seismic-panel-container"),
    earthquakeOverlay: document.getElementById("earthquake-overlay"),
    destinationModal: document.getElementById("destination-modal"),
    btnCancelDestination: document.getElementById("btn-cancel-destination"),
    btnWarpEarth: document.getElementById("btn-warp-earth"),
    btnWarpNova: document.getElementById("btn-warp-nova"),
    
    // Storm panel & shield controls
    stormPanel: document.getElementById("storm-panel-container"),
    btnLockShield: document.getElementById("btn-lock-shield"),
    btnBuyStormProtection: document.getElementById("btn-buy-storm-protection"),
    
    // Sell Ores
    btnSellAll: document.getElementById("btn-sell-all"),
    appreciationCountdown: document.getElementById("appreciation-countdown"),
    
    // Ore list elements
    countIron: document.getElementById("count-iron"),
    countCopper: document.getElementById("count-copper"),
    countGold: document.getElementById("count-gold"),
    countAmethyst: document.getElementById("count-amethyst"),
    countRuby: document.getElementById("count-ruby"),
    
    priceIron: document.getElementById("price-iron"),
    priceCopper: document.getElementById("price-copper"),
    priceGold: document.getElementById("price-gold"),
    priceAmethyst: document.getElementById("price-amethyst"),
    priceRuby: document.getElementById("price-ruby"),
    
    changeGold: document.getElementById("change-gold"),
    changeAmethyst: document.getElementById("change-amethyst"),
    changeRuby: document.getElementById("change-ruby"),
    
    btnSellIron: document.getElementById("btn-sell-iron"),
    btnSellCopper: document.getElementById("btn-sell-copper"),
    btnSellGold: document.getElementById("btn-sell-gold"),
    btnSellAmethyst: document.getElementById("btn-sell-amethyst"),
    btnSellRuby: document.getElementById("btn-sell-ruby"),
    
    // Shop Actions
    btnBuyExcavator: document.getElementById("btn-buy-excavator"),
    btnBuyFuel: document.getElementById("btn-buy-fuel"),
    btnBuyShovel: document.getElementById("btn-buy-shovel"),
    btnBuyLoosener: document.getElementById("btn-buy-loosener"),
    btnBuyDetector: document.getElementById("btn-buy-detector"),
    btnBuyTime: document.getElementById("btn-buy-time"),
    
    // Convergence Table
    actCountIron: document.getElementById("act-count-iron"),
    actCountCopper: document.getElementById("act-count-copper"),
    actCountGold: document.getElementById("act-count-gold"),
    actCountAmethyst: document.getElementById("act-count-amethyst"),
    actCountRuby: document.getElementById("act-count-ruby"),
    actCountCoins: document.getElementById("act-count-coins"),
    
    actRateIron: document.getElementById("act-rate-iron"),
    actRateCopper: document.getElementById("act-rate-copper"),
    actRateGold: document.getElementById("act-rate-gold"),
    actRateAmethyst: document.getElementById("act-rate-amethyst"),
    actRateRuby: document.getElementById("act-rate-ruby"),
    actRateCoins: document.getElementById("act-rate-coins"),
    
    theoProbIron: document.getElementById("theo-prob-iron"),
    theoProbCopper: document.getElementById("theo-prob-copper"),
    theoProbGold: document.getElementById("theo-prob-gold"),
    theoProbAmethyst: document.getElementById("theo-prob-amethyst"),
    theoProbRuby: document.getElementById("theo-prob-ruby"),
    theoProbCoins: document.getElementById("theo-prob-coins"),

    // Console
    logConsole: document.getElementById("log-console"),
    btnClearConsole: document.getElementById("btn-clear-console"),
    
    // Excel Panel
    fxCellAddress: document.getElementById("fx-cell-address"),
    fxFormulaInput: document.getElementById("fx-formula-input"),
    btnResetExcelModel: document.getElementById("btn-reset-excel-model"),
    probValidationWarning: document.getElementById("prob-validation-warning"),
    probSumDisplay: document.getElementById("prob-sum-display"),
    excelStatusBar: document.getElementById("excel-status-bar"),
    
    // End Game Modal
    endGameModal: document.getElementById("end-game-modal"),
    endNetWorth: document.getElementById("end-net-worth"),
    endCash: document.getElementById("end-cash"),
    endInventoryVal: document.getElementById("end-inventory-val"),
    endDigs: document.getElementById("end-digs"),
    endCoins: document.getElementById("end-coins"),
    highScoreBanner: document.getElementById("high-score-banner"),
    btnRestartGame: document.getElementById("btn-restart-game"),
    
    // Sound Toggle
    soundToggleBtn: document.getElementById("sound-toggle-btn"),
    soundIcon: document.getElementById("sound-icon"),

    // Leaderboard, Pilot Name, Pause Game
    pilotNameInput: document.getElementById("player-name-input"),
    btnPauseGame: document.getElementById("btn-pause-game"),
    pausedOverlay: document.getElementById("paused-overlay"),
    btnResumeOverlay: document.getElementById("btn-resume-overlay"),
    fuelTop: document.getElementById("game-fuel-top"),
    leaderboardTableBody: document.getElementById("leaderboard-table-body"),
    btnClearLeaderboard: document.getElementById("btn-clear-leaderboard"),
    modalLeaderboardBody: document.getElementById("modal-leaderboard-body")
  };

  // Canvas elements
  const canvas = document.getElementById("net-worth-canvas");
  const ctx = canvas ? canvas.getContext("2d") : null;

  // ==========================================================================
  // 2. AUDIO SYNTHESIZER (WEB AUDIO API)
  // ==========================================================================
  function initAudio() {
    if (!audioCtx) {
      audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    }
    if (audioCtx.state === "suspended") {
      audioCtx.resume();
    }
  }

  function playSynthSound(type) {
    if (!soundEnabled) return;
    try {
      initAudio();
      const osc = audioCtx.createOscillator();
      const gainNode = audioCtx.createGain();
      osc.connect(gainNode);
      gainNode.connect(audioCtx.destination);
      const now = audioCtx.currentTime;

      if (type === "click") {
        osc.type = "sine";
        osc.frequency.setValueAtTime(500, now);
        osc.frequency.exponentialRampToValueAtTime(100, now + 0.08);
        gainNode.gain.setValueAtTime(0.12, now);
        gainNode.gain.exponentialRampToValueAtTime(0.001, now + 0.08);
        osc.start(now);
        osc.stop(now + 0.08);
      } 
      else if (type === "dig") {
        // Metallic clink sound
        osc.type = "triangle";
        osc.frequency.setValueAtTime(650, now);
        osc.frequency.exponentialRampToValueAtTime(150, now + 0.15);
        
        // High frequency transient clank
        const osc2 = audioCtx.createOscillator();
        const gain2 = audioCtx.createGain();
        osc2.type = "sine";
        osc2.frequency.setValueAtTime(1200, now);
        osc2.frequency.linearRampToValueAtTime(800, now + 0.03);
        osc2.connect(gain2);
        gain2.connect(audioCtx.destination);
        
        gainNode.gain.setValueAtTime(0.15, now);
        gainNode.gain.exponentialRampToValueAtTime(0.001, now + 0.15);
        gain2.gain.setValueAtTime(0.10, now);
        gain2.gain.exponentialRampToValueAtTime(0.001, now + 0.03);
        
        osc.start(now);
        osc2.start(now);
        osc.stop(now + 0.15);
        osc2.stop(now + 0.03);
      } 
      else if (type === "coin") {
        // High pitched retro double beep
        osc.type = "sine";
        osc.frequency.setValueAtTime(987.77, now); // B5
        osc.frequency.setValueAtTime(1318.51, now + 0.07); // E6
        gainNode.gain.setValueAtTime(0.12, now);
        gainNode.gain.linearRampToValueAtTime(0.12, now + 0.07);
        gainNode.gain.exponentialRampToValueAtTime(0.001, now + 0.22);
        osc.start(now);
        osc.stop(now + 0.22);
      } 
      else if (type === "sell") {
        // Cash register "Cha-Ching"
        osc.type = "triangle";
        osc.frequency.setValueAtTime(1200, now);
        osc.frequency.setValueAtTime(1500, now + 0.05);
        gainNode.gain.setValueAtTime(0.15, now);
        gainNode.gain.setValueAtTime(0.15, now + 0.05);
        gainNode.gain.exponentialRampToValueAtTime(0.001, now + 0.25);
        osc.start(now);
        osc.stop(now + 0.25);
      } 
      else if (type === "upgrade") {
        // Futuristic rising sweep
        osc.type = "sine";
        osc.frequency.setValueAtTime(220, now);
        osc.frequency.exponentialRampToValueAtTime(880, now + 0.3);
        gainNode.gain.setValueAtTime(0.15, now);
        gainNode.gain.exponentialRampToValueAtTime(0.01, now + 0.3);
        osc.start(now);
        osc.stop(now + 0.3);
      } 
      else if (type === "price_up") {
        // Chime for appreciation ticker
        const freq = [523.25, 659.25, 783.99]; // C5, E5, G5
        freq.forEach((f, index) => {
          const oscN = audioCtx.createOscillator();
          const gainN = audioCtx.createGain();
          oscN.connect(gainN);
          gainN.connect(audioCtx.destination);
          oscN.type = "sine";
          oscN.frequency.setValueAtTime(f, now + index * 0.05);
          gainN.gain.setValueAtTime(0.08, now + index * 0.05);
          gainN.gain.exponentialRampToValueAtTime(0.001, now + index * 0.05 + 0.2);
          oscN.start(now + index * 0.05);
          oscN.stop(now + index * 0.05 + 0.2);
        });
      } 
      else if (type === "warning") {
        // Low alert pulse
        osc.type = "sawtooth";
        osc.frequency.setValueAtTime(180, now);
        gainNode.gain.setValueAtTime(0.15, now);
        gainNode.gain.exponentialRampToValueAtTime(0.001, now + 0.15);
        osc.start(now);
        osc.stop(now + 0.15);
      } 
      else if (type === "gameover") {
        // Triumphant game over ditty
        const melody = [523.25, 587.33, 659.25, 783.99, 1046.50]; // C5, D5, E5, G5, C6
        melody.forEach((f, index) => {
          const oscN = audioCtx.createOscillator();
          const gainN = audioCtx.createGain();
          oscN.connect(gainN);
          gainN.connect(audioCtx.destination);
          oscN.type = "triangle";
          oscN.frequency.setValueAtTime(f, now + index * 0.12);
          gainN.gain.setValueAtTime(0.10, now + index * 0.12);
          gainN.gain.exponentialRampToValueAtTime(0.001, now + index * 0.12 + 0.3);
          oscN.start(now + index * 0.12);
          oscN.stop(now + index * 0.12 + 0.35);
        });
      } 
      else if (type === "error") {
        osc.type = "sawtooth";
        osc.frequency.setValueAtTime(120, now);
        osc.frequency.setValueAtTime(100, now + 0.1);
        gainNode.gain.setValueAtTime(0.18, now);
        gainNode.gain.exponentialRampToValueAtTime(0.001, now + 0.25);
        osc.start(now);
        osc.stop(now + 0.25);
      }
    } catch (e) {
      console.warn("Audio Context synthesis blocked/failed:", e);
    }
  }

  function playHeavyExplosionSound() {
    if (!soundEnabled) return;
    try {
      initAudio();
      const now = audioCtx.currentTime;
      const numOscs = 5;
      for (let i = 0; i < numOscs; i++) {
        const osc = audioCtx.createOscillator();
        const gainNode = audioCtx.createGain();
        osc.connect(gainNode);
        gainNode.connect(audioCtx.destination);
        
        osc.type = (i % 2 === 0) ? "sawtooth" : "triangle";
        const baseFreq = 50 + (i * 25);
        osc.frequency.setValueAtTime(baseFreq, now);
        osc.frequency.linearRampToValueAtTime(baseFreq * 0.4, now + 3.5);
        
        gainNode.gain.setValueAtTime(0.35, now);
        gainNode.gain.exponentialRampToValueAtTime(0.001, now + 3.5);
        
        osc.start(now);
        osc.stop(now + 3.5);
      }
    } catch (e) {
      console.warn("Audio Context error", e);
    }
  }

  function playWarpSound() {
    if (!soundEnabled) return;
    try {
      initAudio();
      const now = audioCtx.currentTime;
      const osc = audioCtx.createOscillator();
      const gainNode = audioCtx.createGain();
      osc.connect(gainNode);
      gainNode.connect(audioCtx.destination);
      
      osc.type = "sine";
      osc.frequency.setValueAtTime(100, now);
      osc.frequency.exponentialRampToValueAtTime(1400, now + 2.0);
      osc.frequency.exponentialRampToValueAtTime(250, now + 4.0);
      
      gainNode.gain.setValueAtTime(0.01, now);
      gainNode.gain.linearRampToValueAtTime(0.25, now + 2.0);
      gainNode.gain.exponentialRampToValueAtTime(0.001, now + 4.5);
      
      osc.start(now);
      osc.stop(now + 4.5);
    } catch (e) {
      console.warn("Audio Context error", e);
    }
  }

  // ==========================================================================
  // 3. TAB CONTROLS & SIDEBAR NAVIGATION
  // ==========================================================================
  DOM.tabs.forEach(tab => {
    tab.addEventListener("click", () => {
      DOM.tabs.forEach(t => t.classList.remove("active"));
      DOM.tabContents.forEach(tc => tc.classList.remove("active-content"));
      
      tab.classList.add("active");
      const contentId = tab.getAttribute("data-tab");
      document.getElementById(contentId).classList.add("active-content");
      
      playSynthSound("click");

      // Redraw Canvas Chart if switching to the game view
      if (contentId === "simulator-view") {
        drawNetWorthChart();
      } else if (contentId === "leaderboard-view") {
        updateLeaderboardUIs();
      }
    });
  });

  // Sound toggler
  DOM.soundToggleBtn.addEventListener("click", () => {
    soundEnabled = !soundEnabled;
    if (soundEnabled) {
      DOM.soundIcon.textContent = "🔊";
      DOM.soundToggleBtn.classList.remove("muted");
      DOM.soundToggleBtn.querySelector(".text").textContent = "Sound On";
      playSynthSound("click");
    } else {
      DOM.soundIcon.textContent = "🔇";
      DOM.soundToggleBtn.classList.add("muted");
      DOM.soundToggleBtn.querySelector(".text").textContent = "Muted";
    }
  });

  // Recalculates Expected Value contributions in real time
  function recalculateExcelModel() {
    let titanProbSum = 0;
    excelModel.titanOres.forEach(ore => {
      // H_i = D_i * E_i
      const evContr = ore.probability * ore.baseValue;
      const cellEv = document.getElementById(`cell-${ore.cellEV.toLowerCase()}`);
      if (cellEv) {
        cellEv.textContent = evContr.toFixed(2);
      }
      titanProbSum += ore.probability;
    });
    
    excelModel.grossExpectedValue = excelModel.titanOres.reduce((sum, o) => sum + (o.probability * o.baseValue), 0);
    const cellI20 = document.getElementById("cell-i20");
    if (cellI20) {
      cellI20.textContent = excelModel.grossExpectedValue.toFixed(2);
    }

    let voidProbSum = 0;
    excelModel.voidOres.forEach(ore => {
      // H_i = D_i * E_i
      const evContr = ore.probability * ore.baseValue;
      const cellEv = document.getElementById(`cell-${ore.cellEV.toLowerCase()}`);
      if (cellEv) {
        cellEv.textContent = evContr.toFixed(2);
      }
      voidProbSum += ore.probability;
    });
    
    excelModel.voidExpectedValue = excelModel.voidOres.reduce((sum, o) => sum + (o.probability * o.baseValue), 0);
    const cellI38 = document.getElementById("cell-i38");
    if (cellI38) {
      cellI38.textContent = excelModel.voidExpectedValue.toFixed(2);
    }

    let novaProbSum = 0;
    excelModel.novaOres.forEach(ore => {
      // H_i = D_i * E_i
      const evContr = ore.probability * ore.baseValue;
      const cellEv = document.getElementById(`cell-${ore.cellEV.toLowerCase()}`);
      if (cellEv) {
        cellEv.textContent = evContr.toFixed(2);
      }
      novaProbSum += ore.probability;
    });
    
    excelModel.novaExpectedValue = excelModel.novaOres.reduce((sum, o) => sum + (o.probability * o.baseValue), 0);
    const cellI58 = document.getElementById("cell-i58");
    if (cellI58) {
      cellI58.textContent = excelModel.novaExpectedValue.toFixed(2);
    }
    
    // Display validation warnings based on current asteroid
    let activeProbSum = 0;
    if (gameState.currentAsteroid === "TITAN") {
      activeProbSum = titanProbSum;
    } else if (gameState.currentAsteroid === "VOID") {
      activeProbSum = voidProbSum;
    } else if (gameState.currentAsteroid === "NOVA") {
      activeProbSum = novaProbSum;
    }
    
    const sumPct = Math.round(activeProbSum * 100);
    if (DOM.excelStatusBar) {
      DOM.excelStatusBar.textContent = `Ready. Active Asteroid (${gameState.currentAsteroid}) Probability Total: ${sumPct}%`;
    }

    if (Math.abs(activeProbSum - 1.0) > 0.0001) {
      DOM.probValidationWarning.classList.remove("hidden");
      DOM.probSumDisplay.textContent = `${sumPct}%`;
    } else {
      DOM.probValidationWarning.classList.add("hidden");
    }

    // Live sync game parameters
    syncExcelParametersToGame();
    
    // Refresh formula bar display if there is a selected cell
    if (selectedExcelCell) {
      updateFormulaBar(selectedExcelCell);
    }
  }

  // Push values from Excel Model structure directly to game settings
  function syncExcelParametersToGame() {
    const currentAsteroid = gameState.currentAsteroid;
    
    // Set pointers depending on asteroid
    if (currentAsteroid === "TITAN") {
      excelModel.ores = excelModel.titanOres;
      excelModel.gear = excelModel.titanGear;
    } else if (currentAsteroid === "VOID") {
      excelModel.ores = excelModel.voidOres;
      excelModel.gear = excelModel.voidGear;
    } else if (currentAsteroid === "NOVA") {
      excelModel.ores = excelModel.novaOres;
      excelModel.gear = excelModel.novaGear;
    }

    // Starting cash
    if (!gameState.isActive && currentAsteroid === "TITAN") {
      gameState.walletCash = excelModel.startingMoney;
      gameState.timer = excelModel.initialTimer;
      updateUI();
    }
    
    // Prices
    excelModel.ores.forEach(ore => {
      ore.currentPrice = ore.baseValue * (1.0 + (gameState.appreciationTicks * ore.appreciationRate));
      
      const priceLbl = document.getElementById(`price-${ore.id}`);
      if (priceLbl) {
        priceLbl.textContent = `$${Math.round(ore.currentPrice)}`;
      }
      
      // Update appreciation label percentage
      const changeLbl = document.getElementById(`change-${ore.id}`);
      if (changeLbl) {
        const flatIncrease = ore.baseValue * ore.appreciationRate;
        const ratePct = Math.round(ore.appreciationRate * 100);
        changeLbl.textContent = `+$${Math.round(flatIncrease)} (+${ratePct}%)`;
      }
    });

    // Update theoretical headers in Convergence table
    DOM.theoProbIron.textContent = `${(excelModel.ores[0].probability * 100).toFixed(1)}%`;
    DOM.theoProbCopper.textContent = `${(excelModel.ores[1].probability * 100).toFixed(1)}%`;
    DOM.theoProbGold.textContent = `${(excelModel.ores[2].probability * 100).toFixed(1)}%`;
    DOM.theoProbAmethyst.textContent = `${(excelModel.ores[3].probability * 100).toFixed(1)}%`;
    DOM.theoProbRuby.textContent = `${(excelModel.ores[4].probability * 100).toFixed(1)}%`;
    
    if (currentAsteroid === "TITAN") {
      DOM.theoProbCoins.textContent = `${(excelModel.coinDropChance * 100).toFixed(1)}%`;
    } else if (currentAsteroid === "VOID") {
      DOM.theoProbCoins.textContent = `${(excelModel.pearlDropChance * 100).toFixed(1)}%`;
    } else if (currentAsteroid === "NOVA") {
      DOM.theoProbCoins.textContent = `${(excelModel.crystalDropChance * 100).toFixed(1)}%`;
    }
    
    // Recalculate manual digging speed
    recalculateDiggingSpeeds();
  }

  function updateCellEditabilityStyles() {
    const currentAsteroid = gameState.currentAsteroid;
    
    const titanCellIds = [
      "cell-e6", "cell-e7", "cell-e9", "cell-e10", "cell-e11",
      "cell-d21", "cell-d22", "cell-d23", "cell-d24", "cell-d25", "cell-d26",
      "cell-e21", "cell-e22", "cell-e23", "cell-e24", "cell-e25",
      "cell-f21", "cell-f22", "cell-f23", "cell-f24", "cell-f25",
      "cell-k10", "cell-k11", "cell-k12", "cell-k13"
    ];
    
    const voidCellIds = [
      "cell-d32", "cell-d33", "cell-d34", "cell-d35", "cell-d36", "cell-d37",
      "cell-e32", "cell-e33", "cell-e34", "cell-e35", "cell-e36",
      "cell-f32", "cell-f33", "cell-f34", "cell-f35", "cell-f36",
      "cell-k32", "cell-k33", "cell-k34",
      "cell-d39", "cell-d40", "cell-d41"
    ];

    const novaCellIds = [
      "cell-d52", "cell-d53", "cell-d54", "cell-d55", "cell-d56", "cell-d57",
      "cell-e52", "cell-e53", "cell-e54", "cell-e55", "cell-e56",
      "cell-f52", "cell-f53", "cell-f54", "cell-f55", "cell-f56",
      "cell-k50", "cell-k51", "cell-k53", "cell-k54",
      "cell-d59", "cell-d60", "cell-d61", "cell-d63"
    ];

    titanCellIds.forEach(id => {
      const cell = document.getElementById(id);
      if (cell) {
        if (currentAsteroid === "TITAN") {
          cell.classList.remove("cell-disabled");
          cell.setAttribute("tabindex", "0");
        } else {
          cell.classList.add("cell-disabled");
          cell.removeAttribute("tabindex");
        }
      }
    });

    voidCellIds.forEach(id => {
      const cell = document.getElementById(id);
      if (cell) {
        if (currentAsteroid === "VOID") {
          cell.classList.remove("cell-disabled");
          cell.setAttribute("tabindex", "0");
        } else {
          cell.classList.add("cell-disabled");
          cell.removeAttribute("tabindex");
        }
      }
    });

    novaCellIds.forEach(id => {
      const cell = document.getElementById(id);
      if (cell) {
        if (currentAsteroid === "NOVA") {
          cell.classList.remove("cell-disabled");
          cell.setAttribute("tabindex", "0");
        } else {
          cell.classList.add("cell-disabled");
          cell.removeAttribute("tabindex");
        }
      }
    });
  }

  // Synchronize Excel model parameters into DOM cells
  function syncModelToDOM() {
    // Titan Config
    document.getElementById("cell-e6").textContent = excelModel.initialTimer;
    document.getElementById("cell-e7").textContent = excelModel.startingMoney.toFixed(0);
    document.getElementById("cell-e9").textContent = excelModel.excavatorCost.toFixed(0);
    document.getElementById("cell-e10").textContent = excelModel.fuelCost.toFixed(0);
    document.getElementById("cell-e11").textContent = excelModel.travelCost.toFixed(0);
    
    // Titan Ores
    excelModel.titanOres.forEach(ore => {
      document.getElementById(`cell-${ore.cellProb.toLowerCase()}`).textContent = `${(ore.probability * 100).toFixed(1)}%`;
      document.getElementById(`cell-${ore.cellVal.toLowerCase()}`).textContent = ore.baseValue.toFixed(0);
      document.getElementById(`cell-${ore.cellAppr.toLowerCase()}`).textContent = `${(ore.appreciationRate * 100).toFixed(1)}%`;
    });
    
    // Titan Gear Costs
    document.getElementById("cell-k10").textContent = excelModel.titanGear.shovel.costCoins;
    document.getElementById("cell-k11").textContent = excelModel.titanGear.loosener.costCoins;
    document.getElementById("cell-k12").textContent = excelModel.titanGear.detector.costCoins;
    document.getElementById("cell-k13").textContent = excelModel.titanGear.time.costCoins;
    
    // Void Ores
    excelModel.voidOres.forEach(ore => {
      document.getElementById(`cell-${ore.cellProb.toLowerCase()}`).textContent = `${(ore.probability * 100).toFixed(1)}%`;
      document.getElementById(`cell-${ore.cellVal.toLowerCase()}`).textContent = ore.baseValue.toFixed(0);
      document.getElementById(`cell-${ore.cellAppr.toLowerCase()}`).textContent = `${(ore.appreciationRate * 100).toFixed(1)}%`;
    });
    
    // Void Gear & Machinery
    document.getElementById("cell-d26").textContent = `${Math.round(excelModel.coinDropChance * 100)}%`;
    document.getElementById("cell-d37").textContent = `${Math.round(excelModel.pearlDropChance * 100)}%`;
    document.getElementById("cell-d39").textContent = excelModel.drillCost.toFixed(0);
    document.getElementById("cell-d40").textContent = excelModel.gasCost.toFixed(0);
    document.getElementById("cell-d41").textContent = excelModel.stormProtectionCost.toFixed(0);
    
    document.getElementById("cell-k32").textContent = excelModel.voidGear.pickaxe.costPearls;
    document.getElementById("cell-k33").textContent = excelModel.voidGear.dynamite.costPearls;
    document.getElementById("cell-k34").textContent = excelModel.voidGear.time.costPearls;

    // Nova Ores
    excelModel.novaOres.forEach(ore => {
      document.getElementById(`cell-${ore.cellProb.toLowerCase()}`).textContent = `${(ore.probability * 100).toFixed(1)}%`;
      document.getElementById(`cell-${ore.cellVal.toLowerCase()}`).textContent = ore.baseValue.toFixed(0);
      document.getElementById(`cell-${ore.cellAppr.toLowerCase()}`).textContent = `${(ore.appreciationRate * 100).toFixed(1)}%`;
    });
    
    // Nova Gear & Machinery
    document.getElementById("cell-d57").textContent = `${Math.round(excelModel.crystalDropChance * 100)}%`;
    document.getElementById("cell-d59").textContent = excelModel.drillBoreCost.toFixed(0);
    document.getElementById("cell-d60").textContent = excelModel.tyliumCost.toFixed(0);
    document.getElementById("cell-d61").textContent = excelModel.novaStormProtectionCost.toFixed(0);
    document.getElementById("cell-d63").textContent = excelModel.earthReturnCost.toFixed(0);
    
    document.getElementById("cell-k50").textContent = excelModel.novaGear.sledgehammer.costCrystals;
    document.getElementById("cell-k51").textContent = excelModel.novaGear.nuke.costCrystals;
    document.getElementById("cell-k53").textContent = excelModel.novaGear.stabilizer.costCrystals;
    document.getElementById("cell-k54").textContent = excelModel.novaGear.water.costCrystals;

    updateCellEditabilityStyles();
    recalculateExcelModel();
  }

  // Update cell data on blur/save
  function updateCellData(cellId, rawText) {
    let parsedVal = parseFloat(rawText.replace(/[^0-9.-]/g, ""));
    if (isNaN(parsedVal)) parsedVal = 0;
    
    // Titan Core Settings
    if (cellId === "cell-e7") {
      excelModel.startingMoney = Math.max(0, parsedVal);
    } 
    else if (cellId === "cell-e6") {
      excelModel.initialTimer = Math.max(10, Math.round(parsedVal));
    } 
    else if (cellId === "cell-e9") {
      excelModel.excavatorCost = Math.max(1, parsedVal);
    } 
    else if (cellId === "cell-e10") {
      excelModel.fuelCost = Math.max(1, parsedVal);
    } 
    else if (cellId === "cell-e11") {
      excelModel.travelCost = Math.max(1, parsedVal);
    } 
    // Titan Ores
    else if (cellId === "cell-d21") { updateTitanOreProb(0, parsedVal, rawText); }
    else if (cellId === "cell-d22") { updateTitanOreProb(1, parsedVal, rawText); }
    else if (cellId === "cell-d23") { updateTitanOreProb(2, parsedVal, rawText); }
    else if (cellId === "cell-d24") { updateTitanOreProb(3, parsedVal, rawText); }
    else if (cellId === "cell-d25") { updateTitanOreProb(4, parsedVal, rawText); }
    
    else if (cellId === "cell-e21") { excelModel.titanOres[0].baseValue = Math.max(0, parsedVal); }
    else if (cellId === "cell-e22") { excelModel.titanOres[1].baseValue = Math.max(0, parsedVal); }
    else if (cellId === "cell-e23") { excelModel.titanOres[2].baseValue = Math.max(0, parsedVal); }
    else if (cellId === "cell-e24") { excelModel.titanOres[3].baseValue = Math.max(0, parsedVal); }
    else if (cellId === "cell-e25") { excelModel.titanOres[4].baseValue = Math.max(0, parsedVal); }
    
    else if (cellId === "cell-f21") { updateTitanOreAppr(0, parsedVal, rawText); }
    else if (cellId === "cell-f22") { updateTitanOreAppr(1, parsedVal, rawText); }
    else if (cellId === "cell-f23") { updateTitanOreAppr(2, parsedVal, rawText); }
    else if (cellId === "cell-f24") { updateTitanOreAppr(3, parsedVal, rawText); }
    else if (cellId === "cell-f25") { updateTitanOreAppr(4, parsedVal, rawText); }
    
    else if (cellId === "cell-d26") {
      let isPct = rawText.includes("%") || (rawText.includes("0.") === false && parsedVal > 1);
      excelModel.coinDropChance = Math.max(0, Math.min(1, isPct ? parsedVal / 100 : parsedVal));
    }
    
    // Titan Gear Costs
    else if (cellId === "cell-k10") { excelModel.titanGear.shovel.costCoins = Math.max(1, Math.round(parsedVal)); }
    else if (cellId === "cell-k11") { excelModel.titanGear.loosener.costCoins = Math.max(1, Math.round(parsedVal)); }
    else if (cellId === "cell-k12") { excelModel.titanGear.detector.costCoins = Math.max(1, Math.round(parsedVal)); }
    else if (cellId === "cell-k13") { excelModel.titanGear.time.costCoins = Math.max(1, Math.round(parsedVal)); }
    
    // Void Ores
    else if (cellId === "cell-d32") { updateVoidOreProb(0, parsedVal, rawText); }
    else if (cellId === "cell-d33") { updateVoidOreProb(1, parsedVal, rawText); }
    else if (cellId === "cell-d34") { updateVoidOreProb(2, parsedVal, rawText); }
    else if (cellId === "cell-d35") { updateVoidOreProb(3, parsedVal, rawText); }
    else if (cellId === "cell-d36") { updateVoidOreProb(4, parsedVal, rawText); }
    
    else if (cellId === "cell-e32") { excelModel.voidOres[0].baseValue = Math.max(0, parsedVal); }
    else if (cellId === "cell-e33") { excelModel.voidOres[1].baseValue = Math.max(0, parsedVal); }
    else if (cellId === "cell-e34") { excelModel.voidOres[2].baseValue = Math.max(0, parsedVal); }
    else if (cellId === "cell-e35") { excelModel.voidOres[3].baseValue = Math.max(0, parsedVal); }
    else if (cellId === "cell-e36") { excelModel.voidOres[4].baseValue = Math.max(0, parsedVal); }
    
    else if (cellId === "cell-f32") { updateVoidOreAppr(0, parsedVal, rawText); }
    else if (cellId === "cell-f33") { updateVoidOreAppr(1, parsedVal, rawText); }
    else if (cellId === "cell-f34") { updateVoidOreAppr(2, parsedVal, rawText); }
    else if (cellId === "cell-f35") { updateVoidOreAppr(3, parsedVal, rawText); }
    else if (cellId === "cell-f36") { updateVoidOreAppr(4, parsedVal, rawText); }
    
    else if (cellId === "cell-d37") {
      let isPct = rawText.includes("%") || (rawText.includes("0.") === false && parsedVal > 1);
      excelModel.pearlDropChance = Math.max(0, Math.min(1, isPct ? parsedVal / 100 : parsedVal));
    }
    
    // Void Machine Costs
    else if (cellId === "cell-d39") { excelModel.drillCost = Math.max(1, parsedVal); }
    else if (cellId === "cell-d40") { excelModel.gasCost = Math.max(1, parsedVal); }
    else if (cellId === "cell-d41") { excelModel.stormProtectionCost = Math.max(1, parsedVal); }
    
    // Void Gear Costs
    else if (cellId === "cell-k32") { excelModel.voidGear.pickaxe.costPearls = Math.max(1, Math.round(parsedVal)); }
    else if (cellId === "cell-k33") { excelModel.voidGear.dynamite.costPearls = Math.max(1, Math.round(parsedVal)); }
    else if (cellId === "cell-k34") { excelModel.voidGear.time.costPearls = Math.max(1, Math.round(parsedVal)); }

    // Nova Ores
    else if (cellId === "cell-d52") { updateNovaOreProb(0, parsedVal, rawText); }
    else if (cellId === "cell-d53") { updateNovaOreProb(1, parsedVal, rawText); }
    else if (cellId === "cell-d54") { updateNovaOreProb(2, parsedVal, rawText); }
    else if (cellId === "cell-d55") { updateNovaOreProb(3, parsedVal, rawText); }
    else if (cellId === "cell-d56") { updateNovaOreProb(4, parsedVal, rawText); }
    
    else if (cellId === "cell-e52") { excelModel.novaOres[0].baseValue = Math.max(0, parsedVal); }
    else if (cellId === "cell-e53") { excelModel.novaOres[1].baseValue = Math.max(0, parsedVal); }
    else if (cellId === "cell-e54") { excelModel.novaOres[2].baseValue = Math.max(0, parsedVal); }
    else if (cellId === "cell-e55") { excelModel.novaOres[3].baseValue = Math.max(0, parsedVal); }
    else if (cellId === "cell-e56") { excelModel.novaOres[4].baseValue = Math.max(0, parsedVal); }
    
    else if (cellId === "cell-f52") { updateNovaOreAppr(0, parsedVal, rawText); }
    else if (cellId === "cell-f53") { updateNovaOreAppr(1, parsedVal, rawText); }
    else if (cellId === "cell-f54") { updateNovaOreAppr(2, parsedVal, rawText); }
    else if (cellId === "cell-f55") { updateNovaOreAppr(3, parsedVal, rawText); }
    else if (cellId === "cell-f56") { updateNovaOreAppr(4, parsedVal, rawText); }
    
    else if (cellId === "cell-d57") {
      let isPct = rawText.includes("%") || (rawText.includes("0.") === false && parsedVal > 1);
      excelModel.crystalDropChance = Math.max(0, Math.min(1, isPct ? parsedVal / 100 : parsedVal));
    }
    
    // Nova Machine Costs
    else if (cellId === "cell-d59") { excelModel.drillBoreCost = Math.max(1, parsedVal); }
    else if (cellId === "cell-d60") { excelModel.tyliumCost = Math.max(1, parsedVal); }
    else if (cellId === "cell-d61") { excelModel.novaStormProtectionCost = Math.max(1, parsedVal); }
    else if (cellId === "cell-d63") { excelModel.earthReturnCost = Math.max(1, parsedVal); }
    
    // Nova Gear Costs
    else if (cellId === "cell-k50") { excelModel.novaGear.sledgehammer.costCrystals = Math.max(1, Math.round(parsedVal)); }
    else if (cellId === "cell-k51") { excelModel.novaGear.nuke.costCrystals = Math.max(1, Math.round(parsedVal)); }
    else if (cellId === "cell-k53") { excelModel.novaGear.stabilizer.costCrystals = Math.max(1, Math.round(parsedVal)); }
    else if (cellId === "cell-k54") { excelModel.novaGear.water.costCrystals = Math.max(1, Math.round(parsedVal)); }

    syncModelToDOM();
  }

  function updateTitanOreProb(idx, val, raw) {
    let isPct = raw.includes("%") || (raw.includes("0.") === false && val > 1);
    let finalProb = Math.max(0, Math.min(1, isPct ? val / 100 : val));
    excelModel.titanOres[idx].probability = finalProb;
  }

  function updateTitanOreAppr(idx, val, raw) {
    let isPct = raw.includes("%") || (raw.includes("0.") === false && val > 1);
    let finalAppr = Math.max(0, Math.min(5, isPct ? val / 100 : val));
    excelModel.titanOres[idx].appreciationRate = finalAppr;
  }

  function updateVoidOreProb(idx, val, raw) {
    let isPct = raw.includes("%") || (raw.includes("0.") === false && val > 1);
    let finalProb = Math.max(0, Math.min(1, isPct ? val / 100 : val));
    excelModel.voidOres[idx].probability = finalProb;
  }

  function updateVoidOreAppr(idx, val, raw) {
    let isPct = raw.includes("%") || (raw.includes("0.") === false && val > 1);
    let finalAppr = Math.max(0, Math.min(5, isPct ? val / 100 : val));
    excelModel.voidOres[idx].appreciationRate = finalAppr;
  }

  function updateNovaOreProb(idx, val, raw) {
    let isPct = raw.includes("%") || (raw.includes("0.") === false && val > 1);
    let finalProb = Math.max(0, Math.min(1, isPct ? val / 100 : val));
    excelModel.novaOres[idx].probability = finalProb;
  }

  function updateNovaOreAppr(idx, val, raw) {
    let isPct = raw.includes("%") || (raw.includes("0.") === false && val > 1);
    let finalAppr = Math.max(0, Math.min(5, isPct ? val / 100 : val));
    excelModel.novaOres[idx].appreciationRate = finalAppr;
  }

  // Active Excel selection
  function selectCell(element) {
    document.querySelectorAll(".excel-table td").forEach(td => td.classList.remove("selected-cell"));
    selectedExcelCell = element;
    element.classList.add("selected-cell");
    updateFormulaBar(element);
    playSynthSound("click");
  }

  function updateFormulaBar(element) {
    const address = element.getAttribute("data-cell");
    DOM.fxCellAddress.textContent = address;
    
    const hasFormula = element.hasAttribute("data-formula");
    if (hasFormula) {
      DOM.fxFormulaInput.value = element.getAttribute("data-formula");
    } else {
      DOM.fxFormulaInput.value = element.textContent.trim();
    }
  }

  // Wire up Excel Grid Interactions
  document.querySelectorAll(".excel-table td[data-cell]").forEach(cell => {
    cell.addEventListener("click", () => selectCell(cell));
  });

  function isCellEditableOnCurrentAsteroid(cell) {
    const cellId = cell.id;
    const isTitan = (gameState.currentAsteroid === "TITAN");
    
    const titanCellIds = [
      "cell-e6", "cell-e7", "cell-e9", "cell-e10", "cell-e11",
      "cell-d21", "cell-d22", "cell-d23", "cell-d24", "cell-d25", "cell-d26",
      "cell-e21", "cell-e22", "cell-e23", "cell-e24", "cell-e25",
      "cell-f21", "cell-f22", "cell-f23", "cell-f24", "cell-f25",
      "cell-k10", "cell-k11", "cell-k12", "cell-k13"
    ];

    const isTitanCell = titanCellIds.includes(cellId);
    return isTitan ? isTitanCell : !isTitanCell;
  }

  const editableCells = document.querySelectorAll(".excel-table td.editable-cell");
  editableCells.forEach(cell => {
    cell.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        cell.blur();
      }
      if (e.key === "Escape") {
        syncModelToDOM();
        cell.blur();
      }
    });
    
    cell.addEventListener("focus", () => {
      if (!isCellEditableOnCurrentAsteroid(cell)) {
        cell.blur();
        return;
      }
      selectCell(cell);
      cell.contentEditable = "true";
    });
    
    cell.addEventListener("blur", () => {
      cell.contentEditable = "false";
      updateCellData(cell.id, cell.textContent.trim());
    });
  });

  // Reset Excel defaults
  DOM.btnResetExcelModel.addEventListener("click", () => {
    // Deep clone from original defaults backup
    Object.assign(excelModel, JSON.parse(JSON.stringify(backupModel)));
    syncModelToDOM();
    writeConsoleLog("Spreadsheet config reset to standard bootcamp default values.");
    playSynthSound("upgrade");
  });

  // Recalculate manual digging speed based on gear owned
  function recalculateDiggingSpeeds() {
    let duration = 3000;
    const currentAsteroid = gameState.currentAsteroid;
    
    if (currentAsteroid === "TITAN") {
      if (gameState.hasTitaniumShovel) {
        duration = 1500;
      }
      if (gameState.hasDirtLoosener) {
        duration = duration / 1.5; // +50% speed = divide time by 1.5
      }
    } else if (currentAsteroid === "VOID") {
      duration = 1500; // Shiny Shovel speed by default in VOID
      if (gameState.hasShinyPickaxe) {
        duration = 750; // Shiny Pickaxe reduces it to 750ms
      }
    } else if (currentAsteroid === "NOVA") {
      duration = 1000; // Ice Axe base speed is 1000ms in NOVA
      if (gameState.hasSledgehammer) {
        duration = 500; // Sledgehammer reduces it to 500ms
      }
      if (gameState.hasCosmicWater) {
        duration = duration / 1.5; // speeds up by 50%
      }
      if (gameState.hasSledgehammer && gameState.hasCosmicWater) {
        duration = 250; // combines with Sledgehammer to reach 250ms
      }
    }
    manualDigDuration = duration;
  }

  // ==========================================================================
  // 5. GAME SYSTEM LOGIC (ASTEROID mining RUSH)
  // ==========================================================================
  
  // Game Setup & Initialization
  function startNewGame() {
    // Reset state values
    gameState.isActive = true;
    gameState.isPaused = false;
    
    // Reset asteroid back to TITAN
    gameState.currentAsteroid = "TITAN";
    document.body.classList.remove("void-theme");
    document.body.classList.remove("nova-theme");
    DOM.asteroidName.textContent = "TITAN";
    DOM.asteroidName.className = "asteroid-titan-text";
    DOM.asteroidDistance.textContent = "(Near Earth)";
    
    // Reset pointers to Titan
    excelModel.ores = excelModel.titanOres;
    excelModel.gear = excelModel.titanGear;
    
    gameState.timer = excelModel.initialTimer;
    gameState.walletCash = excelModel.startingMoney;
    gameState.walletCoins = 0;
    gameState.walletPearls = 0;
    gameState.walletCrystals = 0;
    gameState.digsCompleted = 0;
    gameState.coinsFound = 0;
    gameState.pearlsFound = 0;
    gameState.crystalsFound = 0;
    
    Object.keys(gameState.inventory).forEach(k => gameState.inventory[k] = 0);
    Object.keys(gameState.tallies).forEach(k => gameState.tallies[k] = 0);
    
    gameState.hasTitaniumShovel = false;
    gameState.hasDirtLoosener = false;
    gameState.hasMetalDetector = false;
    gameState.hasExcavator = false;
    gameState.fuelLitres = 0.0;
    gameState.excavatorRunning = false;
    gameState.excavatorProgress = 0;
    
    gameState.hasShinyPickaxe = false;
    gameState.hasDrillMachine = false;
    gameState.gasUnits = 0.0;
    
    gameState.hasSledgehammer = false;
    gameState.hasCosmicWater = false;
    gameState.hasResonanceStabilizer = false;
    gameState.hasDrillBore = false;
    gameState.tyliumUnits = 0.0;
    gameState.nukePurchasedCount = 0;
    excelModel.novaGear.nuke.costCrystals = 30;
    
    // Reset Storm & Dynamite variables
    gameState.hasStormProtection = false;
    gameState.stormShieldActive = false;
    gameState.dynamitePurchasedCount = 0;
    gameState.elapsedSeconds = 0;
    gameState.scheduledStormSecond = null;
    gameState.stormWarningStarted = false;
    gameState.stormActive = false;
    excelModel.voidGear.dynamite.costPearls = 25;

    // Reset Earthquake variables
    gameState.elapsedEarthquakeSeconds = 0;
    gameState.scheduledEarthquakeSecond = null;
    gameState.earthquakeWarningStarted = false;
    gameState.earthquakeActive = false;

    // Reset UI overlays
    if (DOM.stormPanel) {
      DOM.stormPanel.classList.add("hidden");
      DOM.btnLockShield.disabled = true;
      DOM.btnLockShield.textContent = "🔒 ACTIVATE STORM SHIELD";
      DOM.btnLockShield.style.background = "";
    }
    if (DOM.seismicPanel) {
      DOM.seismicPanel.classList.add("hidden");
    }
    const stOverlay = document.getElementById("storm-overlay");
    if (stOverlay) stOverlay.classList.add("hidden");
    const expOverlay = document.getElementById("explosion-overlay");
    if (expOverlay) expOverlay.classList.add("hidden");
    const warpOverlay = document.getElementById("warp-overlay");
    if (warpOverlay) warpOverlay.classList.add("hidden");
    const destModal = document.getElementById("destination-modal");
    if (destModal) destModal.classList.add("hidden");
    const eqOverlay = document.getElementById("earthquake-overlay");
    if (eqOverlay) eqOverlay.classList.add("hidden");
    
    const crystalPill = document.getElementById("crystal-pill-container");
    if (crystalPill) crystalPill.classList.add("hidden");

    gameState.appreciationTimer = 10;
    gameState.appreciationTicks = 0;
    
    // Clear and initialize history
    gameState.history.timestamps = [0];
    gameState.history.netWorths = [excelModel.startingMoney];
    
    // Reset ore current prices to base
    excelModel.titanOres.forEach(ore => ore.currentPrice = ore.baseValue);
    excelModel.voidOres.forEach(ore => ore.currentPrice = ore.baseValue);
    excelModel.novaOres.forEach(ore => ore.currentPrice = ore.baseValue);
    
    // Visual progress resets
    DOM.digProgressBar.style.width = "0%";
    DOM.digProgressText.textContent = "Ready to Dig";
    isManualDigging = false;
    if (manualDigTimeoutId) clearTimeout(manualDigTimeoutId);
    
    // Reset modal overlays
    DOM.endGameModal.classList.add("hidden");
    DOM.flagModal.classList.add("hidden");
    if (DOM.pausedOverlay) DOM.pausedOverlay.classList.add("hidden");
    if (DOM.btnPauseGame) {
      DOM.btnPauseGame.textContent = "⏸️ Pause";
      DOM.btnPauseGame.classList.remove("paused");
    }
    
    // Telemetry updates
    DOM.telemetryShovel.textContent = "Rusty Shovel";
    DOM.telemetryLoosener.textContent = "None (1x Speed)";
    DOM.telemetryDetector.textContent = "Not Equipped";
    DOM.telemetryDigs.textContent = "0";
    
    // Sync buttons
    syncExcelParametersToGame();
    syncModelToDOM();
    updateUI();
    
    // Start main game loop (runs every 100ms for responsiveness)
    if (gameIntervalId) clearInterval(gameIntervalId);
    gameIntervalId = setInterval(gameLoopTick, 100);
    
    writeConsoleLog("ASTERIOD MINING EXPEDITION STARTED! Miner, deploy shovel.");
    playSynthSound("upgrade");
  }

  // Choose ore drop based on probabilities (normal or boosted by metal detector)
  function rollMiningReward() {
    let probSum = 0;
    const items = [];
    
    if (gameState.currentAsteroid === "TITAN" && gameState.hasMetalDetector) {
      // Tripled probabilities for Gold, Amethyst, Ruby
      const tripleGold = Math.min(1.0, excelModel.ores[2].probability * 3);
      const tripleAmethyst = Math.min(1.0 - tripleGold, excelModel.ores[3].probability * 3);
      const tripleRuby = Math.min(1.0 - tripleGold - tripleAmethyst, excelModel.ores[4].probability * 3);
      
      const sumRares = tripleGold + tripleAmethyst + tripleRuby;
      const remaining = Math.max(0, 1.0 - sumRares);
      
      // Split remaining equally between Iron and Copper
      const halfRemaining = remaining / 2;
      
      items.push({ id: "iron", prob: halfRemaining });
      items.push({ id: "copper", prob: halfRemaining });
      items.push({ id: "gold", prob: tripleGold });
      items.push({ id: "amethyst", prob: tripleAmethyst });
      items.push({ id: "ruby", prob: tripleRuby });
    } else if (gameState.currentAsteroid === "NOVA" && gameState.hasResonanceStabilizer) {
      // Doubled probabilities for Obsidian, Warp-Stone, Infinity Gem
      const doubleObsidian = Math.min(1.0, excelModel.ores[2].probability * 2);
      const doubleWarp = Math.min(1.0 - doubleObsidian, excelModel.ores[3].probability * 2);
      const doubleInfinity = Math.min(1.0 - doubleObsidian - doubleWarp, excelModel.ores[4].probability * 2);
      
      const sumRares = doubleObsidian + doubleWarp + doubleInfinity;
      const remaining = Math.max(0, 1.0 - sumRares);
      const halfRemaining = remaining / 2;
      
      items.push({ id: "cryptonite", prob: halfRemaining });
      items.push({ id: "elerium", prob: halfRemaining });
      items.push({ id: "obsidian", prob: doubleObsidian });
      items.push({ id: "warpStone", prob: doubleWarp });
      items.push({ id: "infinityGem", prob: doubleInfinity });
    } else {
      // Use standard probabilities
      excelModel.ores.forEach(ore => {
        items.push({ id: ore.id, prob: ore.probability });
      });
    }
    
    // Roll
    const totalProb = items.reduce((sum, item) => sum + item.prob, 0);
    const roll = Math.random() * totalProb;
    
    let cumulative = 0;
    for (let i = 0; i < items.length; i++) {
      cumulative += items[i].prob;
      if (roll <= cumulative) {
        return excelModel.ores.find(o => o.id === items[i].id);
      }
    }
    return excelModel.ores[0]; // fallback
  }

  // Register a successful dig
  function executeSuccessfulDig(sourceName) {
    if (!gameState.isActive) return;
    
    gameState.digsCompleted++;
    DOM.telemetryDigs.textContent = gameState.digsCompleted;
    
    // 1. Roll Ore
    const drop = rollMiningReward();
    gameState.inventory[drop.id]++;
    gameState.tallies[drop.id]++;
    
    // 2. Roll Coins/Pearls/Crystals (independent roll)
    let foundSecondary = false;
    let dropChance = 0;
    const isTitan = (gameState.currentAsteroid === "TITAN");
    const isVoid = (gameState.currentAsteroid === "VOID");
    if (isTitan) {
      dropChance = excelModel.coinDropChance;
    } else if (isVoid) {
      dropChance = excelModel.pearlDropChance;
    } else {
      dropChance = excelModel.crystalDropChance;
    }
    
    if (Math.random() < dropChance) {
      foundSecondary = true;
      if (isTitan) {
        gameState.walletCoins++;
        gameState.tallies.coins++;
      } else if (isVoid) {
        gameState.walletPearls++;
        gameState.tallies.pearls++;
      } else {
        gameState.walletCrystals++;
        gameState.tallies.crystals++;
      }
    }
    
    // Visual Effects
    triggerAsteroidDigFeedback(drop, foundSecondary);
    
    // Logs
    const secName = isTitan ? "Coin" : (isVoid ? "Pearl" : "Crystal");
    const secIcon = isTitan ? "🪙" : (isVoid ? "🦪" : "🔮");
    const secMsg = foundSecondary ? ` (+1 ${secIcon} ${secName} found!)` : "";
    writeConsoleLog(`[${sourceName}] Mined ${drop.icon} ${drop.name} worth $${Math.round(drop.currentPrice)}${secMsg}.`);
    
    // Sound chimes
    if (foundSecondary) {
      playSynthSound("coin");
    } else {
      playSynthSound("dig");
    }
    
    // Update Stats & Wallet
    updateUI();
  }

  // Digital visual pops on the Asteroid
  function triggerAsteroidDigFeedback(oreDrop, foundSecondary) {
    // 1. Shake asteroid
    DOM.asteroidWrapper.classList.add("dig-shake");
    setTimeout(() => DOM.asteroidWrapper.classList.remove("dig-shake"), 150);
    
    // 2. Floating text pop
    const isTitan = (gameState.currentAsteroid === "TITAN");
    const isVoid = (gameState.currentAsteroid === "VOID");
    const secIcon = isTitan ? "🪙" : (isVoid ? "🦪" : "🔮");
    const secLabel = isTitan ? "Coin" : (isVoid ? "Pearl" : "Crystal");
    const textPop = foundSecondary ? `+${oreDrop.icon} Mined & +${secIcon} ${secLabel}!` : `+${oreDrop.icon} ${oreDrop.name}`;
    DOM.digFlashText.textContent = textPop;
    DOM.digFlashText.style.color = oreDrop.color;
    
    // Trigger CSS re-flow to restart animation
    DOM.digFlashText.style.animation = 'none';
    DOM.digFlashText.offsetHeight;
    DOM.digFlashText.style.animation = null;
    
    // 3. Spawn stone particles
    createStoneParticles(oreDrop.color);
  }

  // Spawn stone shards when mining
  function createStoneParticles(color) {
    const pCount = 12;
    for (let i = 0; i < pCount; i++) {
      const p = document.createElement("div");
      p.className = "particle";
      p.style.backgroundColor = color;
      p.style.boxShadow = `0 0 6px ${color}`;
      
      const size = Math.random() * 8 + 4;
      p.style.width = `${size}px`;
      p.style.height = `${size}px`;
      
      // Random direction vectors
      const angle = Math.random() * Math.PI * 2;
      const dist = Math.random() * 80 + 40;
      const dx = Math.cos(angle) * dist;
      const dy = Math.sin(angle) * dist;
      
      p.style.setProperty("--dx", `${dx}px`);
      p.style.setProperty("--dy", `${dy}px`);
      
      DOM.particleEmitter.appendChild(p);
      
      // Garbage collect particle after animation completes
      setTimeout(() => p.remove(), 600);
    }
  }

  // Manual Click Mining Trigger
  let manualDigProgress = 0;
  function triggerManualDig() {
    if (!gameState.isActive || gameState.isPaused || isManualDigging) return;
    
    isManualDigging = true;
    DOM.btnManualDig.disabled = true;
    manualDigProgress = 0;
    
    DOM.digProgressText.textContent = "Digging...";
    playSynthSound("click");
    
    // Progress bar tick interval
    function progressTick() {
      if (!gameState.isActive) return;
      if (gameState.isPaused) {
        manualDigTimeoutId = setTimeout(progressTick, 50);
        return;
      }
      
      manualDigProgress += 50;
      const pct = Math.min(100, (manualDigProgress / manualDigDuration) * 100);
      DOM.digProgressBar.style.width = `${pct}%`;
      
      if (manualDigProgress < manualDigDuration) {
        manualDigTimeoutId = setTimeout(progressTick, 50);
      } else {
        // Complete dig
        DOM.digProgressBar.style.width = "0%";
        DOM.digProgressText.textContent = "Ready to Dig";
        isManualDigging = false;
        DOM.btnManualDig.disabled = false;
        
        executeSuccessfulDig("Manual Shovel");
      }
    }
    
    progressTick();
  }

  DOM.btnManualDig.addEventListener("click", triggerManualDig);
  DOM.asteroidTarget.addEventListener("click", triggerManualDig);

  // Excavator / Drill automated toggling
  DOM.btnToggleExcavator.addEventListener("click", () => {
    const isTitan = (gameState.currentAsteroid === "TITAN");
    const isVoid = (gameState.currentAsteroid === "VOID");
    const machineOwned = isTitan ? gameState.hasExcavator : (isVoid ? gameState.hasDrillMachine : gameState.hasDrillBore);
    if (!gameState.isActive || gameState.isPaused || !machineOwned) return;
    
    gameState.excavatorRunning = !gameState.excavatorRunning;
    playSynthSound("click");
    
    if (gameState.excavatorRunning) {
      let stopText = "🛑 Stop Excavator";
      let startLog = "Heavy machinery activated. Excavator running...";
      if (isVoid) {
        stopText = "🛑 Stop Drill Machine";
        startLog = "Heavy machinery activated. Diamond Drilling Machine running...";
      } else if (gameState.currentAsteroid === "NOVA") {
        stopText = "🛑 Stop Drill-Bore";
        startLog = "Heavy machinery activated. Quantum Drill-Bore running...";
      }
      DOM.btnToggleExcavator.textContent = stopText;
      DOM.btnToggleExcavator.classList.add("btn-pulse");
      DOM.btnToggleExcavator.classList.add("active-running");
      writeConsoleLog(startLog);
    } else {
      let startText = "🤖 Run Excavator";
      let stopLog = "Heavy machinery deactivated. Excavator idle.";
      if (isVoid) {
        startText = "🤖 Run Drill Machine";
        stopLog = "Heavy machinery deactivated. Diamond Drilling Machine idle.";
      } else if (gameState.currentAsteroid === "NOVA") {
        startText = "🤖 Run Drill-Bore";
        stopLog = "Heavy machinery deactivated. Quantum Drill-Bore idle.";
      }
      DOM.btnToggleExcavator.textContent = startText;
      DOM.btnToggleExcavator.classList.remove("btn-pulse");
      DOM.btnToggleExcavator.classList.remove("active-running");
      writeConsoleLog(stopLog);
    }
  });

  // Calculate Net Worth: Cash + inventory prices
  function calculateCurrentNetWorth() {
    let inventoryVal = 0;
    excelModel.ores.forEach(ore => {
      inventoryVal += gameState.inventory[ore.id] * ore.currentPrice;
    });
    return gameState.walletCash + inventoryVal;
  }

  // Main 100ms clock ticking loop
  let timeAccumulator = 0;
  function gameLoopTick() {
    if (!gameState.isActive || gameState.isPaused) return;
    
    timeAccumulator += 100;
    
    // A. Excavator / Drill run time & fuel/gas consumption
    if (gameState.excavatorRunning) {
      const isTitan = (gameState.currentAsteroid === "TITAN");
      const isVoid = (gameState.currentAsteroid === "VOID");
      const machineOwned = isTitan ? gameState.hasExcavator : (isVoid ? gameState.hasDrillMachine : gameState.hasDrillBore);
      const fuelLevel = isTitan ? gameState.fuelLitres : (isVoid ? gameState.gasUnits : gameState.tyliumUnits);
      
      if (!machineOwned) {
        gameState.excavatorRunning = false;
      }
      else if (fuelLevel <= 0) {
        // Run out of fuel / gas
        gameState.excavatorRunning = false;
        if (isTitan) {
          gameState.fuelLitres = 0;
          DOM.btnToggleExcavator.textContent = "🤖 Run Excavator";
          writeConsoleLog("⚠️ EXCAVATOR ENGINE HALTED: OUT OF FUEL! Purchase fuel from outpost.");
        } else if (isVoid) {
          gameState.gasUnits = 0;
          DOM.btnToggleExcavator.textContent = "🤖 Run Drill Machine";
          writeConsoleLog("⚠️ DRILL ENGINE HALTED: OUT OF GAS! Purchase gas from outpost.");
        } else {
          gameState.tyliumUnits = 0;
          DOM.btnToggleExcavator.textContent = "🤖 Run Drill-Bore";
          writeConsoleLog("⚠️ DRILL-BORE ENGINE HALTED: OUT OF TYLIUM! Purchase Tylium from outpost.");
        }
        DOM.btnToggleExcavator.classList.remove("btn-pulse");
        DOM.btnToggleExcavator.classList.remove("active-running");
        playSynthSound("error");
      } else {
        // Consume fuel: 0.1 per second -> 0.01 per 100ms
        if (isTitan) {
          gameState.fuelLitres -= 0.01;
          if (gameState.fuelLitres < 0) gameState.fuelLitres = 0;
        } else if (isVoid) {
          gameState.gasUnits -= 0.01;
          if (gameState.gasUnits < 0) gameState.gasUnits = 0;
        } else {
          gameState.tyliumUnits -= 0.01;
          if (gameState.tyliumUnits < 0) gameState.tyliumUnits = 0;
        }
        
        // Automated dig rate
        // Default rate = 1 dig per 1.0s (1000ms). Boosted by loosener = 1 dig per 0.67s (667ms).
        const rateLimit = (isTitan && gameState.hasDirtLoosener) ? 667 : 1000;
        gameState.excavatorProgress += 100;
        if (gameState.excavatorProgress >= rateLimit) {
          gameState.excavatorProgress -= rateLimit;
          let machineName = "Auto Excavator";
          if (isVoid) machineName = "Diamond Drill";
          if (gameState.currentAsteroid === "NOVA") machineName = "Quantum Drill-Bore";
          executeSuccessfulDig(machineName);
        }
      }
      updateUI();
    }
    
    // B. Run 1-second interval items
    if (timeAccumulator >= 1000) {
      timeAccumulator -= 1000;
      
      // Decrement main timer
      gameState.timer--;
      if (gameState.timer <= 0) {
        gameState.timer = 0;
        endGameMission();
      }
      
      // Low time warning beep
      if (gameState.timer > 0 && gameState.timer <= 10) {
        playSynthSound("warning");
      }

      // Space Storm Scheduler (elapsed seconds increment)
      gameState.elapsedSeconds++;
      
      const windowSize = (gameState.currentAsteroid === "NOVA") ? 180 : 360;
      const windowIndex = Math.floor(gameState.elapsedSeconds / windowSize);
      if (gameState.scheduledStormSecond === null) {
        const minOffset = 15;
        const maxOffset = windowSize - 15;
        const randomOffset = minOffset + Math.floor(Math.random() * (maxOffset - minOffset));
        gameState.scheduledStormSecond = windowIndex * windowSize + randomOffset;
        gameState.stormWarningStarted = false;
        gameState.stormActive = false;
      }
      
      // Cross window boundary reset
      if (gameState.elapsedSeconds > (windowIndex + 1) * windowSize) {
        gameState.scheduledStormSecond = null;
      }
      
      // 2. 10-second pre-storm warning
      if (gameState.elapsedSeconds === gameState.scheduledStormSecond - 10) {
        gameState.stormWarningStarted = true;
        gameState.stormShieldActive = false;
        
        if (DOM.stormPanel) {
          DOM.stormPanel.classList.remove("hidden");
          const msgEl = document.getElementById("storm-warning-message");
          if (msgEl) msgEl.textContent = "⚠️ WARNING: DUST STORM IN 10s!";
          
          if (gameState.hasStormProtection) {
            DOM.btnLockShield.disabled = false;
            DOM.btnLockShield.textContent = "🔒 ACTIVATE STORM SHIELD";
            DOM.btnLockShield.style.background = "";
          } else {
            const cost = (gameState.currentAsteroid === "NOVA") ? excelModel.novaStormProtectionCost : 25000;
            DOM.btnLockShield.disabled = true;
            DOM.btnLockShield.textContent = `🔒 SHIELD LOCKED (Storm Gear Needed - $${cost.toLocaleString()})`;
            DOM.btnLockShield.style.background = "";
          }
        }
        playSynthSound("warning");
        writeConsoleLog("⚠️ RADAR WARNING: A massive space dust storm is approaching the Outpost! 10 seconds to impact!");
      }
      // Update warning countdown if warning is active
      else if (gameState.stormWarningStarted && gameState.elapsedSeconds < gameState.scheduledStormSecond) {
        const diff = gameState.scheduledStormSecond - gameState.elapsedSeconds;
        const msgEl = document.getElementById("storm-warning-message");
        if (msgEl) msgEl.textContent = `⚠️ WARNING: DUST STORM IN ${diff}s!`;
        playSynthSound("warning");
      }
      // 3. Storm impact!
      else if (gameState.elapsedSeconds === gameState.scheduledStormSecond) {
        gameState.stormWarningStarted = false;
        if (DOM.stormPanel) DOM.stormPanel.classList.add("hidden");
        
        triggerSpaceDustStorm();
        
        // Pre-schedule storm for the next window
        const nextWindowIdx = windowIndex + 1;
        const minOffset = 15;
        const maxOffset = windowSize - 15;
        const randomOffset = minOffset + Math.floor(Math.random() * (maxOffset - minOffset));
        gameState.scheduledStormSecond = nextWindowIdx * windowSize + randomOffset;
      }
      
      // Earthquake scheduler (Only on NOVA)
      if (gameState.currentAsteroid === "NOVA") {
        gameState.elapsedEarthquakeSeconds++;
        
        const eqWindowSize = 240; // 4 minutes
        const eqWindowIndex = Math.floor(gameState.elapsedEarthquakeSeconds / eqWindowSize);
        if (gameState.scheduledEarthquakeSecond === null) {
          const minOffset = 15;
          const maxOffset = eqWindowSize - 15;
          const randomOffset = minOffset + Math.floor(Math.random() * (maxOffset - minOffset));
          gameState.scheduledEarthquakeSecond = eqWindowIndex * eqWindowSize + randomOffset;
          gameState.earthquakeWarningStarted = false;
          gameState.earthquakeActive = false;
        }
        
        // Cross window boundary reset
        if (gameState.elapsedEarthquakeSeconds > (eqWindowIndex + 1) * eqWindowSize) {
          gameState.scheduledEarthquakeSecond = null;
        }
        
        // 10-second pre-earthquake warning
        if (gameState.elapsedEarthquakeSeconds === gameState.scheduledEarthquakeSecond - 10) {
          gameState.earthquakeWarningStarted = true;
          
          if (DOM.seismicPanel) {
            DOM.seismicPanel.classList.remove("hidden");
            const msgEl = document.getElementById("seismic-warning-message");
            if (msgEl) msgEl.textContent = "⚠️ SEISMIC WARNING: REGULAR TREMORS DETECTED! EARTHQUAKE IN 10s!";
          }
          playSynthSound("warning");
          writeConsoleLog("⚠️ SEISMIC DETECTOR: Micro-tremors detected on NOVA! Earthquake warning active. Sell ores now!");
        }
        // Update warning countdown if active
        else if (gameState.earthquakeWarningStarted && gameState.elapsedEarthquakeSeconds < gameState.scheduledEarthquakeSecond) {
          const diff = gameState.scheduledEarthquakeSecond - gameState.elapsedEarthquakeSeconds;
          const msgEl = document.getElementById("seismic-warning-message");
          if (msgEl) msgEl.textContent = `⚠️ SEISMIC WARNING: EARTHQUAKE IN ${diff}s!`;
          playSynthSound("warning");
        }
        // Earthquake impact!
        else if (gameState.elapsedEarthquakeSeconds === gameState.scheduledEarthquakeSecond) {
          gameState.earthquakeWarningStarted = false;
          if (DOM.seismicPanel) DOM.seismicPanel.classList.add("hidden");
          
          triggerSeismicEarthquake();
          
          // Pre-schedule earthquake for the next window
          const nextWindowIdx = eqWindowIndex + 1;
          const minOffset = 15;
          const maxOffset = eqWindowSize - 15;
          const randomOffset = minOffset + Math.floor(Math.random() * (maxOffset - minOffset));
          gameState.scheduledEarthquakeSecond = nextWindowIdx * eqWindowSize + randomOffset;
        }
      } else {
        gameState.elapsedEarthquakeSeconds = 0;
        gameState.scheduledEarthquakeSecond = null;
        gameState.earthquakeWarningStarted = false;
        gameState.earthquakeActive = false;
        if (DOM.seismicPanel) DOM.seismicPanel.classList.add("hidden");
      }
      
      // Increment appreciation countdown (10s on all asteroids)
      gameState.appreciationTimer--;
      if (gameState.appreciationTimer <= 0) {
        gameState.appreciationTimer = 10;
        gameState.appreciationTicks++;
        
        // Perform ores price appreciation
        appreciateOrePrices();
      }
      
      // Save wealth logs for graphs (every 2s)
      if (gameState.timer % 2 === 0) {
        gameState.history.timestamps.push(excelModel.initialTimer - gameState.timer);
        gameState.history.netWorths.push(calculateCurrentNetWorth());
        drawNetWorthChart();
      }
      
      updateUI();
    }
  }

  // Appreciate prices of held items
  function appreciateOrePrices() {
    let appreciatedOresCount = 0;
    
    excelModel.ores.forEach(ore => {
      if (ore.appreciationRate > 0) {
        const oldValue = ore.currentPrice;
        // Linear appreciation based on E_i
        ore.currentPrice = ore.baseValue * (1.0 + (gameState.appreciationTicks * ore.appreciationRate));
        appreciatedOresCount++;
      }
    });
    
    if (appreciatedOresCount > 0) {
      playSynthSound("price_up");
      
      let oreNamesList = "";
      if (gameState.currentAsteroid === "TITAN") {
        oreNamesList = "Gold, Amethyst, Ruby";
      } else if (gameState.currentAsteroid === "VOID") {
        oreNamesList = "Pyrite, Diamond, Emerald";
      } else {
        oreNamesList = "Obsidian, Warp-Stone, Infinity Gem";
      }
      writeConsoleLog(`📈 RESOURCE APPRAISAL: Ore values appreciated! ${oreNamesList} values increased.`);
      
      // Visual flash on inventory rows
      ["gold", "amethyst", "ruby"].forEach(id => {
        const row = document.getElementById(`ore-item-${id}`);
        if (row) {
          row.style.background = "rgba(0, 230, 118, 0.15)";
          row.style.borderColor = "var(--color-green)";
          setTimeout(() => {
            row.style.background = null;
            row.style.borderColor = null;
          }, 600);
        }
      });
    }
  }

  // Final mission timer expiration
  function endGameMission(timerExpired = true) {
    gameState.isActive = false;
    if (gameIntervalId) clearInterval(gameIntervalId);
    
    const isTitan = (gameState.currentAsteroid === "TITAN");
    const isVoid = (gameState.currentAsteroid === "VOID");
    const isNova = (gameState.currentAsteroid === "NOVA");
    
    // Stop sound
    gameState.excavatorRunning = false;
    if (isTitan) DOM.btnToggleExcavator.textContent = "🤖 Run Excavator";
    else if (isVoid) DOM.btnToggleExcavator.textContent = "🤖 Run Drill Machine";
    else DOM.btnToggleExcavator.textContent = "🤖 Run Drill-Bore";
    DOM.btnToggleExcavator.classList.remove("btn-pulse");
    DOM.btnToggleExcavator.classList.remove("active-running");
    
    // Stop digging Shovel
    isManualDigging = false;
    DOM.btnManualDig.disabled = false;
    DOM.digProgressBar.style.width = "0%";
    DOM.digProgressText.textContent = "Expedition Over";
    
    // Sell all inventory automatically to compute final cash
    let inventoryVal = 0;
    excelModel.ores.forEach(ore => {
      inventoryVal += gameState.inventory[ore.id] * ore.currentPrice;
    });
    
    // Sell all inventory to cash first
    sellAllInventory(true);
    
    let finalNetWorth = gameState.walletCash;
    if (isNova && timerExpired) {
      finalNetWorth = Math.max(0, finalNetWorth - 50000000);
    }
    playSynthSound("gameover");
    
    // Get Pilot Name
    const pilotName = DOM.pilotNameInput ? DOM.pilotNameInput.value.trim() : "Miner One";
    
    // Check personal high score in LocalStorage
    let highScore = parseFloat(safeStorage.getItem("asteroid_rush_highscore") || "0");
    let isNewHighScore = false;
    if (finalNetWorth > highScore) {
      highScore = finalNetWorth;
      safeStorage.setItem("asteroid_rush_highscore", finalNetWorth.toString());
      isNewHighScore = true;
    }
    
    // Load data in End Game Dialog
    const successMsgEl = document.querySelector(".end-game-dialog .success-message");
    const dialogTitleEl = document.querySelector(".end-game-dialog .dialog-title");
    if (successMsgEl) {
      if (timerExpired) {
        if (dialogTitleEl) dialogTitleEl.textContent = "🚀 Asteroid Mission Completed!";
        successMsgEl.textContent = "Your excavation timer has expired. Telemetry analysis complete.";
      } else {
        if (dialogTitleEl) dialogTitleEl.textContent = "🌍 Returned to Earth!";
        successMsgEl.innerHTML = `🌍 <strong>GAME OVER</strong><br>Congratulations on completing your space exploration adventure! You have safely returned to Earth with final winnings of <strong>$${Math.round(finalNetWorth).toLocaleString()}</strong>!`;
      }
    }
    DOM.endNetWorth.textContent = `$${Math.round(finalNetWorth).toLocaleString()}`;
    DOM.endCash.textContent = `$${Math.round(gameState.walletCash).toLocaleString()}`;
    DOM.endInventoryVal.textContent = `$0`;
    DOM.endDigs.textContent = gameState.digsCompleted;
    
    // Dynamically update label for Coins/Pearls/Crystals in End Game dialog
    const endCoinsLabelEl = DOM.endCoins.parentElement.querySelector("span");
    if (endCoinsLabelEl) {
      if (isTitan) {
        endCoinsLabelEl.textContent = "Total Coins Found:";
        DOM.endCoins.textContent = gameState.walletCoins + gameState.tallies.coins;
      } else if (isVoid) {
        endCoinsLabelEl.textContent = "Total Pearls Found:";
        DOM.endCoins.textContent = gameState.walletPearls + gameState.tallies.pearls;
      } else {
        endCoinsLabelEl.textContent = "Total Crystals Found:";
        DOM.endCoins.textContent = gameState.walletCrystals + gameState.tallies.crystals;
      }
    }
    
    if (isNewHighScore) {
      DOM.highScoreBanner.classList.remove("hidden");
    } else {
      DOM.highScoreBanner.classList.add("hidden");
    }
    
    // Save to Leaderboard
    let secTally = 0;
    if (isTitan) secTally = gameState.walletCoins + gameState.tallies.coins;
    else if (isVoid) secTally = gameState.walletPearls + gameState.tallies.pearls;
    else secTally = gameState.walletCrystals + gameState.tallies.crystals;
    
    saveLeaderboardEntry(pilotName, finalNetWorth, gameState.walletCash, secTally, gameState.digsCompleted);
    
    // Show overlay
    DOM.endGameModal.classList.remove("hidden");
    writeConsoleLog(`🚀 MISSION COMPLETED! Final Net Worth: $${Math.round(finalNetWorth).toLocaleString()}.`);
  }

  // restart game trigger
  DOM.btnRestartGame.addEventListener("click", () => {
    startNewGame();
  });

  // ==========================================================================
  // 6. TRADING OPERATIONS (BUY/SELL)
  // ==========================================================================
  
  // Sell individual ore by index
  function sellOreItemByIndex(index) {
    if (gameState.isPaused) return;
    const ore = excelModel.ores[index];
    if (gameState.inventory[ore.id] <= 0) return;
    
    gameState.inventory[ore.id]--;
    gameState.walletCash += ore.currentPrice;
    
    playSynthSound("sell");
    updateUI();
  }

  // Sell all inventory
  function sellAllInventory(bypassPause = false) {
    if (gameState.isPaused && !bypassPause) return;
    let totalEarnings = 0;
    let itemsSold = 0;
    
    excelModel.ores.forEach(ore => {
      const count = gameState.inventory[ore.id];
      if (count > 0) {
        totalEarnings += count * ore.currentPrice;
        itemsSold += count;
        gameState.inventory[ore.id] = 0;
      }
    });
    
    if (itemsSold > 0) {
      gameState.walletCash += totalEarnings;
      playSynthSound("sell");
      writeConsoleLog(`💰 Sold ${itemsSold} ores for a total of $${Math.round(totalEarnings).toLocaleString()} cash.`);
      updateUI();
    } else {
      playSynthSound("error");
    }
  }

  // Wire up sell actions
  DOM.btnSellAll.addEventListener("click", sellAllInventory);
  DOM.btnSellIron.addEventListener("click", () => sellOreItemByIndex(0));
  DOM.btnSellCopper.addEventListener("click", () => sellOreItemByIndex(1));
  DOM.btnSellGold.addEventListener("click", () => sellOreItemByIndex(2));
  DOM.btnSellAmethyst.addEventListener("click", () => sellOreItemByIndex(3));
  DOM.btnSellRuby.addEventListener("click", () => sellOreItemByIndex(4));

  // Shop item purchases
  DOM.btnBuyExcavator.addEventListener("click", () => {
    if (gameState.isPaused) return;
    if (gameState.currentAsteroid === "TITAN") {
      if (gameState.hasExcavator) return;
      if (gameState.walletCash >= excelModel.excavatorCost) {
        gameState.walletCash -= excelModel.excavatorCost;
        gameState.hasExcavator = true;
        
        DOM.btnToggleExcavator.disabled = false;
        DOM.btnToggleExcavator.classList.remove("btn-locked");
        DOM.btnToggleExcavator.textContent = "🤖 Run Excavator";
        
        playSynthSound("upgrade");
        writeConsoleLog("🛠️ SHOP TRANSACTION: Purchased Auto Excavator! Equip fuel to run.");
        updateUI();
      } else {
        playSynthSound("error");
        writeConsoleLog("❌ SHOP ERROR: Insufficient cash to purchase Excavator!");
      }
    } else if (gameState.currentAsteroid === "VOID") {
      // VOID: Diamond Drilling Machine
      if (gameState.hasDrillMachine) return;
      if (gameState.walletCash >= excelModel.drillCost) {
        gameState.walletCash -= excelModel.drillCost;
        gameState.hasDrillMachine = true;
        
        DOM.btnToggleExcavator.disabled = false;
        DOM.btnToggleExcavator.classList.remove("btn-locked");
        DOM.btnToggleExcavator.textContent = "🤖 Run Drill Machine";
        
        playSynthSound("upgrade");
        writeConsoleLog("🛠️ SHOP TRANSACTION: Purchased Diamond Drilling Machine! Equip gas to run.");
        updateUI();
      } else {
        playSynthSound("error");
        writeConsoleLog("❌ SHOP ERROR: Insufficient cash to purchase Diamond Drilling Machine!");
      }
    } else if (gameState.currentAsteroid === "NOVA") {
      // NOVA: Quantum Lattice Drill-Bore
      if (gameState.hasDrillBore) return;
      if (gameState.walletCash >= excelModel.drillBoreCost) {
        gameState.walletCash -= excelModel.drillBoreCost;
        gameState.hasDrillBore = true;
        
        DOM.btnToggleExcavator.disabled = false;
        DOM.btnToggleExcavator.classList.remove("btn-locked");
        DOM.btnToggleExcavator.textContent = "🤖 Run Drill-Bore";
        
        playSynthSound("upgrade");
        writeConsoleLog("🛠️ SHOP TRANSACTION: Purchased Quantum Lattice Drill-Bore! Equip Tylium to run.");
        updateUI();
      } else {
        playSynthSound("error");
        writeConsoleLog("❌ SHOP ERROR: Insufficient cash to purchase Quantum Lattice Drill-Bore!");
      }
    }
  });

  DOM.btnBuyFuel.addEventListener("click", () => {
    if (gameState.isPaused) return;
    if (gameState.currentAsteroid === "TITAN") {
      if (gameState.walletCash >= excelModel.fuelCost) {
        gameState.walletCash -= excelModel.fuelCost;
        gameState.fuelLitres += 1.0;
        
        playSynthSound("upgrade");
        writeConsoleLog("⛽ SHOP TRANSACTION: Purchased 1L fuel. Heavy machinery excavation time extended.");
        updateUI();
      } else {
        playSynthSound("error");
        writeConsoleLog("❌ SHOP ERROR: Insufficient cash to purchase Fuel!");
      }
    } else if (gameState.currentAsteroid === "VOID") {
      // VOID: Gas
      if (gameState.walletCash >= excelModel.gasCost) {
        gameState.walletCash -= excelModel.gasCost;
        gameState.gasUnits += 1.0;
        
        playSynthSound("upgrade");
        writeConsoleLog("🔋 SHOP TRANSACTION: Purchased 1 Unit Gas. Diamond drilling machine run time extended.");
        updateUI();
      } else {
        playSynthSound("error");
        writeConsoleLog("❌ SHOP ERROR: Insufficient cash to purchase Gas!");
      }
    } else if (gameState.currentAsteroid === "NOVA") {
      // NOVA: Tylium Fuel
      if (gameState.walletCash >= excelModel.tyliumCost) {
        gameState.walletCash -= excelModel.tyliumCost;
        gameState.tyliumUnits += 1.0;
        
        playSynthSound("upgrade");
        writeConsoleLog("🔋 SHOP TRANSACTION: Purchased 1 Unit Tylium Fuel. Quantum Drill-Bore run time extended.");
        updateUI();
      } else {
        playSynthSound("error");
        writeConsoleLog("❌ SHOP ERROR: Insufficient cash to purchase Tylium Fuel!");
      }
    }
  });

  DOM.btnBuyShovel.addEventListener("click", () => {
    if (gameState.isPaused) return;
    if (gameState.currentAsteroid === "TITAN") {
      if (gameState.hasTitaniumShovel) return;
      if (gameState.walletCoins >= excelModel.gear.shovel.costCoins) {
        gameState.walletCoins -= excelModel.gear.shovel.costCoins;
        gameState.hasTitaniumShovel = true;
        
        DOM.telemetryShovel.textContent = "Titanium Shovel";
        
        playSynthSound("upgrade");
        writeConsoleLog("⛏️ GEAR EQUIPPED: Equipped Shiny Titanium Shovel! Manual dig speed doubled.");
        updateUI();
      } else {
        playSynthSound("error");
        writeConsoleLog("❌ SHOP ERROR: Insufficient coins to purchase Shiny Titanium Shovel!");
      }
    } else if (gameState.currentAsteroid === "VOID") {
      // VOID: Shiny Pickaxe
      if (gameState.hasShinyPickaxe) return;
      if (gameState.walletPearls >= excelModel.voidGear.pickaxe.costPearls) {
        gameState.walletPearls -= excelModel.voidGear.pickaxe.costPearls;
        gameState.hasShinyPickaxe = true;
        
        DOM.telemetryShovel.textContent = "Shiny Pickaxe (VOID)";
        
        playSynthSound("upgrade");
        writeConsoleLog("⛏️ GEAR EQUIPPED: Equipped Shiny Pickaxe! Manual dig speed increased dramatically (750ms).");
        updateUI();
      } else {
        playSynthSound("error");
        writeConsoleLog("❌ SHOP ERROR: Insufficient pearls to purchase Shiny Pickaxe!");
      }
    } else if (gameState.currentAsteroid === "NOVA") {
      // NOVA: Sledgehammer
      if (gameState.hasSledgehammer) return;
      const cost = excelModel.novaGear.sledgehammer.costCrystals;
      if (gameState.walletCrystals >= cost) {
        gameState.walletCrystals -= cost;
        gameState.hasSledgehammer = true;
        
        DOM.telemetryShovel.textContent = "Sledgehammer (NOVA)";
        
        playSynthSound("upgrade");
        writeConsoleLog("⛏️ GEAR EQUIPPED: Equipped Sledgehammer! Manual dig speed increased dramatically (500ms).");
        updateUI();
      } else {
        playSynthSound("error");
        writeConsoleLog(`❌ SHOP ERROR: Insufficient crystals to purchase Sledgehammer (${cost} Crystals required)!`);
      }
    }
  });

  DOM.btnBuyLoosener.addEventListener("click", () => {
    if (gameState.isPaused) return;
    if (gameState.currentAsteroid === "TITAN") {
      if (gameState.hasDirtLoosener) return;
      if (gameState.walletCoins >= excelModel.gear.loosener.costCoins) {
        gameState.walletCoins -= excelModel.gear.loosener.costCoins;
        gameState.hasDirtLoosener = true;
        
        DOM.telemetryLoosener.textContent = "Loosener (+50% Speed)";
        
        playSynthSound("upgrade");
        writeConsoleLog("🌀 GEAR EQUIPPED: Equipped Dirt Loosener! Digging frequency increased by 50% for all tools.");
        updateUI();
      } else {
        playSynthSound("error");
        writeConsoleLog("❌ SHOP ERROR: Insufficient coins to purchase Dirt Loosener!");
      }
    } else if (gameState.currentAsteroid === "VOID") {
      // VOID: Dynamite & Ignition
      const cost = (gameState.dynamitePurchasedCount === 0) ? 25 : 50;
      if (gameState.walletPearls >= cost) {
        gameState.walletPearls -= cost;
        gameState.dynamitePurchasedCount++;
        excelModel.voidGear.dynamite.costPearls = 50; // Subsequent dynamite always costs 50
        playSynthSound("sell");
        writeConsoleLog(`💥 DYNAMITE PURCHASED (Cost: ${cost} Pearls): Initiating detonation sequence...`);
        
        gameState.isPaused = true;
        triggerDynamiteExplosion();
      } else {
        playSynthSound("error");
        writeConsoleLog(`❌ SHOP ERROR: Insufficient pearls to purchase Dynamite & Ignition (${cost} Pearls needed)!`);
      }
    } else if (gameState.currentAsteroid === "NOVA") {
      // NOVA: Nuclear Explosion
      if (gameState.nukePurchasedCount >= 2) {
        playSynthSound("error");
        writeConsoleLog("❌ SHOP ERROR: Maximum of 2 Nuclear Explosions allowed on NOVA!");
        return;
      }
      const cost = (gameState.nukePurchasedCount === 0) ? 30 : 60;
      if (gameState.walletCrystals >= cost) {
        gameState.walletCrystals -= cost;
        gameState.nukePurchasedCount++;
        excelModel.novaGear.nuke.costCrystals = 60; // Subsequent nuke always costs 60
        playSynthSound("sell");
        writeConsoleLog(`☢️ NUCLEAR EXPLOSION PURCHASED (Cost: ${cost} Crystals): Initiating blast sequence...`);
        
        gameState.isPaused = true;
        triggerNuclearExplosion();
      } else {
        playSynthSound("error");
        writeConsoleLog(`❌ SHOP ERROR: Insufficient crystals to purchase Nuclear Explosion (${cost} Crystals needed)!`);
      }
    }
  });

  DOM.btnBuyDetector.addEventListener("click", () => {
    if (gameState.isPaused) return;
    if (gameState.currentAsteroid === "TITAN") {
      if (gameState.hasMetalDetector) return;
      if (gameState.walletCoins >= excelModel.gear.detector.costCoins) {
        gameState.walletCoins -= excelModel.gear.detector.costCoins;
        gameState.hasMetalDetector = true;
        
        DOM.telemetryDetector.textContent = "Equipped (3x Rares)";
        
        playSynthSound("upgrade");
        writeConsoleLog("🔍 GEAR EQUIPPED: Equipped Metal Detector! Probability of finding Gold, Amethyst, and Ruby drops tripled!");
        updateUI();
      } else {
        playSynthSound("error");
        writeConsoleLog("❌ SHOP ERROR: Insufficient coins to purchase Metal Detector!");
      }
    } else if (gameState.currentAsteroid === "NOVA") {
      // NOVA: Resonance Harmonic Stabilizer
      if (gameState.hasResonanceStabilizer) return;
      const cost = excelModel.novaGear.stabilizer.costCrystals;
      if (gameState.walletCrystals >= cost) {
        gameState.walletCrystals -= cost;
        gameState.hasResonanceStabilizer = true;
        
        DOM.telemetryDetector.textContent = "Equipped (2x Rares)";
        
        playSynthSound("upgrade");
        writeConsoleLog("🔮 GEAR EQUIPPED: Equipped Resonance Harmonic Stabilizer! Probability of finding Obsidian, Warp-Stone, and Infinity Gems doubled!");
        updateUI();
      } else {
        playSynthSound("error");
        writeConsoleLog(`❌ SHOP ERROR: Insufficient crystals to purchase Resonance Harmonic Stabilizer (${cost} Crystals required)!`);
      }
    }
  });

  DOM.btnBuyTime.addEventListener("click", () => {
    if (gameState.isPaused) return;
    if (gameState.currentAsteroid === "TITAN") {
      if (gameState.walletCoins >= excelModel.gear.time.costCoins) {
        gameState.walletCoins -= excelModel.gear.time.costCoins;
        gameState.timer += 60;
        
        playSynthSound("upgrade");
        writeConsoleLog("⏳ TELEPORT STRETCH: Extended mission clock by +60 seconds.");
        updateUI();
      } else {
        playSynthSound("error");
        writeConsoleLog("❌ SHOP ERROR: Insufficient coins to purchase Time Extender!");
      }
    } else if (gameState.currentAsteroid === "VOID") {
      // VOID: Time Extender
      if (gameState.walletPearls >= excelModel.voidGear.time.costPearls) {
        gameState.walletPearls -= excelModel.voidGear.time.costPearls;
        gameState.timer += 60;
        
        playSynthSound("upgrade");
        writeConsoleLog("⏳ VOID EXPEDITION: Extended mission clock by +60 seconds.");
        updateUI();
      } else {
        playSynthSound("error");
        writeConsoleLog("❌ SHOP ERROR: Insufficient pearls to purchase Time Extender!");
      }
    } else if (gameState.currentAsteroid === "NOVA") {
      // NOVA: Cosmic Water
      if (gameState.hasCosmicWater) return;
      const cost = excelModel.novaGear.water.costCrystals;
      if (gameState.walletCrystals >= cost) {
        gameState.walletCrystals -= cost;
        gameState.hasCosmicWater = true;
        
        DOM.telemetryLoosener.textContent = "Cosmic Water (+50% Speed)";
        
        playSynthSound("upgrade");
        writeConsoleLog("💧 GEAR EQUIPPED: Consumed Cosmic Water! Manual digging speed increased by 50%.");
        updateUI();
      } else {
        playSynthSound("error");
        writeConsoleLog(`❌ SHOP ERROR: Insufficient crystals to purchase Cosmic Water (${cost} Crystals required)!`);
      }
    }
  });

  DOM.btnBuyStormProtection.addEventListener("click", () => {
    if (gameState.isPaused) return;
    if (gameState.hasStormProtection) return;
    
    let cost = 25000;
    if (gameState.currentAsteroid === "NOVA") {
      cost = excelModel.novaStormProtectionCost;
    } else if (gameState.currentAsteroid === "TITAN") {
      cost = 10000;
    } else if (gameState.currentAsteroid === "VOID") {
      cost = excelModel.stormProtectionCost;
    }
    if (gameState.walletCash >= cost) {
      gameState.walletCash -= cost;
      gameState.hasStormProtection = true;
      playSynthSound("upgrade");
      writeConsoleLog("🛡️ PROTECTION GEAR PURCHASED: Base storm protection acquired! Enable the shield during dust storms.");
      updateUI();
    } else {
      playSynthSound("error");
      writeConsoleLog(`❌ SHOP ERROR: Insufficient cash for Storm Protection Gear ($${cost.toLocaleString()} required)!`);
    }
  });

  DOM.btnLockShield.addEventListener("click", () => {
    if (gameState.hasStormProtection && !gameState.stormShieldActive && gameState.stormWarningStarted) {
      gameState.stormShieldActive = true;
      DOM.btnLockShield.disabled = true;
      DOM.btnLockShield.style.background = "linear-gradient(135deg, #00e676 0%, #00b0ff 100%)";
      DOM.btnLockShield.textContent = "🛡️ STORM SHIELD ACTIVE!";
      playSynthSound("upgrade");
      writeConsoleLog("🛡️ SHIELD ACTIVATED: Base sealed. Ores, fuel, and machinery locked down.");
    }
  });

  function triggerDynamiteExplosion() {
    const expOverlay = document.getElementById("explosion-overlay");
    const expTitle = document.getElementById("explosion-title");
    const expIcon = document.getElementById("explosion-fuse-icon");
    
    if (expOverlay && expTitle && expIcon) {
      expOverlay.classList.remove("hidden");
      expOverlay.style.background = "#000";
      expIcon.textContent = "🧨";
      
      // t=0s
      expTitle.textContent = "DETONATING IN 3...";
      playSynthSound("warning");
      
      // t=1s
      setTimeout(() => {
        expTitle.textContent = "DETONATING IN 2...";
        playSynthSound("warning");
      }, 1000);
      
      // t=2s
      setTimeout(() => {
        expTitle.textContent = "DETONATING IN 1...";
        playSynthSound("warning");
      }, 2000);
      
      // t=3s
      setTimeout(() => {
        expOverlay.style.background = "linear-gradient(135deg, rgba(255, 69, 0, 0.95) 0%, rgba(255, 140, 0, 0.9) 100%)";
        expTitle.textContent = "💥 BOOM! DETONATION ACTIVE! 💥";
        expIcon.textContent = "💥";
        
        playHeavyExplosionSound();
        document.body.classList.add("explosion-shake");
        if (DOM.asteroidWrapper) DOM.asteroidWrapper.classList.add("explosion-shake");
        
        // Execute the digs
        executeDynamiteExplosionDigs();
      }, 3000);
      
      // t=4.5s
      setTimeout(() => {
        document.body.classList.remove("explosion-shake");
        if (DOM.asteroidWrapper) DOM.asteroidWrapper.classList.remove("explosion-shake");
        expOverlay.classList.add("hidden");
        expOverlay.style.background = "#000";
        expIcon.textContent = "🧨";
        
        gameState.isPaused = false;
        recalculateExcelModel();
        updateUI();
      }, 4500);
    } else {
      // Fallback if overlay elements don't exist
      executeDynamiteExplosionDigs();
      gameState.isPaused = false;
      updateUI();
    }
  }

  function executeDynamiteExplosionDigs() {
    const rollSummary = {};
    let pearlsFoundDuringExplosion = 0;

    for (let i = 0; i < 100; i++) {
      gameState.digsCompleted++;
      const drop = rollMiningReward();
      gameState.inventory[drop.id]++;
      gameState.tallies[drop.id]++;

      if (!rollSummary[drop.name]) {
        rollSummary[drop.name] = { count: 0, icon: drop.icon };
      }
      rollSummary[drop.name].count++;

      // Roll Pearls
      if (Math.random() < excelModel.pearlDropChance) {
        gameState.walletPearls++;
        gameState.tallies.pearls++;
        pearlsFoundDuringExplosion++;
      }
    }

    DOM.telemetryDigs.textContent = gameState.digsCompleted;

    writeConsoleLog("🔥 EXPLOSION STATUS: Detonation successful!");
    Object.keys(rollSummary).forEach(name => {
      const info = rollSummary[name];
      writeConsoleLog(`  - Gathered ${info.icon} ${info.count}x ${name}`);
    });
    if (pearlsFoundDuringExplosion > 0) {
      writeConsoleLog(`  - Recovered 🦪 ${pearlsFoundDuringExplosion} Pearls from crater fissures!`);
    }

    for (let p = 0; p < 15; p++) {
      setTimeout(() => {
        createStoneParticles("#d033ff");
        createStoneParticles("#2ecc71");
        createStoneParticles("#a0e6ff");
      }, p * 100);
    }
  }

  function triggerNuclearExplosion() {
    const expOverlay = document.getElementById("explosion-overlay");
    const expTitle = document.getElementById("explosion-title");
    const expIcon = document.getElementById("explosion-fuse-icon");
    
    if (expOverlay && expTitle && expIcon) {
      expOverlay.classList.remove("hidden");
      expOverlay.style.background = "#000";
      expIcon.textContent = "☢️";
      
      // t=0s
      expTitle.textContent = "☢️ NUKE DETONATING IN 3...";
      playSynthSound("warning");
      
      // t=1s
      setTimeout(() => {
        expTitle.textContent = "☢️ NUKE DETONATING IN 2...";
        playSynthSound("warning");
      }, 1000);
      
      // t=2s
      setTimeout(() => {
        expTitle.textContent = "☢️ NUKE DETONATING IN 1...";
        playSynthSound("warning");
      }, 2000);
      
      // t=3s
      setTimeout(() => {
        expOverlay.style.background = "linear-gradient(135deg, rgba(0, 229, 255, 0.95) 0%, rgba(208, 51, 255, 0.9) 100%)";
        expTitle.textContent = "☢️ BOOM! COLD WAR BLAST ACTIVE! ☢️";
        expIcon.textContent = "💥";
        
        playHeavyExplosionSound();
        document.body.classList.add("explosion-shake");
        if (DOM.asteroidWrapper) DOM.asteroidWrapper.classList.add("explosion-shake");
        
        // Execute the digs
        executeNuclearExplosionDigs();
      }, 3000);
      
      // t=4.5s
      setTimeout(() => {
        document.body.classList.remove("explosion-shake");
        if (DOM.asteroidWrapper) DOM.asteroidWrapper.classList.remove("explosion-shake");
        expOverlay.classList.add("hidden");
        expOverlay.style.background = "#000";
        expIcon.textContent = "🧨";
        
        gameState.isPaused = false;
        recalculateExcelModel();
        updateUI();
      }, 4500);
    } else {
      executeNuclearExplosionDigs();
      gameState.isPaused = false;
      updateUI();
    }
  }

  function executeNuclearExplosionDigs() {
    const rollSummary = {};
    let crystalsFoundDuringExplosion = 0;

    for (let i = 0; i < 1000; i++) {
      gameState.digsCompleted++;
      const drop = rollMiningReward();
      gameState.inventory[drop.id]++;
      gameState.tallies[drop.id]++;

      if (!rollSummary[drop.name]) {
        rollSummary[drop.name] = { count: 0, icon: drop.icon };
      }
      rollSummary[drop.name].count++;

      // Roll Crystals
      if (Math.random() < excelModel.crystalDropChance) {
        gameState.walletCrystals++;
        gameState.tallies.crystals++;
        crystalsFoundDuringExplosion++;
      }
    }

    DOM.telemetryDigs.textContent = gameState.digsCompleted;

    writeConsoleLog("☢️ NUKE STATUS: Cold war nuke successfully detonated on NOVA!");
    Object.keys(rollSummary).forEach(name => {
      const info = rollSummary[name];
      writeConsoleLog(`  - Gathered ${info.icon} ${info.count}x ${name}`);
    });
    if (crystalsFoundDuringExplosion > 0) {
      writeConsoleLog(`  - Recovered 🔮 ${crystalsFoundDuringExplosion} Crystals from seismic blast fissures!`);
    }

    for (let p = 0; p < 30; p++) {
      setTimeout(() => {
        createStoneParticles("#00e676");
        createStoneParticles("#00b0ff");
        createStoneParticles("#d033ff");
      }, p * 50);
    }
  }

  function triggerSpaceDustStorm() {
    gameState.stormActive = true;
    gameState.isPaused = true; // Pause clock and manual digging/buying
    
    // Show Storm Overlay
    const overlay = document.getElementById("storm-overlay");
    const desc = document.getElementById("storm-overlay-desc");
    if (overlay && desc) {
      overlay.classList.remove("hidden");
      desc.textContent = gameState.stormShieldActive ? "🛡️ STORM SHIELD HOLDING STRONG..." : "🌪️ BATTLING SOLAR WINDS - SYSTEM INTEGRITY COMPROMISED...";
    }
    
    // Play rumbles and shake screen
    playHeavyExplosionSound();
    document.body.classList.add("explosion-shake");
    if (DOM.asteroidWrapper) DOM.asteroidWrapper.classList.add("explosion-shake");
    
    writeConsoleLog("🌪️ SYSTEM ALERT: A massive space dust storm is currently hitting the asteroid Outpost!");
    
    // Duration is 4.5 seconds
    setTimeout(() => {
      // Clear overlay and screen shake
      document.body.classList.remove("explosion-shake");
      if (DOM.asteroidWrapper) DOM.asteroidWrapper.classList.remove("explosion-shake");
      if (overlay) overlay.classList.add("hidden");
      
      // Resolve storm damage
      if (gameState.stormShieldActive) {
        playSynthSound("upgrade");
        writeConsoleLog("🛡️ SHIELD SECURED: Space dust storm passed safely. All machinery and ores protected!");
      } else {
        playSynthSound("error");
        writeConsoleLog("🌪️ STORM DAMAGE: Space dust storm swept away all unsold ores, fuel/gas/tylium, and destroyed your heavy machinery!");
        
        // Wipe ores
        let oresToWipe;
        if (gameState.currentAsteroid === "TITAN") {
          oresToWipe = ["iron", "copper", "gold", "amethyst", "ruby"];
        } else if (gameState.currentAsteroid === "VOID") {
          oresToWipe = ["bauxite", "magnetite", "pyrite", "diamond", "emerald"];
        } else {
          oresToWipe = ["cryptonite", "elerium", "obsidian", "warpStone", "infinityGem"];
        }
        oresToWipe.forEach(id => {
          gameState.inventory[id] = 0;
        });
        
        // Wipe fuel/gas/tylium and destroy machinery
        if (gameState.currentAsteroid === "TITAN") {
          gameState.fuelLitres = 0;
          gameState.hasExcavator = false;
          gameState.excavatorRunning = false;
          gameState.excavatorProgress = 0;
        } else if (gameState.currentAsteroid === "VOID") {
          gameState.gasUnits = 0;
          gameState.hasDrillMachine = false;
          gameState.excavatorRunning = false;
        } else {
          gameState.tyliumUnits = 0;
          gameState.hasDrillBore = false;
          gameState.excavatorRunning = false;
        }
      }
      
      // Clean up variables & Consume Storm Protection gear
      gameState.hasStormProtection = false;
      gameState.stormShieldActive = false;
      gameState.stormActive = false;
      gameState.isPaused = false;
      
      recalculateExcelModel();
      updateUI();
    }, 4500);
  }

  function triggerSeismicEarthquake() {
    gameState.earthquakeActive = true;
    gameState.isPaused = true; // Pause clock and manual digging/buying
    
    // Show Earthquake Overlay
    const overlay = document.getElementById("earthquake-overlay");
    if (overlay) {
      overlay.classList.remove("hidden");
    }
    
    // Play rumbles and shake screen
    playHeavyExplosionSound();
    document.body.classList.add("explosion-shake");
    if (DOM.asteroidWrapper) DOM.asteroidWrapper.classList.add("explosion-shake");
    
    writeConsoleLog("⚠️ SEISMIC UPDATE: A massive earthquake is shaking Asteroid NOVA!");
    
    // Duration is 4.5 seconds
    setTimeout(() => {
      // Clear overlay and screen shake
      document.body.classList.remove("explosion-shake");
      if (DOM.asteroidWrapper) DOM.asteroidWrapper.classList.remove("explosion-shake");
      if (overlay) overlay.classList.add("hidden");
      
      playSynthSound("error");
      writeConsoleLog("⚠️ SEISMIC HAZARD: Tremors destroyed all unsold ores in your inventory!");
      
      // Wipe Nova ores
      const oresToWipe = ["cryptonite", "elerium", "obsidian", "warpStone", "infinityGem"];
      oresToWipe.forEach(id => {
        gameState.inventory[id] = 0;
      });
      
      // Clean up variables
      gameState.earthquakeActive = false;
      gameState.isPaused = false;
      
      recalculateExcelModel();
      updateUI();
    }, 4500);
  }

  function playExplosionSound() {
    if (!soundEnabled) return;
    try {
      initAudio();
      const now = audioCtx.currentTime;

      const osc = audioCtx.createOscillator();
      const gainNode = audioCtx.createGain();
      
      osc.type = "sawtooth";
      osc.frequency.setValueAtTime(120, now);
      osc.frequency.linearRampToValueAtTime(10, now + 0.8);
      
      const filter = audioCtx.createBiquadFilter();
      filter.type = "lowpass";
      filter.frequency.setValueAtTime(200, now);
      filter.frequency.linearRampToValueAtTime(50, now + 0.8);

      osc.connect(filter);
      filter.connect(gainNode);
      gainNode.connect(audioCtx.destination);

      gainNode.gain.setValueAtTime(0.3, now);
      gainNode.gain.exponentialRampToValueAtTime(0.001, now + 0.8);

      osc.start(now);
      osc.stop(now + 0.8);
    } catch (e) {
      console.warn("Explosion audio failed:", e);
    }
  }

  // ==========================================================================
  // 7. USER INTERFACE UPDATING
  // ==========================================================================

  // Shop Config & Update functions
  const shopConfig = {
    TITAN: {
      excavator: {
        title: "🤖 Auto Excavator",
        desc: "Digs automatically (1s rate). Requires fuel.",
        costText: () => `Buy: $${Math.round(excelModel.excavatorCost)}`,
        canBuy: () => !gameState.hasExcavator && gameState.walletCash >= excelModel.excavatorCost,
        isOwned: () => gameState.hasExcavator,
        buttonClass: "btn btn-sm btn-buy-cash"
      },
      fuel: {
        title: "⛽ Fuel (1 Litre)",
        desc: "Powers the excavator for 10 seconds.",
        costText: () => `Buy: $${Math.round(excelModel.fuelCost)}`,
        canBuy: () => gameState.walletCash >= excelModel.fuelCost,
        buttonClass: "btn btn-sm btn-buy-cash"
      },
      shovel: {
        title: "⛏️ Shiny Titanium Shovel",
        desc: "Doubles manual digging speed. Needs no fuel.",
        costText: () => `Equip: ${excelModel.titanGear.shovel.costCoins} Coins`,
        canBuy: () => !gameState.hasTitaniumShovel && gameState.walletCoins >= excelModel.titanGear.shovel.costCoins,
        isOwned: () => gameState.hasTitaniumShovel,
        buttonClass: "btn btn-sm btn-buy-coins"
      },
      loosener: {
        title: "🌀 Dirt Loosener Attachment",
        desc: "Speeds up all digging (shovel & excavator) by 50%.",
        costText: () => `Equip: ${excelModel.titanGear.loosener.costCoins} Coins`,
        canBuy: () => !gameState.hasDirtLoosener && gameState.walletCoins >= excelModel.titanGear.loosener.costCoins,
        isOwned: () => gameState.hasDirtLoosener,
        buttonClass: "btn btn-sm btn-buy-coins"
      },
      detector: {
        title: "🔍 Metal Detector",
        desc: "Triples probability of Gold, Amethyst, and Ruby drops.",
        costText: () => `Equip: ${excelModel.titanGear.detector.costCoins} Coins`,
        canBuy: () => !gameState.hasMetalDetector && gameState.walletCoins >= excelModel.titanGear.detector.costCoins,
        isOwned: () => gameState.hasMetalDetector,
        buttonClass: "btn btn-sm btn-buy-coins"
      },
      time: {
        title: "⏳ Quantum Time Extender",
        desc: "Injects an extra 60 seconds into the game clock.",
        costText: () => `Extend: ${excelModel.titanGear.time.costCoins} Coin`,
        canBuy: () => gameState.walletCoins >= excelModel.titanGear.time.costCoins,
        buttonClass: "btn btn-sm btn-buy-coins"
      }
    },
    VOID: {
      excavator: {
        title: "💎 Diamond Drilling Machine",
        desc: "Runs automatically (1s rate). Requires Gas.",
        costText: () => `Buy: $${Math.round(excelModel.drillCost)}`,
        canBuy: () => !gameState.hasDrillMachine && gameState.walletCash >= excelModel.drillCost,
        isOwned: () => gameState.hasDrillMachine,
        buttonClass: "btn btn-sm btn-buy-cash"
      },
      fuel: {
        title: "🔋 Gas (1 Unit)",
        desc: "Powers the drilling machine for 10 seconds.",
        costText: () => `Buy: $${Math.round(excelModel.gasCost)}`,
        canBuy: () => gameState.walletCash >= excelModel.gasCost,
        buttonClass: "btn btn-sm btn-buy-cash"
      },
      shovel: {
        title: "⛏️ Shiny Pickaxe",
        desc: "Reduces manual digging duration to 750ms.",
        costText: () => `Equip: ${excelModel.voidGear.pickaxe.costPearls} Pearls`,
        canBuy: () => !gameState.hasShinyPickaxe && gameState.walletPearls >= excelModel.voidGear.pickaxe.costPearls,
        isOwned: () => gameState.hasShinyPickaxe,
        buttonClass: "btn btn-sm btn-buy-coins"
      },
      loosener: {
        title: "💥 Dynamite & Ignition",
        desc: "Triggers an explosion rolling 100 digs instantly.",
        costText: () => `Buy: ${excelModel.voidGear.dynamite.costPearls} Pearls`,
        canBuy: () => gameState.walletPearls >= excelModel.voidGear.dynamite.costPearls,
        isOwned: () => false,
        buttonClass: "btn btn-sm btn-buy-coins"
      },
      detector: {
        title: "Unavailable",
        desc: "Not usable in the VOID.",
        costText: () => "N/A",
        canBuy: () => false,
        isOwned: () => false,
        buttonClass: "btn btn-sm btn-locked"
      },
      time: {
        title: "⏳ Quantum Time Extender",
        desc: "Injects an extra 60 seconds into the game clock.",
        costText: () => `Extend: ${excelModel.voidGear.time.costPearls} Pearls`,
        canBuy: () => gameState.walletPearls >= excelModel.voidGear.time.costPearls,
        buttonClass: "btn btn-sm btn-buy-coins"
      }
    },
    NOVA: {
      excavator: {
        title: "⚛️ Quantum Drill-Bore",
        desc: "Runs automatically. Requires Tylium.",
        costText: () => `Buy: $${Math.round(excelModel.drillBoreCost).toLocaleString()}`,
        canBuy: () => !gameState.hasDrillBore && gameState.walletCash >= excelModel.drillBoreCost,
        isOwned: () => gameState.hasDrillBore,
        buttonClass: "btn btn-sm btn-buy-cash"
      },
      fuel: {
        title: "⚛️ Tylium Fuel (1 Unit)",
        desc: "Powers the Quantum Drill-Bore.",
        costText: () => `Buy: $${Math.round(excelModel.tyliumCost).toLocaleString()}`,
        canBuy: () => gameState.walletCash >= excelModel.tyliumCost,
        buttonClass: "btn btn-sm btn-buy-cash"
      },
      shovel: {
        title: "🔨 Sledgehammer",
        desc: "Reduces manual digging speed to 500ms.",
        costText: () => `Equip: ${excelModel.novaGear.sledgehammer.costCrystals} Crystals`,
        canBuy: () => !gameState.hasSledgehammer && gameState.walletCrystals >= excelModel.novaGear.sledgehammer.costCrystals,
        isOwned: () => gameState.hasSledgehammer,
        buttonClass: "btn btn-sm btn-buy-coins"
      },
      loosener: {
        title: "☢️ Nuclear Explosion",
        desc: "Clears 1000 digs instantly. Max 2 uses per game.",
        costText: () => gameState.nukePurchasedCount >= 2 ? "SOLD OUT" : `Buy: ${excelModel.novaGear.nuke.costCrystals} Crystals`,
        canBuy: () => gameState.nukePurchasedCount < 2 && gameState.walletCrystals >= excelModel.novaGear.nuke.costCrystals,
        isOwned: () => false,
        buttonClass: "btn btn-sm btn-buy-coins"
      },
      detector: {
        title: "🌀 Resonance Stabilizer",
        desc: "Doubles the rare drop rates on Nova.",
        costText: () => `Equip: ${excelModel.novaGear.stabilizer.costCrystals} Crystals`,
        canBuy: () => !gameState.hasResonanceStabilizer && gameState.walletCrystals >= excelModel.novaGear.stabilizer.costCrystals,
        isOwned: () => gameState.hasResonanceStabilizer,
        buttonClass: "btn btn-sm btn-buy-coins"
      },
      time: {
        title: "🧪 Cosmic Water",
        desc: "Speeds up manual digging by 50% (combines with Sledgehammer to 250ms).",
        costText: () => `Equip: ${excelModel.novaGear.water.costCrystals} Crystals`,
        canBuy: () => !gameState.hasCosmicWater && gameState.walletCrystals >= excelModel.novaGear.water.costCrystals,
        isOwned: () => gameState.hasCosmicWater,
        buttonClass: "btn btn-sm btn-buy-coins"
      }
    }
  };

  function updateShopDisplay() {
    const isTitan = (gameState.currentAsteroid === "TITAN");
    const isVoid = (gameState.currentAsteroid === "VOID");
    const isNova = (gameState.currentAsteroid === "NOVA");
    const config = isTitan ? shopConfig.TITAN : (isVoid ? shopConfig.VOID : shopConfig.NOVA);

    // 1. Excavator / Drilling Machine slot
    const itemExc = document.getElementById("shop-item-excavator");
    if (itemExc) {
      itemExc.querySelector(".shop-item-title").textContent = config.excavator.title;
      itemExc.querySelector(".shop-item-desc").textContent = config.excavator.desc;
      const btn = DOM.btnBuyExcavator;
      if (config.excavator.isOwned()) {
        btn.textContent = "Owned";
        btn.disabled = true;
        btn.className = "btn btn-sm btn-buy-cash btn-owned";
      } else {
        btn.textContent = config.excavator.costText();
        btn.disabled = !config.excavator.canBuy();
        btn.className = config.excavator.buttonClass;
      }
    }

    // 2. Fuel / Gas slot
    const itemFuel = document.getElementById("shop-item-fuel");
    if (itemFuel) {
      itemFuel.querySelector(".shop-item-title").textContent = config.fuel.title;
      itemFuel.querySelector(".shop-item-desc").textContent = config.fuel.desc;
      const btn = DOM.btnBuyFuel;
      btn.textContent = config.fuel.costText();
      btn.disabled = !config.fuel.canBuy();
      btn.className = config.fuel.buttonClass;
    }

    // 3. Shovel / Pickaxe slot
    const itemShovel = document.getElementById("shop-item-shovel");
    if (itemShovel) {
      itemShovel.querySelector(".shop-item-title").textContent = config.shovel.title;
      itemShovel.querySelector(".shop-item-desc").textContent = config.shovel.desc;
      const btn = DOM.btnBuyShovel;
      if (config.shovel.isOwned()) {
        btn.textContent = "Equipped";
        btn.disabled = true;
        btn.className = "btn btn-sm btn-owned";
      } else {
        btn.textContent = config.shovel.costText();
        btn.disabled = !config.shovel.canBuy();
        btn.className = config.shovel.buttonClass;
      }
    }

    // 4. Dirt Loosener / Dynamite slot
    const itemLoosener = document.getElementById("shop-item-loosener");
    if (itemLoosener) {
      itemLoosener.querySelector(".shop-item-title").textContent = config.loosener.title;
      itemLoosener.querySelector(".shop-item-desc").textContent = config.loosener.desc;
      const btn = DOM.btnBuyLoosener;
      if (config.loosener.isOwned()) {
        btn.textContent = "Equipped";
        btn.disabled = true;
        btn.className = "btn btn-sm btn-owned";
      } else {
        btn.textContent = config.loosener.costText();
        btn.disabled = !config.loosener.canBuy();
        btn.className = config.loosener.buttonClass;
      }
    }

    // 5. Metal Detector / Resonance Stabilizer slot
    const itemDetector = document.getElementById("shop-item-detector");
    if (itemDetector) {
      if (isTitan || isNova) {
        itemDetector.style.display = "";
        itemDetector.querySelector(".shop-item-title").textContent = config.detector.title;
        itemDetector.querySelector(".shop-item-desc").textContent = config.detector.desc;
        const btn = DOM.btnBuyDetector;
        if (config.detector.isOwned()) {
          btn.textContent = "Equipped";
          btn.disabled = true;
          btn.className = "btn btn-sm btn-owned";
        } else {
          btn.textContent = config.detector.costText();
          btn.disabled = !config.detector.canBuy();
          btn.className = config.detector.buttonClass;
        }
      } else {
        itemDetector.style.display = "none";
      }
    }

    // 6. Time Extender slot
    const itemTime = document.getElementById("shop-item-time");
    if (itemTime) {
      itemTime.querySelector(".shop-item-title").textContent = config.time.title;
      itemTime.querySelector(".shop-item-desc").textContent = config.time.desc;
      const btn = DOM.btnBuyTime;
      btn.textContent = config.time.costText();
      btn.disabled = !config.time.canBuy();
      btn.className = config.time.buttonClass;
    }

    // 7. Storm Protection slot
    const itemProtection = document.getElementById("shop-item-protection");
    if (itemProtection) {
      const btn = DOM.btnBuyStormProtection;
      if (gameState.hasStormProtection) {
        btn.textContent = "Owned";
        btn.disabled = true;
        btn.className = "btn btn-sm btn-buy-cash btn-owned";
      } else {
        let cost = 25000;
        if (isNova) cost = excelModel.novaStormProtectionCost;
        else if (isTitan) cost = 10000;
        else if (isVoid) cost = excelModel.stormProtectionCost;
        btn.textContent = `Buy: $${cost.toLocaleString()}`;
        btn.disabled = gameState.walletCash < cost;
        btn.className = "btn btn-sm btn-buy-cash";
      }
    }
  }
  
  function updateUI() {
    const isTitan = (gameState.currentAsteroid === "TITAN");
    const isVoid = (gameState.currentAsteroid === "VOID");
    const isNova = (gameState.currentAsteroid === "NOVA");

    // 1. Timer
    const mins = Math.floor(gameState.timer / 60);
    const secs = gameState.timer % 60;
    DOM.gameTimer.textContent = `${mins.toString().padStart(2, "0")}:${secs.toString().padStart(2, "0")}`;
    
    // 2. Wallet Cash
    DOM.walletCash.textContent = `$${Math.round(gameState.walletCash).toLocaleString()}`;
    
    // 3. Wallet Coins / Pearls / Crystals
    if (isTitan) {
      DOM.coinPillContainer.classList.remove("hidden");
      DOM.pearlPillContainer.classList.add("hidden");
      if (DOM.crystalPillContainer) DOM.crystalPillContainer.classList.add("hidden");
      DOM.walletCoins.textContent = `${gameState.walletCoins} Coin${gameState.walletCoins === 1 ? "" : "s"}`;
    } else if (isVoid) {
      DOM.coinPillContainer.classList.add("hidden");
      DOM.pearlPillContainer.classList.remove("hidden");
      if (DOM.crystalPillContainer) DOM.crystalPillContainer.classList.add("hidden");
      DOM.walletPearls.textContent = `${gameState.walletPearls} Pearl${gameState.walletPearls === 1 ? "" : "s"}`;
    } else if (isNova) {
      DOM.coinPillContainer.classList.add("hidden");
      DOM.pearlPillContainer.classList.add("hidden");
      if (DOM.crystalPillContainer) {
        DOM.crystalPillContainer.classList.remove("hidden");
        DOM.walletCrystals.textContent = `${gameState.walletCrystals} Crystal${gameState.walletCrystals === 1 ? "" : "s"}`;
      }
    }
    
    // 4. Net Worth
    const nw = calculateCurrentNetWorth();
    DOM.gameNetWorth.textContent = `$${Math.round(nw).toLocaleString()}`;

    // 5. Top fuel / gas / tylium display
    if (DOM.fuelTop) {
      if (isTitan) {
        DOM.fuelTop.textContent = `${gameState.fuelLitres.toFixed(1)} L`;
      } else if (isVoid) {
        DOM.fuelTop.textContent = `${gameState.gasUnits.toFixed(1)} U`;
      } else {
        DOM.fuelTop.textContent = `${gameState.tyliumUnits.toFixed(1)} U`;
      }
      
      const labelEl = DOM.fuelTop.parentElement.querySelector(".status-lbl");
      if (labelEl) {
        if (isTitan) labelEl.textContent = "⛽ EXCAVATOR FUEL";
        else if (isVoid) labelEl.textContent = "🔋 DRILL GAS";
        else labelEl.textContent = "⚛️ TYLIUM FUEL";
      }
    }

    // 6. Shovel/Pickaxe Telemetry
    if (isTitan) {
      DOM.telemetryShovel.textContent = gameState.hasTitaniumShovel ? "Titanium Shovel" : "Rusty Shovel";
      DOM.telemetryLoosener.textContent = gameState.hasDirtLoosener ? "Loosener (+50% Speed)" : "None (1x Speed)";
      DOM.telemetryDetector.textContent = gameState.hasMetalDetector ? "Equipped (3x Rares)" : "Not Equipped";
    } else if (isVoid) {
      DOM.telemetryShovel.textContent = gameState.hasShinyPickaxe ? "Shiny Pickaxe (VOID)" : "Shiny Shovel";
      DOM.telemetryLoosener.textContent = "N/A (Left on Titan)";
      DOM.telemetryDetector.textContent = "N/A (Left on Titan)";
    } else if (isNova) {
      DOM.telemetryShovel.textContent = gameState.hasSledgehammer ? "Sledgehammer" : "Ice Axe";
      DOM.telemetryLoosener.textContent = gameState.hasCosmicWater ? "Cosmic Water (+50% Speed)" : "None";
      DOM.telemetryDetector.textContent = gameState.hasResonanceStabilizer ? "Resonance Stabilizer (2x Rares)" : "None";
    }
    DOM.telemetryDigs.textContent = gameState.digsCompleted;

    // 7. Heavy Machinery Run Button & Subtext
    let activeMachine = false;
    let activeFuelName = "";
    let activeFuelValue = 0;
    let activeFuelUnit = "";
    let machineName = "";
    
    if (isTitan) {
      activeMachine = gameState.hasExcavator;
      activeFuelName = "Fuel";
      activeFuelValue = gameState.fuelLitres;
      activeFuelUnit = "L";
      machineName = "Excavator";
    } else if (isVoid) {
      activeMachine = gameState.hasDrillMachine;
      activeFuelName = "Gas";
      activeFuelValue = gameState.gasUnits;
      activeFuelUnit = "Units";
      machineName = "Drill Machine";
    } else if (isNova) {
      activeMachine = gameState.hasDrillBore;
      activeFuelName = "Tylium";
      activeFuelValue = gameState.tyliumUnits;
      activeFuelUnit = "U";
      machineName = "Drill-Bore";
    }
    
    if (activeMachine) {
      DOM.btnToggleExcavator.disabled = false;
      DOM.btnToggleExcavator.classList.remove("btn-locked");
      if (!gameState.excavatorRunning) {
        DOM.btnToggleExcavator.textContent = `🤖 Run ${machineName}`;
      }
      DOM.fuelDisplaySubtext.textContent = `${activeFuelName} level: ${activeFuelValue.toFixed(1)}${activeFuelUnit} (${Math.round(activeFuelValue * 10)}s left)`;
    } else {
      DOM.btnToggleExcavator.disabled = true;
      DOM.btnToggleExcavator.classList.add("btn-locked");
      DOM.btnToggleExcavator.textContent = `🤖 Run ${machineName} (Locked)`;
      DOM.fuelDisplaySubtext.textContent = `Requires ${machineName} & ${activeFuelName}`;
    }

    // 8. Dynamic Ore List update (Names, Icons, Counts, Prices, Changes, Sell Buttons)
    const oreRows = ["iron", "copper", "gold", "amethyst", "ruby"];
    oreRows.forEach((id, index) => {
      const ore = excelModel.ores[index];
      const rowEl = document.getElementById(`ore-item-${id}`);
      if (rowEl) {
        const iconEl = rowEl.querySelector(".ore-icon");
        if (iconEl) iconEl.textContent = ore.icon;

        const nameEl = rowEl.querySelector(".ore-name");
        if (nameEl) nameEl.textContent = ore.name;

        const countValEl = rowEl.querySelector(".ore-count-val");
        if (countValEl) countValEl.textContent = `${gameState.inventory[ore.id]} held`;

        const priceEl = rowEl.querySelector(".ore-price");
        if (priceEl) priceEl.textContent = `$${Math.round(ore.currentPrice)}`;

        const changeEl = rowEl.querySelector(".ore-change");
        if (changeEl) {
          if (ore.appreciationRate > 0) {
            const flatIncrease = ore.baseValue * ore.appreciationRate;
            const ratePct = Math.round(ore.appreciationRate * 100);
            changeEl.textContent = `+$${Math.round(flatIncrease)} (+${ratePct}%)`;
            changeEl.className = "ore-change positive";
          } else {
            changeEl.textContent = `+$0 (0%)`;
            changeEl.className = "ore-change flat";
          }
        }

        const sellBtn = rowEl.querySelector(".btn-sell");
        if (sellBtn) {
          sellBtn.disabled = gameState.inventory[ore.id] === 0;
        }
      }
    });

    // 9. Appreciation Countdown
    DOM.appreciationCountdown.textContent = `${gameState.appreciationTimer}s`;

    // 10. Update Shop Display
    updateShopDisplay();

    // 11. Convergence Table labels, counts & rates
    const convergenceRows = [
      { class: "iron", defaultLabel: "⬜ Iron Ore", voidLabel: "⬜ Bauxite", novaLabel: "🟢 Cryptonite" },
      { class: "copper", defaultLabel: "🟫 Copper Ore", voidLabel: "🟫 Magnetite", novaLabel: "🔵 Elerium" },
      { class: "gold", defaultLabel: "🟨 Gold Ore", voidLabel: "🟨 Pyrite", novaLabel: "⚫ Obsidian" },
      { class: "amethyst", defaultLabel: "🟪 Amethyst Ore", voidLabel: "💎 Diamond", novaLabel: "🌀 Warp-Stone" },
      { class: "ruby", defaultLabel: "🟥 Ruby Ore", voidLabel: "🟩 Emerald", novaLabel: "💎 Infinity Gem" },
      { class: "coins", defaultLabel: "🪙 Extra Coins", voidLabel: "🦪 Extra Pearls", novaLabel: "🔮 Extra Crystals" }
    ];

    convergenceRows.forEach(row => {
      const rowEl = document.querySelector(`.data-table .res-row.${row.class}`);
      if (rowEl) {
        const labelCell = rowEl.querySelector("td:first-child");
        if (labelCell) {
          if (isTitan) labelCell.textContent = row.defaultLabel;
          else if (isVoid) labelCell.textContent = row.voidLabel;
          else if (isNova) labelCell.textContent = row.novaLabel;
        }
      }
    });

    const totalMined = isTitan ? 
      (gameState.tallies.iron + gameState.tallies.copper + gameState.tallies.gold + gameState.tallies.amethyst + gameState.tallies.ruby) :
      (isVoid ? (gameState.tallies.bauxite + gameState.tallies.magnetite + gameState.tallies.pyrite + gameState.tallies.diamond + gameState.tallies.emerald) :
                (gameState.tallies.cryptonite + gameState.tallies.elerium + gameState.tallies.obsidian + gameState.tallies.warpStone + gameState.tallies.infinityGem));

    excelModel.ores.forEach((ore, index) => {
      const count = gameState.tallies[ore.id];
      const domId = ["iron", "copper", "gold", "amethyst", "ruby"][index];
      const countCell = document.getElementById(`act-count-${domId}`);
      const rateCell = document.getElementById(`act-rate-${domId}`);
      
      if (countCell) countCell.textContent = count;
      if (rateCell) rateCell.textContent = totalMined > 0 ? `${((count / totalMined) * 100).toFixed(1)}%` : "0.0%";
    });

    const totalDigs = gameState.digsCompleted;
    const secondaryCount = isTitan ? gameState.tallies.coins : (isVoid ? gameState.tallies.pearls : gameState.tallies.crystals);
    DOM.actCountCoins.textContent = secondaryCount;
    DOM.actRateCoins.textContent = totalDigs > 0 ? `${((secondaryCount / totalDigs) * 100).toFixed(1)}%` : "0.0%";

    // 12. Teleport button panel status
    if (isTitan) {
      DOM.btnTeleportVoid.parentElement.classList.remove("hidden");
      DOM.btnTeleportVoid.textContent = "🌌 Teleport to VOID";
      const currentNetWorth = calculateCurrentNetWorth();
      const canTeleport = currentNetWorth >= excelModel.travelCost;
      DOM.btnTeleportVoid.disabled = !canTeleport;
      
      if (canTeleport) {
        DOM.btnTeleportVoid.classList.remove("disabled");
        DOM.btnTeleportVoid.classList.add("btn-teleport-active");
        DOM.teleportDescText.textContent = "Ready to plant flag and teleport!";
      } else {
        DOM.btnTeleportVoid.classList.add("disabled");
        DOM.btnTeleportVoid.classList.remove("btn-teleport-active");
        DOM.teleportDescText.textContent = `Requires $${Math.round(excelModel.travelCost).toLocaleString()} net worth to travel (Current: $${Math.round(currentNetWorth).toLocaleString()}).`;
      }
    } else if (isVoid) {
      DOM.btnTeleportVoid.parentElement.classList.remove("hidden");
      DOM.btnTeleportVoid.textContent = "🌌 Teleport out of Void";
      DOM.btnTeleportVoid.disabled = false;
      DOM.btnTeleportVoid.classList.remove("disabled");
      DOM.btnTeleportVoid.classList.add("btn-teleport-active");
      DOM.teleportDescText.textContent = "Choose your destination: Earth ($200,000) or NOVA ($500,000).";
    } else if (isNova) {
      DOM.btnTeleportVoid.parentElement.classList.remove("hidden");
      DOM.btnTeleportVoid.textContent = "🌎 Return to Earth";
      const canTeleport = gameState.walletCash >= 50000000;
      DOM.btnTeleportVoid.disabled = !canTeleport;
      if (canTeleport) {
        DOM.btnTeleportVoid.classList.remove("disabled");
        DOM.btnTeleportVoid.classList.add("btn-teleport-active");
        DOM.teleportDescText.textContent = "Ready to return to Earth!";
      } else {
        DOM.btnTeleportVoid.classList.add("disabled");
        DOM.btnTeleportVoid.classList.remove("btn-teleport-active");
        DOM.teleportDescText.textContent = `Requires $50,000,000 cash to return to Earth (Current: $${Math.round(gameState.walletCash).toLocaleString()}).`;
      }
    } else {
      DOM.btnTeleportVoid.parentElement.classList.add("hidden");
    }
  }

  // Streaming Telemetry logging Console
  function writeConsoleLog(msg) {
    const p = document.createElement("p");
    p.textContent = `> ${msg}`;
    DOM.logConsole.appendChild(p);
    
    // Scroll to bottom
    DOM.logConsole.scrollTop = DOM.logConsole.scrollHeight;
    
    // Cap lines at 30 to prevent DOM leakage
    while (DOM.logConsole.children.length > 30) {
      DOM.logConsole.removeChild(DOM.logConsole.firstChild);
    }
  }

  DOM.btnClearConsole.addEventListener("click", () => {
    DOM.logConsole.innerHTML = "";
    writeConsoleLog("Telemetry terminal cleared.");
    playSynthSound("click");
  });

  // ==========================================================================
  // 8. CANVAS WEALTH PLOTTING CHART
  // ==========================================================================
  function drawNetWorthChart() {
    if (!canvas) return;
    
    const width = canvas.width;
    const height = canvas.height;
    
    // Clear canvas
    ctx.clearRect(0, 0, width, height);
    
    // 1. Grid Background lines
    ctx.strokeStyle = "rgba(255, 255, 255, 0.05)";
    ctx.lineWidth = 1;
    for (let x = 40; x < width; x += 50) {
      ctx.beginPath();
      ctx.moveTo(x, 0);
      ctx.lineTo(x, height - 25);
      ctx.stroke();
    }
    for (let y = 10; y < height - 25; y += 35) {
      ctx.beginPath();
      ctx.moveTo(40, y);
      ctx.lineTo(width, y);
      ctx.stroke();
    }
    
    // Draw Y axis border
    ctx.strokeStyle = "rgba(255, 255, 255, 0.15)";
    ctx.beginPath();
    ctx.moveTo(40, 5);
    ctx.lineTo(40, height - 25);
    ctx.lineTo(width - 5, height - 25);
    ctx.stroke();
    
    const count = gameState.history.netWorths.length;
    if (count < 2) {
      // Print loading message
      ctx.fillStyle = "var(--color-text-secondary)";
      ctx.font = "11px Arial";
      ctx.fillText("Gathering telemetry data...", width / 2 - 60, height / 2);
      return;
    }
    
    // Scale parameters
    const maxNW = Math.max(...gameState.history.netWorths, 500); // minimum scale limit
    const minNW = 0;
    
    // Draw Y labels (Min, Mid, Max)
    ctx.fillStyle = "var(--color-text-secondary)";
    ctx.font = "9px var(--font-display)";
    ctx.textAlign = "right";
    ctx.fillText(`$${Math.round(maxNW)}`, 35, 12);
    ctx.fillText(`$${Math.round(maxNW / 2)}`, 35, (height - 30) / 2 + 5);
    ctx.fillText("$0", 35, height - 25);
    
    // Plot Line
    ctx.strokeStyle = "var(--accent-cyan)";
    ctx.lineWidth = 3;
    ctx.beginPath();
    
    const plotWidth = width - 50;
    const plotHeight = height - 35;
    
    for (let i = 0; i < count; i++) {
      const px = 40 + (i / (count - 1)) * plotWidth;
      const py = (height - 25) - ((gameState.history.netWorths[i] - minNW) / (maxNW - minNW)) * plotHeight;
      
      if (i === 0) {
        ctx.moveTo(px, py);
      } else {
        ctx.lineTo(px, py);
      }
    }
    ctx.stroke();
    
    // Gradient fill under the line
    const grad = ctx.createLinearGradient(0, 0, 0, height - 25);
    grad.addColorStop(0, "rgba(0, 229, 255, 0.25)");
    grad.addColorStop(1, "rgba(0, 229, 255, 0)");
    
    ctx.fillStyle = grad;
    ctx.beginPath();
    for (let i = 0; i < count; i++) {
      const px = 40 + (i / (count - 1)) * plotWidth;
      const py = (height - 25) - ((gameState.history.netWorths[i] - minNW) / (maxNW - minNW)) * plotHeight;
      if (i === 0) {
        ctx.moveTo(px, height - 25);
        ctx.lineTo(px, py);
      } else {
        ctx.lineTo(px, py);
      }
    }
    ctx.lineTo(40 + plotWidth, height - 25);
    ctx.closePath();
    ctx.fill();
    
    // Plot marker at latest point
    const latestIdx = count - 1;
    const lx = 40 + plotWidth;
    const ly = (height - 25) - ((gameState.history.netWorths[latestIdx] - minNW) / (maxNW - minNW)) * plotHeight;
    
    ctx.fillStyle = "var(--accent-gold)";
    ctx.beginPath();
    ctx.arc(lx, ly, 5, 0, Math.PI * 2);
    ctx.fill();
    ctx.strokeStyle = "#fff";
    ctx.lineWidth = 1;
    ctx.stroke();
  }

  // ==========================================================================
  // 9. APP START
  // ==========================================================================
  
  // ==========================================================================
  // 10. LEADERBOARD & PAUSE SYSTEM logic
  // ==========================================================================

  // Save leaderboard entry to localStorage
  function saveLeaderboardEntry(name, netWorth, cash, coins, digs) {
    let leaderboard = [];
    try {
      leaderboard = JSON.parse(safeStorage.getItem("asteroid_rush_leaderboard") || "[]");
    } catch (e) {
      leaderboard = [];
    }
    
    const isTitan = (gameState.currentAsteroid === "TITAN");
    const isVoid = (gameState.currentAsteroid === "VOID");
    const isNova = (gameState.currentAsteroid === "NOVA");
    
    const newEntry = {
      name: name || "Miner",
      netWorth: Math.round(netWorth),
      cash: Math.round(cash),
      asteroid: gameState.currentAsteroid,
      coins: isTitan ? coins : 0,
      pearls: isVoid ? coins : 0,
      crystals: isNova ? coins : 0,
      digs: digs,
      date: new Date().toLocaleString()
    };
    
    leaderboard.push(newEntry);
    leaderboard.sort((a, b) => b.netWorth - a.netWorth);
    leaderboard = leaderboard.slice(0, 10);
    
    safeStorage.setItem("asteroid_rush_leaderboard", JSON.stringify(leaderboard));
    updateLeaderboardUIs();
  }

  // Update Leaderboard tables in tab and modal
  function updateLeaderboardUIs() {
    let leaderboard = [];
    try {
      leaderboard = JSON.parse(safeStorage.getItem("asteroid_rush_leaderboard") || "[]");
    } catch (e) {
      leaderboard = [];
    }

    // A. Main Leaderboard Tab View
    if (DOM.leaderboardTableBody) {
      if (leaderboard.length === 0) {
        DOM.leaderboardTableBody.innerHTML = `
          <tr>
            <td colspan="7" style="text-align: center; color: var(--color-text-secondary); padding: 30px;">
              No expedition telemetry recorded yet. Start mining to establish records!
            </td>
          </tr>
        `;
      } else {
        DOM.leaderboardTableBody.innerHTML = leaderboard.map((entry, index) => {
          let rankIcon = index === 0 ? "🥇 " : (index === 1 ? "🥈 " : (index === 2 ? "🥉 " : ""));
          let locationText = entry.asteroid || "TITAN";
          let dropsText = "";
          if (locationText === "TITAN") {
            dropsText = `🪙 ${entry.coins || 0}`;
          } else if (locationText === "VOID") {
            dropsText = `🦪 ${entry.pearls || 0}`;
          } else if (locationText === "NOVA") {
            dropsText = `🔮 ${entry.crystals || 0}`;
          } else {
            dropsText = `🪙 ${entry.coins || 0}`;
          }
          
          return `
            <tr>
              <td style="text-align: center; font-weight: bold; color: ${index < 3 ? 'var(--accent-gold)' : 'inherit'}">${rankIcon}${index + 1}</td>
              <td style="font-weight: 600; color: #fff;">${escapeHtml(entry.name)}</td>
              <td style="text-align: right; font-weight: bold; color: var(--accent-gold);">$${entry.netWorth.toLocaleString()}</td>
              <td style="text-align: right;">$${entry.cash.toLocaleString()}</td>
              <td style="text-align: center; font-size: 0.85rem; font-weight: 500;">
                <span style="color: ${locationText === 'NOVA' ? 'var(--accent-purple)' : (locationText === 'VOID' ? 'var(--accent-cyan)' : 'var(--accent-gold)')}; font-weight: bold;">
                  ${locationText}
                </span>
                <span style="color: var(--color-text-secondary); margin-left: 6px;">(${dropsText})</span>
              </td>
              <td style="text-align: right;">${entry.digs}</td>
              <td style="color: var(--color-text-secondary); font-size: 0.8rem;">${entry.date}</td>
            </tr>
          `;
        }).join("");
      }
    }

    // B. Mini-leaderboard inside End Game Modal
    if (DOM.modalLeaderboardBody) {
      if (leaderboard.length === 0) {
        DOM.modalLeaderboardBody.innerHTML = `
          <tr>
            <td colspan="4" style="text-align: center; color: var(--color-text-secondary); padding: 10px;">
              No records.
            </td>
          </tr>
        `;
      } else {
        DOM.modalLeaderboardBody.innerHTML = leaderboard.slice(0, 5).map((entry, index) => {
          let rankColor = index === 0 ? 'var(--accent-gold)' : (index === 1 ? '#d1d1d1' : (index === 2 ? '#cd7f32' : 'inherit'));
          return `
            <tr style="border-bottom: 1px solid rgba(255, 255, 255, 0.05);">
              <td style="text-align: center; padding: 4px; font-weight: bold; color: ${rankColor};">${index + 1}</td>
              <td style="padding: 4px; font-weight: 600; color: #fff; max-width: 100px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${escapeHtml(entry.name)}</td>
              <td style="padding: 4px; text-align: right; font-weight: bold; color: var(--accent-gold);">$${entry.netWorth.toLocaleString()}</td>
              <td style="padding: 4px; text-align: right; color: var(--color-text-secondary);">${entry.digs}</td>
            </tr>
          `;
        }).join("");
      }
    }
  }

  function escapeHtml(text) {
    return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  // Clear Leaderboard
  if (DOM.btnClearLeaderboard) {
    DOM.btnClearLeaderboard.addEventListener("click", () => {
      if (confirm("Are you sure you want to permanently delete all leaderboard records?")) {
        safeStorage.removeItem("asteroid_rush_leaderboard");
        safeStorage.removeItem("asteroid_rush_highscore");
        updateLeaderboardUIs();
        writeConsoleLog("🏆 Leaderboard records have been reset.");
        playSynthSound("click");
      }
    });
  }

  // Toggle Pause Game
  function togglePauseGame() {
    if (!gameState.isActive) return;
    
    gameState.isPaused = !gameState.isPaused;
    playSynthSound("click");
    
    if (gameState.isPaused) {
      if (DOM.btnPauseGame) {
        DOM.btnPauseGame.textContent = "▶️ Resume";
        DOM.btnPauseGame.classList.add("paused");
      }
      if (DOM.pausedOverlay) DOM.pausedOverlay.classList.remove("hidden");
      writeConsoleLog("⏸️ EXPEDITION PAUSED. Timers, appreciation and automation offline.");
    } else {
      if (DOM.btnPauseGame) {
        DOM.btnPauseGame.textContent = "⏸️ Pause";
        DOM.btnPauseGame.classList.remove("paused");
      }
      if (DOM.pausedOverlay) DOM.pausedOverlay.classList.add("hidden");
      writeConsoleLog("▶️ EXPEDITION RESUMED. Operations are back online.");
    }
  }

  if (DOM.btnPauseGame) {
    DOM.btnPauseGame.addEventListener("click", togglePauseGame);
  }
  if (DOM.btnResumeOverlay) {
    DOM.btnResumeOverlay.addEventListener("click", togglePauseGame);
  }

  // Pilot Name input binding
  if (DOM.pilotNameInput) {
    const savedName = safeStorage.getItem("asteroid_rush_player_name");
    if (savedName) {
      DOM.pilotNameInput.value = savedName;
    }
    DOM.pilotNameInput.addEventListener("input", () => {
      safeStorage.setItem("asteroid_rush_player_name", DOM.pilotNameInput.value.trim());
    });
  }

  // Render leaderboard initially
  updateLeaderboardUIs();

  // Flag and Teleportation Logic
  function saveFlagPlanted(pilotName, flagMessage) {
    let flags = [];
    try {
      flags = JSON.parse(safeStorage.getItem("asteroid_rush_flags") || "[]");
    } catch (e) {
      flags = [];
    }
    
    const newFlag = {
      name: pilotName || "Miner One",
      message: flagMessage,
      date: new Date().toLocaleString()
    };
    
    flags.push(newFlag);
    safeStorage.setItem("asteroid_rush_flags", JSON.stringify(flags));
    updateFlagsTableUI();
  }

  function updateFlagsTableUI() {
    let flags = [];
    try {
      flags = JSON.parse(safeStorage.getItem("asteroid_rush_flags") || "[]");
    } catch (e) {
      flags = [];
    }
    
    if (DOM.flagsTableBody) {
      if (flags.length === 0) {
        DOM.flagsTableBody.innerHTML = `
          <tr>
            <td colspan="3" style="text-align: center; color: var(--color-text-secondary); padding: 30px;">
              No flags planted on TITAN yet. Earn $100,000 to teleport and plant yours!
            </td>
          </tr>
        `;
      } else {
        DOM.flagsTableBody.innerHTML = flags.map(flag => {
          return `
            <tr>
              <td style="font-weight: 600; color: #fff;">${escapeHtml(flag.name)}</td>
              <td style="color: var(--accent-gold); font-style: italic;">🚩 ${escapeHtml(flag.message)}</td>
              <td style="color: var(--color-text-secondary); font-size: 0.8rem;">${flag.date}</td>
            </tr>
          `;
        }).join("");
      }
    }
  }

  // Teleport Button Click Handler
  if (DOM.btnTeleportVoid) {
    DOM.btnTeleportVoid.addEventListener("click", () => {
      if (!gameState.isActive || gameState.isPaused) return;
      
      if (gameState.currentAsteroid === "TITAN") {
        // Pause automation and timer
        gameState.isPaused = true;
        if (DOM.pausedOverlay) DOM.pausedOverlay.classList.add("hidden"); // hide regular pause overlay
        DOM.flagModal.classList.remove("hidden");
        playSynthSound("click");
        
        const pilotName = DOM.pilotNameInput ? DOM.pilotNameInput.value.trim() : "Miner One";
        DOM.flagTextInput.value = `${pilotName}'s Titan Outpost`;
      } else if (gameState.currentAsteroid === "VOID") {
        // Sell all VOID inventory to cash
        sellAllInventory(true);
        
        // Show destination selection modal
        gameState.isPaused = true;
        
        // Enable option buttons in the destination modal based on cash
        const btnWarpEarth = document.getElementById("btn-warp-earth");
        const btnWarpNova = document.getElementById("btn-warp-nova");
        const destCardEarth = document.getElementById("dest-card-earth");
        const destCardNova = document.getElementById("dest-card-nova");
        
        if (btnWarpEarth && btnWarpNova) {
          const cash = gameState.walletCash;
          
          if (cash >= 200000) {
            btnWarpEarth.disabled = false;
            btnWarpEarth.classList.remove("disabled");
            if (destCardEarth) destCardEarth.style.opacity = "1";
          } else {
            btnWarpEarth.disabled = true;
            btnWarpEarth.classList.add("disabled");
            if (destCardEarth) destCardEarth.style.opacity = "0.5";
          }
          
          if (cash >= 500000) {
            btnWarpNova.disabled = false;
            btnWarpNova.classList.remove("disabled");
            if (destCardNova) destCardNova.style.opacity = "1";
          } else {
            btnWarpNova.disabled = true;
            btnWarpNova.classList.add("disabled");
            if (destCardNova) destCardNova.style.opacity = "0.5";
          }
        }
        
        const destModal = document.getElementById("destination-modal");
        if (destModal) {
          destModal.classList.remove("hidden");
        }
        playSynthSound("click");
      } else if (gameState.currentAsteroid === "NOVA") {
        // Teleport back to Earth from Nova
        sellAllInventory(true);
        if (gameState.walletCash >= 50000000) {
          if (confirm("Are you sure you want to return to Earth? This costs $50,000,000 cash and will end the game.")) {
            gameState.walletCash -= 50000000;
            writeConsoleLog("🌎 TELEPORT TO EARTH CONFIRMED: Deducted $50,000,000 cash warp fee.");
            endGameMission(false); // timerExpired = false
          }
        } else {
          writeConsoleLog("❌ WARP ERROR: Insufficient cash to return to Earth ($50,000,000 required)!");
        }
      }
    });
  }

  // Confirm Flag & Teleport Click Handler (TITAN to VOID transition)
  if (DOM.btnConfirmFlag) {
    DOM.btnConfirmFlag.addEventListener("click", () => {
      const flagMessage = DOM.flagTextInput.value.trim() || "Claimed TITAN!";
      const pilotName = DOM.pilotNameInput ? DOM.pilotNameInput.value.trim() : "Miner One";
      
      // Save flag
      saveFlagPlanted(pilotName, flagMessage);
      
      // 1. Hide flag modal & show warp overlay
      DOM.flagModal.classList.add("hidden");
      const warpOverlay = document.getElementById("warp-overlay");
      const desc = warpOverlay ? warpOverlay.querySelector("p") : null;
      const title = warpOverlay ? warpOverlay.querySelector("h1") : null;
      if (warpOverlay) {
        if (title) {
          title.textContent = "🌌 INITIATING HYPERSPACE JUMP...";
          title.style.color = "var(--accent-purple)";
          title.style.textShadow = "0 0 15px rgba(208, 51, 255, 0.8)";
        }
        if (desc) desc.textContent = "Preparing warp engines. Teleporting from TITAN to VOID...";
        warpOverlay.classList.remove("hidden");
      }
      
      // 2. Pause game and start shake
      gameState.isPaused = true;
      document.body.classList.add("explosion-shake");
      
      // 3. Play warp sounds
      playWarpSound();
      
      // 4. Run transition after 4.5 seconds
      setTimeout(() => {
        // Clear screen shake and hide warp overlay
        document.body.classList.remove("explosion-shake");
        if (warpOverlay) warpOverlay.classList.add("hidden");
        
        // Sell all Titan inventory to cash
        sellAllInventory(true);
        
        // Deduct travel cost
        gameState.walletCash -= excelModel.travelCost;
        
        // Reset Titan gear and machinery
        gameState.hasTitaniumShovel = false;
        gameState.hasDirtLoosener = false;
        gameState.hasMetalDetector = false;
        gameState.hasExcavator = false;
        gameState.fuelLitres = 0;
        gameState.excavatorRunning = false;
        gameState.excavatorProgress = 0;
        
        // Reset automation state for VOID
        gameState.hasShinyPickaxe = false;
        gameState.hasDrillMachine = false;
        gameState.gasUnits = 0;
        gameState.dynamitePurchasedCount = 0;
        gameState.hasStormProtection = false;
        gameState.stormShieldActive = false;
        excelModel.voidGear.dynamite.costPearls = 25;
        
        // Reset storm scheduler for the new VOID phase
        gameState.elapsedSeconds = 0;
        gameState.scheduledStormSecond = null;
        gameState.stormWarningStarted = false;
        gameState.stormActive = false;
        if (DOM.stormPanel) {
          DOM.stormPanel.classList.add("hidden");
          DOM.btnLockShield.disabled = true;
          DOM.btnLockShield.textContent = "🔒 ACTIVATE STORM SHIELD";
          DOM.btnLockShield.style.background = "";
        }
        
        // Switch current asteroid to "VOID"
        gameState.currentAsteroid = "VOID";
        excelModel.ores = excelModel.voidOres;
        excelModel.gear = excelModel.voidGear;
        
        // Reset timer to 5 minutes (300s)
        gameState.timer = 300;
        excelModel.initialTimer = 300;
        
        // Reset appreciation timer to 10s for VOID
        gameState.appreciationTimer = 10;
        gameState.appreciationTicks = 0;
        
        // Reset ore current prices for VOID
        excelModel.voidOres.forEach(ore => ore.currentPrice = ore.baseValue);
        
        // Toggle theme
        document.body.classList.add("void-theme");
        
        // Update badge & asteroid labels
        DOM.asteroidName.textContent = "VOID";
        DOM.asteroidName.className = "asteroid-void-text";
        DOM.asteroidDistance.textContent = "(Deep Space)";
        
        // Resume game
        gameState.isPaused = false;
        
        // Update pointers & sync Excel
        syncExcelParametersToGame();
        recalculateExcelModel();
        updateUI();
        
        // Reset history for graph scaling in Void
        gameState.history.timestamps = [0];
        gameState.history.netWorths = [gameState.walletCash];
        drawNetWorthChart();
        
        playSynthSound("upgrade");
        writeConsoleLog("🌌 TELEPORTATION SUCCESSFUL! Welcome to the VOID.");
        writeConsoleLog(`🚩 Pilot flag planted: "${flagMessage}"`);
        writeConsoleLog("⚠️ All Titanium gear and fuel left behind on TITAN.");
        writeConsoleLog("ℹ️ Manual tool upgraded to Shiny Shovel. Mine VOID ores to earn cash!");
      }, 4500);
    });
  }

  // Cancel Destination selector
  const btnCancelDestination = document.getElementById("btn-cancel-destination");
  if (btnCancelDestination) {
    btnCancelDestination.addEventListener("click", () => {
      const destModal = document.getElementById("destination-modal");
      if (destModal) destModal.classList.add("hidden");
      gameState.isPaused = false;
      playSynthSound("click");
      updateUI();
    });
  }

  // Warp option Earth
  const btnWarpEarth = document.getElementById("btn-warp-earth");
  if (btnWarpEarth) {
    btnWarpEarth.addEventListener("click", () => {
      if (gameState.walletCash >= 200000) {
        triggerHyperspaceReturnToEarth();
      }
    });
  }

  // Warp option Nova
  const btnWarpNova = document.getElementById("btn-warp-nova");
  if (btnWarpNova) {
    btnWarpNova.addEventListener("click", () => {
      if (gameState.walletCash >= 500000) {
        triggerHyperspaceWarpToNova();
      }
    });
  }

  function triggerHyperspaceReturnToEarth() {
    const destModal = document.getElementById("destination-modal");
    if (destModal) destModal.classList.add("hidden");
    
    // Deduct $200,000 cash
    gameState.walletCash -= 200000;
    
    // Play warp sounds and show warp overlay
    const warpOverlay = document.getElementById("warp-overlay");
    const desc = warpOverlay ? warpOverlay.querySelector("p") : null;
    const title = warpOverlay ? warpOverlay.querySelector("h1") : null;
    if (warpOverlay) {
      if (title) {
        title.textContent = "🌎 RETURNING TO EARTH...";
        title.style.color = "var(--accent-cyan)";
        title.style.textShadow = "0 0 15px rgba(0, 229, 255, 0.8)";
      }
      if (desc) desc.textContent = "Coordinating home coordinates. Warp engines at 100%...";
      warpOverlay.classList.remove("hidden");
    }
    
    document.body.classList.add("explosion-shake");
    playWarpSound();
    
    setTimeout(() => {
      document.body.classList.remove("explosion-shake");
      if (warpOverlay) warpOverlay.classList.add("hidden");
      
      writeConsoleLog("🌎 MISSION COMPLETE: Safely returned to Earth. Earnings recorded!");
      endGameMission(false); // timerExpired = false
    }, 4500);
  }

  function triggerHyperspaceWarpToNova() {
    const destModal = document.getElementById("destination-modal");
    if (destModal) destModal.classList.add("hidden");
    
    // Deduct $500,000 cash
    gameState.walletCash -= 500000;
    
    // Reset Void gear and machinery
    gameState.hasShinyPickaxe = false;
    gameState.hasDrillMachine = false;
    gameState.gasUnits = 0;
    gameState.dynamitePurchasedCount = 0;
    gameState.hasStormProtection = false;
    gameState.stormShieldActive = false;
    
    // Switch to NOVA
    gameState.currentAsteroid = "NOVA";
    excelModel.ores = excelModel.novaOres;
    excelModel.gear = excelModel.novaGear;
    
    // Show warp overlay
    const warpOverlay = document.getElementById("warp-overlay");
    const desc = warpOverlay ? warpOverlay.querySelector("p") : null;
    const title = warpOverlay ? warpOverlay.querySelector("h1") : null;
    if (warpOverlay) {
      if (title) {
        title.textContent = "🌌 WARPING TO ASTEROID NOVA...";
        title.style.color = "var(--accent-purple)";
        title.style.textShadow = "0 0 15px rgba(208, 51, 255, 0.8)";
      }
      if (desc) desc.textContent = "Destination: Asteroid NOVA (10x Size). Initializing warp tunnel jump...";
      warpOverlay.classList.remove("hidden");
    }
    
    document.body.classList.add("explosion-shake");
    playWarpSound();
    
    setTimeout(() => {
      document.body.classList.remove("explosion-shake");
      if (warpOverlay) warpOverlay.classList.add("hidden");
      
      // Reset timer to 4 minutes (240s)
      gameState.timer = 240;
      excelModel.initialTimer = 240;
      
      // Reset appreciation timer to 10s for NOVA
      gameState.appreciationTimer = 10;
      gameState.appreciationTicks = 0;
      excelModel.novaOres.forEach(ore => ore.currentPrice = ore.baseValue);
      
      // Theme classes
      document.body.classList.remove("void-theme");
      document.body.classList.add("nova-theme");
      
      // Update badge
      DOM.asteroidName.textContent = "NOVA";
      DOM.asteroidName.className = "asteroid-nova-text";
      DOM.asteroidDistance.textContent = "(Far Outer Ring)";
      
      // Crystals wallet display
      const crystalPill = document.getElementById("crystal-pill-container");
      if (crystalPill) crystalPill.classList.remove("hidden");
      
      // Reset storm & earthquake schedulers
      gameState.elapsedSeconds = 0;
      gameState.scheduledStormSecond = null;
      gameState.stormWarningStarted = false;
      gameState.stormActive = false;
      
      gameState.elapsedEarthquakeSeconds = 0;
      gameState.scheduledEarthquakeSecond = null;
      gameState.earthquakeWarningStarted = false;
      gameState.earthquakeActive = false;
      
      if (DOM.stormPanel) {
        DOM.stormPanel.classList.add("hidden");
        DOM.btnLockShield.disabled = true;
        DOM.btnLockShield.textContent = "🔒 ACTIVATE STORM SHIELD";
        DOM.btnLockShield.style.background = "";
      }
      if (DOM.seismicPanel) DOM.seismicPanel.classList.add("hidden");
      
      // Reset gear variables
      gameState.hasSledgehammer = false;
      gameState.hasCosmicWater = false;
      gameState.hasResonanceStabilizer = false;
      gameState.hasDrillBore = false;
      gameState.tyliumUnits = 0;
      gameState.nukePurchasedCount = 0;
      excelModel.novaGear.nuke.costCrystals = 30; // Reset nuke cost
      
      // Recalculate speeds
      recalculateDiggingSpeeds();
      
      // Reset graph
      gameState.history.timestamps = [0];
      gameState.history.netWorths = [gameState.walletCash];
      drawNetWorthChart();
      
      gameState.isPaused = false;
      syncExcelParametersToGame();
      recalculateExcelModel();
      updateUI();
      
      writeConsoleLog("🌌 WARP JUMP COMPLETE! Welcome to Asteroid NOVA.");
      writeConsoleLog("⚠️ All VOID gear and gas left behind in orbit.");
      writeConsoleLog("🔮 Start manual digs to harvest NOVA ores and search for Crystals!");
    }, 4500);
  }

  // Initial flags list update
  updateFlagsTableUI();

  // Set up Excel table and bindings
  syncModelToDOM();
  
  // Wire up Welcome Screen modal overlay instead of starting immediately
  const welcomeModal = document.getElementById("welcome-modal");
  const welcomeInput = document.getElementById("welcome-name-input");
  const launchBtn = document.getElementById("btn-launch-game");
  
  if (welcomeModal && welcomeInput && launchBtn) {
    // Load last saved name into welcome input
    const savedName = safeStorage.getItem("asteroid_rush_player_name");
    if (savedName) {
      welcomeInput.value = savedName;
      if (DOM.pilotNameInput) DOM.pilotNameInput.value = savedName;
    }
    
    // Ensure game starts inactive and shows modal
    gameState.isActive = false;
    welcomeModal.classList.remove("hidden");
    
    // Clicking Launch starts the game
    launchBtn.addEventListener("click", () => {
      const pName = welcomeInput.value.trim() || "Miner One";
      safeStorage.setItem("asteroid_rush_player_name", pName);
      if (DOM.pilotNameInput) DOM.pilotNameInput.value = pName;
      
      welcomeModal.classList.add("hidden");
      startNewGame();
    });
  } else {
    // Fallback if welcome modal elements don't exist
    startNewGame();
  }
});
