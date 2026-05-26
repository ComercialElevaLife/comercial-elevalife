const BRL = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
const NUM = new Intl.NumberFormat("pt-BR");
const PCT = new Intl.NumberFormat("pt-BR", { style: "percent", minimumFractionDigits: 1, maximumFractionDigits: 1 });

const CHART_COLORS = {
  blue900: "#004657",
  blue700: "#007B82",
  blue500: "#00A1A0",
  blue200: "#7FC8CD",
  coral700: "#CC5E56",
  coral500: "#FF807D",
  coral300: "#FFBCB8",
};
const CHART_THEME = {
  realized: CHART_COLORS.blue500,
  quantity: CHART_COLORS.blue700,
  planned: CHART_COLORS.coral700,
  forecast: CHART_COLORS.coral500,
  conversion: CHART_COLORS.blue900,
  support: CHART_COLORS.blue200,
};
const UNIT_GEO_COORDS = {
  ARACARIGUAMA: { lat: -23.437, lon: -47.061 },
  CAMPINAS: { lat: -22.905, lon: -47.061 },
  CORPORATIVO: { lat: -23.55, lon: -46.633 },
  COSIGUA: { lat: -22.766, lon: -43.43 },
  COTIA: { lat: -23.603, lon: -46.918 },
  CUMBICA: { lat: -23.435, lon: -46.473 },
  HORTOLANDIA: { lat: -22.858, lon: -47.22 },
  LORENA: { lat: -22.733, lon: -45.124 },
  "MOGI DAS CRUZES": { lat: -23.522, lon: -46.188 },
  PINDAMONHANGABA: { lat: -22.924, lon: -45.462 },
  SBC: { lat: -23.691, lon: -46.565 },
  SJC: { lat: -23.189, lon: -45.884 },
  SP: { lat: -23.55, lon: -46.633 },
  "SANTA ROSA/RS": { lat: -27.87, lon: -54.48 },
  SANTOS: { lat: -23.961, lon: -46.332 },
  SOROCABA: { lat: -23.501, lon: -47.458 },
  SUZANO: { lat: -23.544, lon: -46.311 },
};

const STORAGE_KEY = "kpi_custom_vendas_v1";
const STORAGE_KEY_GERAL = "kpi_custom_geral_v1";
const STORAGE_KEY_CRM_VIEW = "kpi_crm_view_v1";
const STORAGE_KEY_MODE = "kpi_data_mode_v1";
const AUTO_CROSSMATCH_LEADS_FRONT = false;
const MAX_SALE_BEFORE_LEAD_DAYS = 15;
const HIDDEN_YEARS_FRONT = new Set(["2025"]);
const MONTH_ORDER = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"];
const MONTH_FULL = {
  JAN: "JANEIRO",
  FEV: "FEVEREIRO",
  MAR: "MARÇO",
  ABR: "ABRIL",
  MAI: "MAIO",
  JUN: "JUNHO",
  JUL: "JULHO",
  AGO: "AGOSTO",
  SET: "SETEMBRO",
  OUT: "OUTUBRO",
  NOV: "NOVEMBRO",
  DEZ: "DEZEMBRO",
};
const MONTH_ALIAS = {
  JANEIRO: "JAN", JANEIROS: "JAN", JAN: "JAN",
  FEVEREIRO: "FEV", FEV: "FEV",
  "MARÇO": "MAR", MARCO: "MAR", MAR: "MAR",
  ABRIL: "ABR", ABR: "ABR",
  MAIO: "MAI", MAI: "MAI",
  JUNHO: "JUN", JUN: "JUN",
  JULHO: "JUL", JUL: "JUL",
  AGOSTO: "AGO", AGO: "AGO",
  SETEMBRO: "SET", SET: "SET",
  OUTUBRO: "OUT", OUT: "OUT",
  NOVEMBRO: "NOV", NOV: "NOV",
  DEZEMBRO: "DEZ", DEZ: "DEZ",
};
const QUARTER_ORDER = ["Q1", "Q2", "Q3", "Q4"];
const KANBAN_STAGES = ["Lead", "Proposta", "Venda"];
const chartIds = [
  "chartFunilComercial",
  "chartMapaVendas",
  "chartPropSN",
  "chartLeadIcp",
  "chartLeadEntradas",
  "chartPropEnviadas",
  "chartPropSemana",
  "chartPropVolume",
  "chartVendas",
  "chartVendasServico",
  "chartVendasTipo",
  "chartVendasSocio",
  "chartAcumuladoMeta",
  "chartMetaGauge",
  "chartMetaCompare",
  "chartConvResponsavel",
  "chartConvCanal",
];
const state = {
  raw: null,
  mode: "original",
  customVendas: [],
  customGeral: [],
  filteredVendas: [],
  selectedLeadId: "",
  selectedLeadTaskId: "",
  editingLeadId: "",
  entradaSearch: "",
  servicoSearch: "",
  serviceSearch: "",
  leadServicesSelected: [],
  notifiedTaskKeys: new Set(),
  crmView: {
    kanbanHidden: false,
    kanbanCompact: true,
  },
  leadView: {
    page: 1,
    pageSize: 20,
    sortColumn: "Data",
    sortDirection: "desc",
    search: "",
    stage: "TODOS",
  },
  detailsView: {
    title: "",
    meta: "",
    columns: [],
    rows: [],
    page: 1,
    pageSize: 20,
    sortColumn: "",
    sortDirection: "asc",
  },
  filters: {
    ano: new Set(),
    quarter: new Set(),
    mes: new Set(),
    origem: new Set(),
    entrada: new Set(),
    servico: new Set(),
    tipo: new Set(),
  },
};

const refs = {
  updatedAt: document.getElementById("updatedAt"),
  chipsMes: document.getElementById("chipsMes"),
  chipsAno: document.getElementById("chipsAno"),
  chipsQuarter: document.getElementById("chipsQuarter"),
  chipsOrigem: document.getElementById("chipsOrigem"),
  entradaDropdownBtn: document.getElementById("entradaDropdownBtn"),
  entradaDropdown: document.getElementById("entradaDropdown"),
  servicoDropdownBtn: document.getElementById("servicoDropdownBtn"),
  servicoDropdown: document.getElementById("servicoDropdown"),
  chipsTipo: document.getElementById("chipsTipo"),
  btnRefreshData: document.getElementById("btnRefreshData"),
  clearFilters: document.getElementById("clearFilters"),
  btnExportExcel: document.getElementById("btnExportExcel"),
  btnExportPdf: document.getElementById("btnExportPdf"),
  kpiCount: document.getElementById("kpiCount"),
  kpiValue: document.getElementById("kpiValue"),
  kpiTicket: document.getElementById("kpiTicket"),
  kpiRecurring: document.getElementById("kpiRecurring"),
  toggleAdvancedAnalytics: document.getElementById("toggleAdvancedAnalytics"),
  advancedAnalyticsSection: document.getElementById("advancedAnalyticsSection"),
  aiLeadToProp: document.getElementById("aiLeadToProp"),
  aiPropToSale: document.getElementById("aiPropToSale"),
  aiWinRate: document.getElementById("aiWinRate"),
  aiForecast: document.getElementById("aiForecast"),
  aiAnomalies: document.getElementById("aiAnomalies"),
  detailsPanel: document.getElementById("detailsPanel"),
  detailsTitle: document.getElementById("detailsTitle"),
  detailsMeta: document.getElementById("detailsMeta"),
  detailsTable: document.getElementById("detailsTable"),
  detailsPageSize: document.getElementById("detailsPageSize"),
  detailsPrevPage: document.getElementById("detailsPrevPage"),
  detailsNextPage: document.getElementById("detailsNextPage"),
  detailsPageInfo: document.getElementById("detailsPageInfo"),
  closeDetails: document.getElementById("closeDetails"),
  overlay: document.getElementById("overlay"),
  tabButtons: [...document.querySelectorAll(".tab-btn[data-tab]")],
  painelLayout: document.getElementById("painelLayout"),
  toggleFilters: document.getElementById("toggleFilters"),
  tabPainel: document.getElementById("tab-painel"),
  tabInput: document.getElementById("tab-input"),
  fileInput: document.getElementById("fileInput"),
  btnUseOriginal: document.getElementById("btnUseOriginal"),
  btnUseCustom: document.getElementById("btnUseCustom"),
  btnClearCustom: document.getElementById("btnClearCustom"),
  dataModeInfo: document.getElementById("dataModeInfo"),
  leadForm: document.getElementById("leadForm"),
  leadFormSubmit: document.getElementById("leadFormSubmit"),
  cancelLeadEdit: document.getElementById("cancelLeadEdit"),
  leadServiceInput: document.getElementById("leadServiceInput"),
  leadServiceBtn: document.getElementById("leadServiceBtn"),
  leadServiceDropdown: document.getElementById("leadServiceDropdown"),
  leadServiceSearch: document.getElementById("leadServiceSearch"),
  leadServiceOptions: document.getElementById("leadServiceOptions"),
  leadServiceNew: document.getElementById("leadServiceNew"),
  leadServiceAdd: document.getElementById("leadServiceAdd"),
  leadEmpresaInput: document.getElementById("leadEmpresaInput"),
  leadNomeInput: document.getElementById("leadNomeInput"),
  leadCanalInput: document.getElementById("leadCanalInput"),
  leadResponsavelInput: document.getElementById("leadResponsavelInput"),
  leadCampanhaInput: document.getElementById("leadCampanhaInput"),
  leadNomeHint: document.getElementById("leadNomeHint"),
  leadEmpresaHint: document.getElementById("leadEmpresaHint"),
  leadEmpresaList: document.getElementById("leadEmpresaList"),
  leadNomeList: document.getElementById("leadNomeList"),
  leadCanalList: document.getElementById("leadCanalList"),
  leadResponsavelList: document.getElementById("leadResponsavelList"),
  leadCampanhaList: document.getElementById("leadCampanhaList"),
  recurringMonthsWrap: document.getElementById("recurringMonthsWrap"),
  recurringMonths: document.getElementById("recurringMonths"),
  recurringMonthlyValueWrap: document.getElementById("recurringMonthlyValueWrap"),
  recurringMonthlyValue: document.getElementById("recurringMonthlyValue"),
  recurringTotalWrap: document.getElementById("recurringTotalWrap"),
  recurringTotal: document.getElementById("recurringTotal"),
  leadTable: document.getElementById("leadTable"),
  leadModal: document.getElementById("leadModal"),
  leadModalOverlay: document.getElementById("leadModalOverlay"),
  closeLeadModal: document.getElementById("closeLeadModal"),
  leadModalMeta: document.getElementById("leadModalMeta"),
  leadModalForm: document.getElementById("leadModalForm"),
  leadDataSnapshot: document.getElementById("leadDataSnapshot"),
  leadChangeHistory: document.getElementById("leadChangeHistory"),
  leadSoldFlag: document.getElementById("leadSoldFlag"),
  leadDiscountFlag: document.getElementById("leadDiscountFlag"),
  leadSaleStatus: document.getElementById("leadSaleStatus"),
  leadNewValueWrap: document.getElementById("leadNewValueWrap"),
  leadNewValue: document.getElementById("leadNewValue"),
  leadSaleDate: document.getElementById("leadSaleDate"),
  leadSaleSocio: document.getElementById("leadSaleSocio"),
  leadSaleObs: document.getElementById("leadSaleObs"),
  saveLeadConversion: document.getElementById("saveLeadConversion"),
  leadTaskForm: document.getElementById("leadTaskForm"),
  taskType: document.getElementById("taskType"),
  taskOwner: document.getElementById("taskOwner"),
  taskDueDate: document.getElementById("taskDueDate"),
  taskStatus: document.getElementById("taskStatus"),
  taskDescription: document.getElementById("taskDescription"),
  taskSuggestion: document.getElementById("taskSuggestion"),
  leadTaskList: document.getElementById("leadTaskList"),
  leadTaskHistory: document.getElementById("leadTaskHistory"),
  leadPageSize: document.getElementById("leadPageSize"),
  leadPrevPage: document.getElementById("leadPrevPage"),
  leadNextPage: document.getElementById("leadNextPage"),
  leadPageInfo: document.getElementById("leadPageInfo"),
  leadSearch: document.getElementById("leadSearch"),
  leadStageFilter: document.getElementById("leadStageFilter"),
  crmTotalLeads: document.getElementById("crmTotalLeads"),
  crmTotalPropostas: document.getElementById("crmTotalPropostas"),
  crmTotalVendas: document.getElementById("crmTotalVendas"),
  crmTaxaConversao: document.getElementById("crmTaxaConversao"),
  crmValorPropostas: document.getElementById("crmValorPropostas"),
  crmValorVendas: document.getElementById("crmValorVendas"),
  crmStageBars: document.getElementById("crmStageBars"),
  crmKanban: document.getElementById("crmKanban"),
  crmWorkspace: document.getElementById("crmWorkspace"),
  crmKanbanCard: document.getElementById("crmKanbanCard"),
  toggleKanbanVisibility: document.getElementById("toggleKanbanVisibility"),
  toggleKanbanCompact: document.getElementById("toggleKanbanCompact"),
  enableTaskNotifications: document.getElementById("enableTaskNotifications"),
  taskNotificationStatus: document.getElementById("taskNotificationStatus"),
  customPreviewSection: document.getElementById("customPreviewSection"),
  customTable: document.getElementById("customTable"),
};

function isLocalRuntime() {
  const host = String(window.location.hostname || "").toLowerCase();
  return host === "localhost" || host === "127.0.0.1" || host === "::1";
}

function nonEmpty(value) {
  return value !== undefined && value !== null && String(value).trim() !== "";
}

function valueOrDash(value) {
  return nonEmpty(value) ? String(value) : "-";
}

function formatDateBr(value) {
  const text = String(value || "").trim();
  if (!text) return "-";
  const parsed = parseDateInput(text);
  if (parsed && !Number.isNaN(parsed.getTime())) return parsed.toLocaleDateString("pt-BR");
  return text;
}

function isDateColumn(column) {
  const key = normalizeKey(column);
  if (key === "PRAZO" || key === "DUEDATE") return true;
  if (key.startsWith("DATA")) return true;
  return false;
}

function isCurrencyColumn(column) {
  return /VALOR/.test(normalizeKey(column));
}

function formatColumnLabel(column) {
  if (isCurrencyColumn(column) && !String(column).includes("R$")) return `${column} (R$)`;
  return column;
}

function formatCellValue(column, value) {
  if (isDateColumn(column)) return formatDateBr(value);
  if (isCurrencyColumn(column)) return BRL.format(toNumber(value));
  if (typeof value === "number") return NUM.format(value);
  return valueOrDash(value);
}

function normalizePhoneDigits(value) {
  return String(value || "").replace(/\D/g, "");
}

function buildWhatsAppLink(value) {
  const digits = normalizePhoneDigits(value);
  if (!digits) return "";
  return `https://wa.me/${digits}`;
}

function toNumber(value) {
  if (typeof value === "number") return value;
  const text = String(value ?? "").trim();
  if (!text) return 0;
  let cleaned = text.replace(/[^0-9,.-]/g, "");
  if (!cleaned) return 0;
  if (cleaned.includes(",")) {
    cleaned = cleaned.replace(/\./g, "").replace(",", ".");
  } else if ((cleaned.match(/\./g) || []).length > 1) {
    const parts = cleaned.split(".");
    const last = parts.pop();
    cleaned = `${parts.join("")}.${last}`;
  }
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function parseDateInput(value) {
  const raw = String(value || "").trim();
  if (!raw) return null;
  const isoLike = raw.slice(0, 10);
  if (/^\d{4}-\d{2}-\d{2}$/.test(isoLike)) {
    const d = new Date(`${isoLike}T00:00:00`);
    return Number.isNaN(d.getTime()) ? null : d;
  }
  const brMatch = raw.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
  if (brMatch) {
    const dd = Number(brMatch[1]);
    const mm = Number(brMatch[2]);
    const yyRaw = Number(brMatch[3]);
    const yyyy = yyRaw < 100 ? 2000 + yyRaw : yyRaw;
    const d = new Date(yyyy, mm - 1, dd);
    if (!Number.isNaN(d.getTime()) && d.getDate() === dd && d.getMonth() === mm - 1 && d.getFullYear() === yyyy) {
      return d;
    }
  }
  const fallback = new Date(raw);
  return Number.isNaN(fallback.getTime()) ? null : fallback;
}

function normalizeMonth(text) {
  const t = String(text ?? "").trim().toUpperCase();
  return MONTH_ALIAS[t] || t;
}

function normalizeChannelName(value) {
  const cleaned = String(value || "").trim().replace(/\s+/g, " ");
  if (!cleaned) return "";
  if (/^[A-Z0-9&/ -]{2,10}$/.test(cleaned)) return cleaned.toUpperCase();
  return cleaned
    .toLowerCase()
    .split(" ")
    .map((part) => part ? part[0].toUpperCase() + part.slice(1) : part)
    .join(" ");
}

function serviceKey(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[.\s_-]/g, "");
}

function normalizeServiceName(value) {
  const cleaned = String(value || "").trim().replace(/\s+/g, " ");
  if (!cleaned) return "";
  const upper = cleaned.toUpperCase().replace(/\./g, "");
  const aliases = {
    ERGONOMIA: "ERGO",
    "GINASTICA LABORAL": "GL",
  };
  if (aliases[upper]) return aliases[upper];
  if (/^[A-Z0-9/& -]{2,18}$/.test(upper)) return upper;
  return cleaned
    .toLowerCase()
    .split(" ")
    .map((part) => part ? part[0].toUpperCase() + part.slice(1) : part)
    .join(" ");
}

function normalizeOrigemValue(value) {
  const raw = String(value || "").trim().toUpperCase();
  if (!raw) return "Não informado";
  if (raw === "S" || raw === "SOCIO" || raw === "SÓCIO" || raw === "SOCIOS" || raw === "SÓCIOS") return "Sócio";
  if (raw === "N" || raw === "NAO SOCIO" || raw === "NÃO SÓCIO" || raw === "NAO SOCIOS" || raw === "NÃO SÓCIOS") return "Não Sócio";
  if (raw.includes("NAO SOCIO") || raw.includes("NÃO SÓCIO") || raw.includes("NAO SOCIOS") || raw.includes("NÃO SÓCIOS")) return "Não Sócio";
  if (raw.includes("SOCIO") || raw.includes("SÓCIO")) return "Sócio";
  return "Não informado";
}

function origemToSocioFlag(origem) {
  const normalized = normalizeOrigemValue(origem);
  if (normalized === "Sócio") return "S";
  if (normalized === "Não Sócio") return "N";
  return "-";
}

function normalizeSaleStatus(value) {
  const raw = String(value || "").trim().toUpperCase();
  if (raw === "CANCELADA" || raw === "CANCELADO" || raw === "S" || raw === "SIM" || raw === "TRUE" || raw === "1") return "Cancelada";
  return "Ativa";
}

function isSaleCancelled(row) {
  return normalizeSaleStatus(row?.["Status Venda"] || row?.["Cancelada?"] || row?.Cancelada) === "Cancelada";
}

function getServiceCatalog() {
  const pool = [];
  const add = (value) => {
    const normalized = normalizeServiceName(value);
    if (normalized) pool.push(normalized);
  };
  (state.raw?.records?.vendas || []).forEach((row) => add(row["Serviço"]));
  (state.raw?.records?.geral || []).forEach((row) => add(row["Serviço"]));
  (state.customGeral || []).forEach((row) => add(row["Serviço"]));
  const seen = new Map();
  for (const item of pool) {
    const key = serviceKey(item);
    if (!seen.has(key)) seen.set(key, item);
  }
  return [...seen.values()].sort((a, b) => a.localeCompare(b, "pt-BR"));
}

function getSelectedLeadServices() {
  return [...state.leadServicesSelected];
}

function setSelectedLeadServices(values) {
  const normalizedList = (Array.isArray(values) ? values : [values])
    .map((value) => normalizeServiceName(value))
    .filter(nonEmpty);
  const uniq = [...new Set(normalizedList.map((value) => serviceKey(value)))]
    .map((key) => normalizedList.find((value) => serviceKey(value) === key))
    .filter(nonEmpty);
  state.leadServicesSelected = uniq;
  refs.leadServiceInput.value = uniq.join(", ");
  refs.leadServiceBtn.textContent = uniq.length ? `${uniq.length} serviço(s) selecionado(s)` : "Selecionar serviço(s)";
  refs.leadServiceInput.setCustomValidity(uniq.length ? "" : "Selecione ao menos um serviço.");
}

function updateLeadSuggestionLists() {
  const leads = getEditableLeads();
  const vendas = getActiveVendas();
  const toSorted = (items) => [...new Set(items.filter(nonEmpty))].sort((a, b) => String(a).localeCompare(String(b), "pt-BR"));

  const empresas = toSorted([
    ...leads.map((r) => String(r.Empresa || "").trim()),
    ...vendas.map((r) => String(r.Cliente || r.Empresa || "").trim()),
  ]);
  const nomes = toSorted(leads.map((r) => String(r.Nome || "").trim()));
  const canais = toSorted([
    ...leads.map((r) => String(r.Canal || "").trim()),
    ...vendas.map((r) => String(r.Canal || r.Entrada || "").trim()),
  ]);
  const responsaveis = toSorted(leads.map((r) => String(r.Entrada || "").trim()));
  const campanhas = toSorted(leads.map((r) => String(r.Campanha || "").trim()));

  const fillDatalist = (ref, values) => {
    if (!ref) return;
    ref.innerHTML = "";
    for (const value of values) {
      const option = document.createElement("option");
      option.value = value;
      ref.appendChild(option);
    }
  };

  fillDatalist(refs.leadEmpresaList, empresas);
  fillDatalist(refs.leadNomeList, nomes);
  fillDatalist(refs.leadCanalList, canais);
  fillDatalist(refs.leadResponsavelList, responsaveis);
  fillDatalist(refs.leadCampanhaList, campanhas);
  updateCreateNewHints({ empresas, nomes });
}

function updateCreateNewHints({ empresas = [], nomes = [] } = {}) {
  const setHint = (inputRef, hintRef, domain, label) => {
    if (!inputRef || !hintRef) return;
    const current = String(inputRef.value || "").trim();
    if (!current) {
      hintRef.textContent = "";
      return;
    }
    const exists = domain.some((item) => String(item || "").trim().toLowerCase() === current.toLowerCase());
    hintRef.textContent = exists
      ? `${label} já encontrado(a) na base.`
      : `${label} não encontrado(a). Continue digitando para cadastrar novo.`;
  };
  setHint(refs.leadEmpresaInput, refs.leadEmpresaHint, empresas, "Empresa");
  setHint(refs.leadNomeInput, refs.leadNomeHint, nomes, "Contato");
}

function refreshCreateNewHintsFromDatalist() {
  const empresas = [...(refs.leadEmpresaList?.querySelectorAll("option") || [])].map((opt) => String(opt.value || ""));
  const nomes = [...(refs.leadNomeList?.querySelectorAll("option") || [])].map((opt) => String(opt.value || ""));
  updateCreateNewHints({ empresas, nomes });
}

function renderLeadServiceOptions() {
  if (!refs.leadServiceOptions) return;
  const selected = getSelectedLeadServices();
  const selectedKeys = new Set(selected.map((service) => serviceKey(service)));
  const search = String(state.serviceSearch || "").toLowerCase();
  const options = getServiceCatalog().filter((service) => service.toLowerCase().includes(search));
  refs.leadServiceOptions.innerHTML = "";
  if (!options.length) {
    refs.leadServiceOptions.innerHTML = '<div class="hint">Nenhum serviço encontrado.</div>';
    return;
  }
  for (const service of options) {
    const row = document.createElement("label");
    row.className = "service-option";
    const input = document.createElement("input");
    input.type = "checkbox";
    input.checked = selectedKeys.has(serviceKey(service));
    input.addEventListener("change", () => {
      const next = new Set(getSelectedLeadServices().map((item) => serviceKey(item)));
      if (input.checked) next.add(serviceKey(service));
      else next.delete(serviceKey(service));
      const resolved = getServiceCatalog().filter((item) => next.has(serviceKey(item)));
      setSelectedLeadServices(resolved);
      renderLeadServiceOptions();
    });
    const text = document.createElement("span");
    text.textContent = service;
    row.appendChild(input);
    row.appendChild(text);
    refs.leadServiceOptions.appendChild(row);
  }
}

function calculateRecurringProposalValue() {
  const months = Math.max(0, Math.floor(toNumber(refs.recurringMonths?.value)));
  const monthly = toNumber(refs.recurringMonthlyValue?.value);
  const total = months > 0 && monthly > 0 ? months * monthly : 0;
  if (refs.recurringTotal) refs.recurringTotal.value = total > 0 ? BRL.format(total) : "";
  const proposalInput = refs.leadForm?.elements?.namedItem("Valor Proposta");
  if (proposalInput && String(refs.leadForm?.elements?.namedItem("TIPO")?.value || "").toUpperCase() === "R") {
    proposalInput.value = total > 0 ? String(Number(total.toFixed(2))) : "";
  }
}

function toggleRecurringFields() {
  const tipo = String(refs.leadForm?.elements?.namedItem("TIPO")?.value || "").toUpperCase();
  const isRecurring = tipo === "R";
  const proposalInput = refs.leadForm?.elements?.namedItem("Valor Proposta");
  if (refs.recurringMonthsWrap) refs.recurringMonthsWrap.hidden = !isRecurring;
  if (refs.recurringMonthlyValueWrap) refs.recurringMonthlyValueWrap.hidden = !isRecurring;
  if (refs.recurringTotalWrap) refs.recurringTotalWrap.hidden = !isRecurring;
  if (proposalInput) proposalInput.readOnly = isRecurring;
  if (!isRecurring) {
    if (refs.recurringMonths) refs.recurringMonths.value = "";
    if (refs.recurringMonthlyValue) refs.recurringMonthlyValue.value = "";
    if (refs.recurringTotal) refs.recurringTotal.value = "";
  } else {
    calculateRecurringProposalValue();
  }
}

function updateTaskNotificationStatus() {
  if (!refs.taskNotificationStatus) return;
  if (!("Notification" in window)) {
    refs.taskNotificationStatus.textContent = "Alertas desktop não suportados neste navegador.";
    return;
  }
  if (Notification.permission === "granted") {
    refs.taskNotificationStatus.textContent = "Alertas de tarefa: ativados";
  } else if (Notification.permission === "denied") {
    refs.taskNotificationStatus.textContent = "Alertas de tarefa: bloqueados no navegador.";
  } else {
    refs.taskNotificationStatus.textContent = "Alertas de tarefa: desativados";
  }
}

function notifyTask(title, body, key = "") {
  if (!("Notification" in window) || Notification.permission !== "granted") return;
  if (key && state.notifiedTaskKeys.has(key)) return;
  new Notification(title, { body });
  if (key) state.notifiedTaskKeys.add(key);
}

function remindDueTasks(leads) {
  const today = new Date();
  const todayKey = today.toISOString().slice(0, 10);
  for (const lead of leads || []) {
    const tasks = Array.isArray(lead.Tarefas) ? lead.Tarefas : [];
    for (const task of tasks) {
      const status = normalizeTaskStatus(task.status);
      if (["Concluída", "Cancelada"].includes(status)) continue;
      const due = String(task.prazo || "").slice(0, 10);
      if (!due || due > todayKey) continue;
      const key = `${lead.LeadId}|${task.id}|${due}`;
      notifyTask("Tarefa pendente no CRM", `${lead.Empresa} | ${task.tipo} | prazo ${due}`, key);
    }
  }
}

function monthFromDate(iso) {
  if (!iso || typeof iso !== "string") return "";
  const date = parseDateInput(iso);
  if (!date || Number.isNaN(date.getTime())) return "";
  return MONTH_ORDER[date.getMonth()] || "";
}

function yearFromDate(iso) {
  if (!iso || typeof iso !== "string") return "";
  const date = parseDateInput(iso);
  if (!date || Number.isNaN(date.getTime())) return "";
  return String(date.getFullYear());
}

function quarterFromMonthAbbrev(monthAbbrev) {
  const idx = MONTH_ORDER.indexOf(normalizeMonth(monthAbbrev));
  if (idx < 0) return "";
  return `Q${Math.floor(idx / 3) + 1}`;
}

function weekOfMonth(iso) {
  if (!iso || typeof iso !== "string") return "Sem data";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "Sem data";
  const dom = d.getDate();
  const week = Math.min(5, Math.floor((dom - 1) / 7) + 1);
  return `${week}ª semana`;
}

function uniqueSorted(records, key) {
  return [...new Set(records.map((r) => r[key]).filter(nonEmpty))].sort((a, b) => String(a).localeCompare(String(b), "pt-BR"));
}

function prepareVendasRecords(records) {
  const fallbackYear = String(new Date(state.raw?.generatedAt || Date.now()).getFullYear());
  return (records || []).map((row) => {
    const cliente = String(row.Cliente || row.Empresa || "").trim();
    const month = normalizeMonth(row["MÊS VENDA"] || monthFromDate(row["Data"]));
    const ano = String(row.Ano || yearFromDate(row["Data"]) || fallbackYear);
    const quarter = String(row.Quarter || quarterFromMonthAbbrev(month));
    const origem = normalizeOrigemValue(row.Origem || row.ORIGEM || row["Sócio?"] || row["SÓCIOS"] || row["SÓCIOS?"]);
    const socio = origemToSocioFlag(origem);
    const canal = normalizeChannelName(row["Canal"] || row["Entrada"] || "");
    const statusVenda = normalizeSaleStatus(row["Status Venda"] || row["Cancelada?"] || row.Cancelada);
    return {
      ...row,
      Cliente: cliente || row.Cliente,
      Empresa: cliente || row.Empresa,
      "MÊS VENDA": month,
      Ano: ano,
      Quarter: quarter,
      "Sócio?": socio,
      Origem: origem,
      Canal: canal,
      "Status Venda": statusVenda,
    };
  });
}

function prepareGeralRecords(records) {
  return (records || []).map((row) => {
    const month = normalizeMonth(row["Mês"] || monthFromDate(row["Data"]));
    const ano = String(yearFromDate(row["Data"]) || new Date(state.raw?.generatedAt || Date.now()).getFullYear());
    const quarter = quarterFromMonthAbbrev(month);
    const canal = normalizeChannelName(row["Canal"] || row["Fonte"] || row["Entrada"] || "");
    return normalizeLeadRecord({ ...row, "Mês": month, Ano: ano, Quarter: quarter, Canal: canal });
  }).filter((row) => row !== null);
}

function vendaMergeKey(row) {
  return [
    String(row?.Data || "").trim(),
    String(row?.Cliente || "").trim().toUpperCase(),
    String(row?.["Serviço"] || "").trim().toUpperCase(),
    String(row?.Tipo || "").trim().toUpperCase(),
    String(row?.["MÊS VENDA"] || "").trim().toUpperCase(),
  ].join("|");
}

function mergeVendas(baseRows, customRows) {
  const map = new Map();
  for (const row of baseRows || []) map.set(vendaMergeKey(row), row);
  for (const row of customRows || []) map.set(vendaMergeKey(row), row);
  return [...map.values()];
}

function geralMergeKey(row) {
  const leadId = String(row?.LeadId || "").trim();
  if (leadId) return `LEAD:${leadId}`;
  return [
    String(row?.Data || "").trim(),
    String(row?.Empresa || "").trim().toUpperCase(),
    String(row?.Nome || "").trim().toUpperCase(),
    String(row?.Serviço || "").trim().toUpperCase(),
  ].join("|");
}

function mergeGeral(baseRows, customRows) {
  const map = new Map();
  for (const row of baseRows || []) map.set(geralMergeKey(row), row);
  for (const row of customRows || []) map.set(geralMergeKey(row), row);
  return [...map.values()];
}

function isYearVisibleOnFront(year) {
  const normalized = String(year || "").trim();
  if (!normalized) return true;
  return !HIDDEN_YEARS_FRONT.has(normalized);
}

function getMergedVendasRaw() {
  const baseRows = state.raw?.records?.vendas || [];
  const merged = state.mode === "custom" && state.customVendas.length > 0
    ? mergeVendas(baseRows, state.customVendas)
    : baseRows;
  return merged.filter((row) => isYearVisibleOnFront(row?.Ano || yearFromDate(row?.Data)));
}

function getLatestVendasYear() {
  const years = getMergedVendasRaw()
    .map((row) => String(row?.Ano || "").trim())
    .filter((year) => /^\d{4}$/.test(year));
  if (!years.length) return "";
  return years.sort((a, b) => Number(a) - Number(b))[years.length - 1];
}

function ensureDefaultYearFilter() {
  if (state.filters.ano.size > 0) return;
  const latestYear = getLatestVendasYear();
  if (latestYear) state.filters.ano.add(latestYear);
}

function getActiveVendas() {
  return getMergedVendasRaw().filter((row) => !isSaleCancelled(row));
}

function getActiveGeral() {
  const baseRows = state.raw?.records?.geral || [];
  const merged = state.mode === "custom" && state.customGeral.length > 0
    ? mergeGeral(baseRows, state.customGeral)
    : baseRows;
  return merged.filter((row) => isYearVisibleOnFront(row?.Ano || yearFromDate(row?.Data)));
}

function getComparableGeral() {
  const geral = getActiveGeral();
  const entradaDomain = new Set(geral.map((r) => String(r["Canal"] || "").trim()).filter(nonEmpty));
  const entradaFiltro = [...state.filters.entrada].filter((v) => entradaDomain.has(v));
  const entradaFiltroAtivo = entradaFiltro.length > 0;
  return geral.filter((row) => {
    const leadAno = String(row["Ano"] || "");
    const leadQuarter = String(row["Quarter"] || "");
    const leadMes = normalizeMonth(row["Mês"]);

    const envioData = String(row["Data de envio"] || "").slice(0, 10);
    const envioAno = yearFromDate(envioData);
    const envioMes = normalizeMonth(monthFromDate(envioData));
    const envioQuarter = quarterFromMonthAbbrev(envioMes);

    if (state.filters.ano.size && !state.filters.ano.has(leadAno) && !state.filters.ano.has(envioAno)) return false;
    if (state.filters.quarter.size && !state.filters.quarter.has(leadQuarter) && !state.filters.quarter.has(envioQuarter)) return false;
    if (state.filters.mes.size && !state.filters.mes.has(leadMes) && !state.filters.mes.has(envioMes)) return false;
    if (entradaFiltroAtivo && !entradaFiltro.includes(String(row["Canal"] || "").trim())) return false;
    if (state.filters.servico.size && !state.filters.servico.has(row["Serviço"])) return false;
    if (state.filters.tipo.size && !state.filters.tipo.has(String(row["TIPO"] || "").toUpperCase())) return false;
    return true;
  });
}

function clearFilterSets() {
  Object.values(state.filters).forEach((set) => set.clear());
  state.entradaSearch = "";
  state.servicoSearch = "";
  ensureDefaultYearFilter();
}

function sanitizeFilterSelections() {
  const vendas = getActiveVendas();
  const geral = getActiveGeral();
  const valid = {
    ano: new Set(vendas.map((r) => r.Ano).filter(nonEmpty)),
    quarter: new Set(vendas.map((r) => r.Quarter).filter(nonEmpty)),
    mes: new Set(vendas.map((r) => normalizeMonth(r["MÊS VENDA"])).filter(nonEmpty)),
    origem: new Set(vendas.map((r) => r.Origem).filter(nonEmpty)),
    entrada: new Set([
      ...vendas.map((r) => String(r["Canal"] || "").trim()).filter(nonEmpty),
      ...geral.map((r) => String(r["Canal"] || "").trim()).filter(nonEmpty),
    ]),
    servico: new Set(vendas.map((r) => r["Serviço"]).filter(nonEmpty)),
    tipo: new Set(vendas.map((r) => String(r["Tipo"] || "").toUpperCase()).filter(nonEmpty)),
  };

  for (const key of Object.keys(state.filters)) {
    const set = state.filters[key];
    const allowed = valid[key] || new Set();
    for (const value of [...set]) {
      if (!allowed.has(value)) set.delete(value);
    }
  }
}

function getFilterValue(row, filterKey) {
  if (filterKey === "ano") return row["Ano"];
  if (filterKey === "quarter") return row["Quarter"];
  if (filterKey === "mes") return row["MÊS VENDA"];
  if (filterKey === "origem") return row["Origem"];
  if (filterKey === "entrada") return row["Canal"];
  if (filterKey === "servico") return row["Serviço"];
  if (filterKey === "tipo") return row["Tipo"];
  return "";
}

function passesSet(value, set) {
  return set.size === 0 || set.has(value);
}

function filterVendasRows(excludeKey = null) {
  const rows = getActiveVendas();
  return rows.filter((row) => {
    if (excludeKey !== "ano" && !passesSet(row["Ano"], state.filters.ano)) return false;
    if (excludeKey !== "quarter" && !passesSet(row["Quarter"], state.filters.quarter)) return false;
    if (excludeKey !== "mes" && !passesSet(row["MÊS VENDA"], state.filters.mes)) return false;
    if (excludeKey !== "origem" && !passesSet(row["Origem"], state.filters.origem)) return false;
    if (excludeKey !== "entrada" && !passesSet(row["Canal"], state.filters.entrada)) return false;
    if (excludeKey !== "servico" && !passesSet(row["Serviço"], state.filters.servico)) return false;
    if (excludeKey !== "tipo" && !passesSet(row["Tipo"], state.filters.tipo)) return false;
    return true;
  });
}

function relatedValueSet(filterKey) {
  const rows = filterVendasRows(filterKey);
  return new Set(rows.map((row) => getFilterValue(row, filterKey)).filter(nonEmpty));
}

function renderChips(container, values, filterKey, relatedSet = new Set()) {
  container.innerHTML = "";
  const hasAnyFilter = Object.values(state.filters).some((set) => set.size > 0);
  for (const value of values) {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "chip";
    chip.textContent = value;
    chip.dataset.value = value;
    if (state.filters[filterKey].has(value)) chip.classList.add("active");
    if (!state.filters[filterKey].has(value) && hasAnyFilter) {
      chip.classList.add(relatedSet.has(value) ? "related" : "unrelated");
    }
    chip.addEventListener("click", () => {
      const set = state.filters[filterKey];
      if (filterKey === "ano") {
        if (set.size === 1 && set.has(value)) {
          set.clear();
          ensureDefaultYearFilter();
        } else {
          set.clear();
          set.add(value);
        }
      } else if (set.has(value)) {
        set.delete(value);
      } else {
        set.add(value);
      }
      applyFilters();
    });
    container.appendChild(chip);
  }
}

function rebuildChipFilters() {
  const vendas = getActiveVendas();
  const geral = getActiveGeral();
  const anos = [...new Set(vendas.map((r) => r.Ano).filter(nonEmpty))].sort((a, b) => Number(a) - Number(b));
  const quarters = [...new Set(vendas.map((r) => r.Quarter).filter(nonEmpty))].sort((a, b) => QUARTER_ORDER.indexOf(a) - QUARTER_ORDER.indexOf(b));
  const meses = [...new Set(vendas.map((r) => normalizeMonth(r["MÊS VENDA"])).filter(nonEmpty))].sort((a, b) => MONTH_ORDER.indexOf(a) - MONTH_ORDER.indexOf(b));
  const origens = uniqueSorted(vendas, "Origem");
  const entradasVendas = uniqueSorted(vendas, "Canal");
  const entradasLeads = [...new Set(geral.map((r) => String(r["Canal"] || "").trim()).filter(nonEmpty))].sort((a, b) => a.localeCompare(b, "pt-BR"));
  const entradas = [...new Set([...entradasVendas, ...entradasLeads])];

  renderChips(refs.chipsAno, anos, "ano", relatedValueSet("ano"));
  renderChips(refs.chipsQuarter, quarters, "quarter", relatedValueSet("quarter"));
  renderChips(refs.chipsMes, meses, "mes", relatedValueSet("mes"));
  renderChips(refs.chipsOrigem, origens, "origem", relatedValueSet("origem"));
  renderEntradaDropdown(entradas, relatedValueSet("entrada"));
  renderServicoDropdown(uniqueSorted(vendas, "Serviço"), relatedValueSet("servico"));
  renderChips(refs.chipsTipo, uniqueSorted(vendas, "Tipo"), "tipo", relatedValueSet("tipo"));
}

function applyFilters() {
  ensureDefaultYearFilter();
  sanitizeFilterSelections();
  state.filteredVendas = filterVendasRows();
  rebuildChipFilters();
  renderAll();
}

function groupSum(records, key, valueKey) {
  const map = new Map();
  for (const row of records) {
    const k = valueOrDash(row[key]);
    map.set(k, (map.get(k) || 0) + toNumber(row[valueKey]));
  }
  return [...map.entries()].map(([label, value]) => ({ label, value }));
}

function groupCount(records, key) {
  const map = new Map();
  for (const row of records) {
    const k = valueOrDash(row[key]);
    map.set(k, (map.get(k) || 0) + 1);
  }
  return [...map.entries()].map(([label, value]) => ({ label, value }));
}

function openDetails(title, meta, rows, columns) {
  state.detailsView.title = title;
  state.detailsView.meta = meta;
  state.detailsView.columns = [...columns];
  state.detailsView.rows = [...rows];
  state.detailsView.page = 1;
  state.detailsView.sortColumn = "";
  state.detailsView.sortDirection = "asc";

  refs.detailsTitle.textContent = state.detailsView.title;
  refs.detailsMeta.textContent = state.detailsView.meta;
  refs.detailsPageSize.value = String(state.detailsView.pageSize);
  renderDetailsTable();

  refs.detailsPanel.classList.add("open");
  refs.detailsPanel.setAttribute("aria-hidden", "false");
  refs.overlay.hidden = false;
}

function parseSortNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const text = String(value ?? "").trim();
  if (!text) return null;
  const cleaned = text.replace(/\./g, "").replace(",", ".").replace(/[^0-9.-]/g, "");
  if (!cleaned || cleaned === "-" || cleaned === "." || cleaned === "-.") return null;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function compareDetailValues(a, b) {
  const aNum = parseSortNumber(a);
  const bNum = parseSortNumber(b);
  if (aNum !== null && bNum !== null) return aNum - bNum;
  return String(a ?? "").localeCompare(String(b ?? ""), "pt-BR", { sensitivity: "base" });
}

function getSortedDetailsRows() {
  const rows = [...state.detailsView.rows];
  const column = state.detailsView.sortColumn;
  if (!column) return rows;
  const direction = state.detailsView.sortDirection === "desc" ? -1 : 1;
  rows.sort((r1, r2) => direction * compareDetailValues(r1[column], r2[column]));
  return rows;
}

function renderDetailsTable() {
  const { columns, page, pageSize, sortColumn, sortDirection } = state.detailsView;
  const thead = refs.detailsTable.querySelector("thead");
  const tbody = refs.detailsTable.querySelector("tbody");
  thead.innerHTML = "";
  tbody.innerHTML = "";

  if (!columns.length) {
    refs.detailsPageInfo.textContent = "Página 1 de 1";
    refs.detailsPrevPage.disabled = true;
    refs.detailsNextPage.disabled = true;
    return;
  }

  const sortedRows = getSortedDetailsRows();
  const totalPages = Math.max(1, Math.ceil(sortedRows.length / pageSize));
  const safePage = Math.min(Math.max(1, page), totalPages);
  state.detailsView.page = safePage;
  const start = (safePage - 1) * pageSize;
  const pageRows = sortedRows.slice(start, start + pageSize);

  const headerRow = document.createElement("tr");
  for (const col of columns) {
    const th = document.createElement("th");
    th.textContent = formatColumnLabel(col);
    if (sortColumn === col) th.classList.add(sortDirection === "asc" ? "sort-asc" : "sort-desc");
    th.addEventListener("click", () => {
      if (state.detailsView.sortColumn === col) {
        state.detailsView.sortDirection = state.detailsView.sortDirection === "asc" ? "desc" : "asc";
      } else {
        state.detailsView.sortColumn = col;
        state.detailsView.sortDirection = "asc";
      }
      state.detailsView.page = 1;
      renderDetailsTable();
    });
    headerRow.appendChild(th);
  }
  thead.appendChild(headerRow);

  for (const row of pageRows) {
    const tr = document.createElement("tr");
    for (const col of columns) {
      const td = document.createElement("td");
      const val = row[col];
      if (normalizeKey(col).includes("TELEFONE")) {
        const url = buildWhatsAppLink(val);
        if (url) {
          const link = document.createElement("a");
          link.href = url;
          link.target = "_blank";
          link.rel = "noopener noreferrer";
          link.textContent = formatCellValue(col, val);
          td.appendChild(link);
        } else {
          td.textContent = formatCellValue(col, val);
        }
      } else {
        td.textContent = formatCellValue(col, val);
      }
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }

  refs.detailsPageInfo.textContent = `Página ${safePage} de ${totalPages}`;
  refs.detailsPrevPage.disabled = safePage <= 1;
  refs.detailsNextPage.disabled = safePage >= totalPages;
}

function closeDetails() {
  refs.detailsPanel.classList.remove("open");
  refs.detailsPanel.setAttribute("aria-hidden", "true");
  refs.overlay.hidden = true;
}

function baseLayout(extra = {}) {
  return {
    margin: { l: 72, r: 42, t: 72, b: 122 },
    paper_bgcolor: "rgba(0,0,0,0)",
    plot_bgcolor: "rgba(0,0,0,0)",
    font: { family: "Montserrat, sans-serif", color: CHART_COLORS.blue900, size: 14 },
    legend: { orientation: "h", y: -0.28, x: 0, xanchor: "left", font: { size: 13 }, itemsizing: "constant" },
    xaxis: { automargin: true, tickfont: { size: 12 }, tickangle: -20 },
    yaxis: { automargin: true, tickfont: { size: 12 } },
    hovermode: "x unified",
    uniformtext: { minsize: 10, mode: "hide" },
    ...extra,
  };
}

function updateKpis(records) {
  const count = records.length;
  const total = records.reduce((sum, r) => sum + toNumber(r["Valor"]), 0);
  const ticket = count ? total / count : 0;
  const recurringCount = records.filter((r) => String(r["Tipo"] || "").toUpperCase() === "R").length;
  const recurringRate = count ? recurringCount / count : 0;

  refs.kpiCount.textContent = NUM.format(count);
  refs.kpiValue.textContent = BRL.format(total);
  refs.kpiTicket.textContent = BRL.format(ticket);
  refs.kpiRecurring.textContent = PCT.format(recurringRate);
}

function monthSeries(rows, valueFn, options = {}) {
  const {
    dateField = "Data",
    monthField = "MÊS VENDA",
    yearField = "Ano",
    rowFilter = () => true,
  } = options;
  const byPeriod = new Map();
  for (const row of rows) {
    if (!rowFilter(row)) continue;
    const key = periodKeyFrom(row, dateField, monthField, yearField);
    if (!key) continue;
    byPeriod.set(key, (byPeriod.get(key) || 0) + toNumber(valueFn(row)));
  }
  const keys = [...byPeriod.keys()].sort((a, b) => a.localeCompare(b));
  return {
    keys,
    labels: keys.map(periodLabel),
    values: keys.map((k) => toNumber(byPeriod.get(k))),
  };
}

function movingAverage(values, windowSize = 3) {
  const out = [];
  for (let i = 0; i < values.length; i += 1) {
    const start = Math.max(0, i - windowSize + 1);
    const slice = values.slice(start, i + 1);
    const avg = slice.reduce((s, v) => s + toNumber(v), 0) / slice.length;
    out.push(avg);
  }
  return out;
}

function linearForecast(values, periodsAhead = 1) {
  if (!Array.isArray(values) || values.length < 2) return null;
  const xs = values.map((_, idx) => idx + 1);
  const ys = values.map((v) => toNumber(v));
  const n = ys.length;
  const sumX = xs.reduce((s, v) => s + v, 0);
  const sumY = ys.reduce((s, v) => s + v, 0);
  const sumXY = xs.reduce((s, x, i) => s + x * ys[i], 0);
  const sumXX = xs.reduce((s, x) => s + x * x, 0);
  const den = (n * sumXX) - (sumX * sumX);
  if (den === 0) return null;
  const slope = ((n * sumXY) - (sumX * sumY)) / den;
  const intercept = (sumY - (slope * sumX)) / n;
  const nextX = n + periodsAhead;
  return Math.max(0, intercept + (slope * nextX));
}

function detectAnomalies(values) {
  if (!Array.isArray(values) || values.length < 4) return [];
  const avg = values.reduce((s, v) => s + toNumber(v), 0) / values.length;
  const variance = values.reduce((s, v) => s + ((toNumber(v) - avg) ** 2), 0) / values.length;
  const std = Math.sqrt(variance);
  if (std === 0) return [];
  const anomalies = [];
  for (let i = 0; i < values.length; i += 1) {
    const z = (toNumber(values[i]) - avg) / std;
    if (Math.abs(z) >= 1.5) anomalies.push({ index: i, value: values[i], z });
  }
  return anomalies;
}

function renderAdvancedAnalytics(vendasRows, geralRows) {
  const totalLeads = geralRows.length;
  const propostasRows = geralRows.filter((r) => normalizePropostaStatus(r["Proposta Enviada ?"]) === "S");
  const totalPropostas = propostasRows.length;
  const totalVendas = vendasRows.length;
  const valorPropostas = propostasRows.reduce((s, r) => s + toNumber(r["Valor Proposta"]), 0);
  const valorVendas = vendasRows.reduce((s, r) => s + toNumber(r["Valor"]), 0);

  const leadToProp = totalLeads > 0 ? totalPropostas / totalLeads : 0;
  const propToSale = totalPropostas > 0 ? totalVendas / totalPropostas : 0;
  const winRateValue = valorPropostas > 0 ? valorVendas / valorPropostas : 0;

  const vendasSeries = monthSeries(vendasRows, (r) => toNumber(r["Valor"]));
  const forecast = linearForecast(vendasSeries.values, 1);
  const anomalies = detectAnomalies(vendasSeries.values);

  refs.aiLeadToProp.textContent = PCT.format(leadToProp);
  refs.aiPropToSale.textContent = PCT.format(propToSale);
  refs.aiWinRate.textContent = PCT.format(winRateValue);
  refs.aiForecast.textContent = forecast !== null ? BRL.format(forecast) : "n/d";
  refs.aiAnomalies.textContent = anomalies.length ? `${anomalies.length} ponto(s)` : "Sem desvios";
}

function getLeadsBySharedFilters(vendasRows) {
  return getEditableLeads().filter((lead) => passesCrmSharedFilters(lead, vendasRows));
}

function renderChartFunilComercial(vendasRows) {
  const perf = perfInfoVendas(vendasRows, { valueFn: (r) => toNumber(r["Valor"]) });
  const leads = getLeadsBySharedFilters(vendasRows);
  const propostas = leads.filter((lead) => leadHasProposta(lead, vendasRows));
  const vendas = leads.filter((lead) => leadHasVenda(lead, vendasRows));

  const etapas = ["Leads", "Propostas", "Vendas"];
  const valores = [leads.length, propostas.length, vendas.length];
  if (!valores.some((v) => v > 0)) return plotEmpty("chartFunilComercial", "Sem dados no funil para o filtro atual.", perf.text);

  Plotly.newPlot(
    "chartFunilComercial",
    [{
      type: "funnel",
      y: etapas,
      x: valores,
      textinfo: "value+percent previous",
      textposition: "inside",
      marker: { color: [CHART_COLORS.blue200, CHART_COLORS.blue500, CHART_COLORS.blue700] },
      customdata: etapas,
      hovertemplate: "<b>%{y}</b><br>Quantidade: %{x}<extra></extra>",
    }],
    baseLayout({
      margin: { l: 72, r: 32, t: 62, b: 58 },
      xaxis: { title: "Quantidade" },
      annotations: [
        ...perfAnnotation(perf.text),
        {
          xref: "paper",
          yref: "paper",
          x: 0.5,
          y: -0.2,
          showarrow: false,
          text: `Conversão Lead→Proposta: ${leads.length ? ((propostas.length / leads.length) * 100).toFixed(1).replace(".", ",") : "0,0"}% | Conversão Proposta→Venda: ${propostas.length ? ((vendas.length / propostas.length) * 100).toFixed(1).replace(".", ",") : "0,0"}%`,
          font: { size: 12, color: CHART_COLORS.blue900 },
        },
      ],
    }),
    { responsive: true }
  );

  document.getElementById("chartFunilComercial").on("plotly_click", (event) => {
    const etapa = String(event.points?.[0]?.customdata || "");
    let detailRows = leads;
    if (etapa === "Propostas") detailRows = propostas;
    if (etapa === "Vendas") detailRows = vendas;
    openDetails(
      `Funil Comercial: ${etapa}`,
      `${detailRows.length} lead(s) na etapa ${etapa}`,
      detailRows,
      ["Data", "Nome", "Empresa", "Canal", "Entrada", "Serviço", "Proposta Enviada ?", "Valor Proposta", "Status Lead"]
    );
  });
}

function normalizePropostaStatus(value) {
  const raw = String(value || "").trim().toUpperCase();
  return raw === "S" ? "S" : raw === "N" ? "N" : "-";
}

function normalizeLeadEntradaKey(value) {
  const text = String(value ?? "").trim().replace(/\s+/g, " ");
  if (!text) return "-";
  return normalizeKey(text);
}

function formatLeadEntradaLabel(value) {
  const text = String(value ?? "").trim().replace(/\s+/g, " ");
  if (!text) return "-";
  return text
    .split(" ")
    .map((word) => {
      const w = String(word || "");
      if (!w) return "";
      if (w === w.toUpperCase() && w.length <= 3) return w.toUpperCase();
      const lower = w.toLowerCase();
      return `${lower.charAt(0).toUpperCase()}${lower.slice(1)}`;
    })
    .join(" ");
}

function normalizeSocioStatus(row) {
  return normalizeOrigemValue(row["Origem"] || row["ORIGEM"] || row["Sócio?"] || row["SÓCIOS"] || row["SÓCIOS?"]);
}

function monthIndex(month) {
  const idx = MONTH_ORDER.indexOf(normalizeMonth(month));
  return idx >= 0 ? idx : 99;
}

function periodKeyFrom(row, dateField, monthField = "Mês", yearField = "Ano") {
  const iso = String(row[dateField] || "").slice(0, 10);
  const parsed = parseDateInput(iso);
  if (parsed && !Number.isNaN(parsed.getTime())) return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}`;
  const month = normalizeMonth(row[monthField]);
  const year = String(row[yearField] || "");
  const idx = MONTH_ORDER.indexOf(month);
  if (idx >= 0 && year) return `${year}-${String(idx + 1).padStart(2, "0")}`;
  return "";
}

function periodLabel(periodKey) {
  const [year, month] = String(periodKey || "").split("-");
  const idx = Number(month) - 1;
  if (!year || Number.isNaN(idx) || idx < 0 || idx > 11) return periodKey;
  return `${MONTH_ORDER[idx]}/${String(year).slice(-2)}`;
}

function periodKeyToDate(periodKey) {
  const [year, month] = String(periodKey || "").split("-");
  const y = Number(year);
  const m = Number(month);
  if (!Number.isFinite(y) || !Number.isFinite(m) || m < 1 || m > 12) return null;
  return new Date(y, m - 1, 1);
}

function dateToPeriodKey(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function addMonths(date, amount) {
  const out = new Date(date.getFullYear(), date.getMonth(), 1);
  out.setMonth(out.getMonth() + amount);
  return out;
}

function monthsBetweenInclusive(startDate, endDate) {
  return ((endDate.getFullYear() - startDate.getFullYear()) * 12) + (endDate.getMonth() - startDate.getMonth()) + 1;
}

function sumBucketRange(bucket, startDate, endDate) {
  let total = 0;
  const cursor = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  while (cursor <= endDate) {
    total += toNumber(bucket.get(dateToPeriodKey(cursor)));
    cursor.setMonth(cursor.getMonth() + 1);
  }
  return total;
}

function hasTemporalFilterSelected() {
  return state.filters.ano.size > 0 || state.filters.quarter.size > 0 || state.filters.mes.size > 0;
}

function passesVendasNonTemporalFilters(row) {
  if (!passesSet(row["Origem"], state.filters.origem)) return false;
  if (!passesSet(row["Canal"], state.filters.entrada)) return false;
  if (!passesSet(row["Serviço"], state.filters.servico)) return false;
  if (!passesSet(row["Tipo"], state.filters.tipo)) return false;
  return true;
}

function passesGeralNonTemporalFilters(row) {
  if (state.filters.entrada.size && !state.filters.entrada.has(String(row["Canal"] || "").trim())) return false;
  if (state.filters.servico.size && !state.filters.servico.has(row["Serviço"])) return false;
  if (state.filters.tipo.size && !state.filters.tipo.has(String(row["TIPO"] || "").toUpperCase())) return false;
  return true;
}

function perfInfo(rows, options) {
  const {
    dateField = "Data",
    monthField = "Mês",
    yearField = "Ano",
    rowFilter = () => true,
    valueFn = () => 1,
    baseRows = rows,
    nonTemporalFilter = () => true,
  } = options || {};

  const bucket = new Map();
  for (const row of baseRows) {
    if (!nonTemporalFilter(row) || !rowFilter(row)) continue;
    const key = periodKeyFrom(row, dateField, monthField, yearField);
    if (!key) continue;
    bucket.set(key, (bucket.get(key) || 0) + toNumber(valueFn(row)));
  }

  const selectedKeys = [...new Set(
    rows
      .filter((row) => rowFilter(row))
      .map((row) => periodKeyFrom(row, dateField, monthField, yearField))
      .filter(Boolean)
  )].sort((a, b) => a.localeCompare(b));
  if (!selectedKeys.length) return { text: "n/d", pct: null };

  let currentStart = periodKeyToDate(selectedKeys[0]);
  let currentEnd = periodKeyToDate(selectedKeys[selectedKeys.length - 1]);
  if (!currentStart || !currentEnd) return { text: "n/d", pct: null };

  if (!hasTemporalFilterSelected()) {
    currentStart = currentEnd;
  }

  const windowMonths = Math.max(1, monthsBetweenInclusive(currentStart, currentEnd));
  const prevEnd = addMonths(currentStart, -1);
  const prevStart = addMonths(prevEnd, -(windowMonths - 1));
  const current = sumBucketRange(bucket, currentStart, currentEnd);
  const previous = sumBucketRange(bucket, prevStart, prevEnd);
  if (previous <= 0) return { text: "n/d", pct: null };
  const pct = (current - previous) / previous;
  const signal = pct >= 0 ? "+" : "";
  return { text: `${signal}${PCT.format(pct)}`, pct };
}

function perfInfoVendas(rows, options = {}) {
  return perfInfo(rows, {
    dateField: "Data",
    monthField: "MÊS VENDA",
    yearField: "Ano",
    baseRows: getActiveVendas(),
    nonTemporalFilter: passesVendasNonTemporalFilters,
    ...options,
  });
}

function perfInfoGeral(rows, options = {}) {
  return perfInfo(rows, {
    dateField: "Data",
    monthField: "Mês",
    yearField: "Ano",
    baseRows: getActiveGeral(),
    nonTemporalFilter: passesGeralNonTemporalFilters,
    ...options,
  });
}

function perfAnnotation(perfText) {
  const text = String(perfText ?? "n/d");
  let color = CHART_COLORS.blue900;
  if (text.startsWith("+")) color = "#1E8E3E";
  if (text.startsWith("-")) color = "#C7352E";
  return [{
    xref: "paper",
    yref: "paper",
    x: 1,
    y: 1.18,
    xanchor: "right",
    showarrow: false,
    align: "right",
    text,
    font: { size: 11, color },
  }];
}

function plotEmpty(chartId, message, perfText) {
  const title = String(message || "Sem dados para exibir.");
  Plotly.newPlot(
    chartId,
    [],
    baseLayout({
      shapes: [{
        type: "rect",
        xref: "paper",
        yref: "paper",
        x0: 0.16,
        x1: 0.84,
        y0: 0.33,
        y1: 0.67,
        line: { color: "#CFE4E6", width: 1 },
        fillcolor: "#F7FCFC",
        layer: "below",
      }],
      annotations: [
        ...perfAnnotation(perfText),
        { xref: "paper", yref: "paper", x: 0.5, y: 0.56, showarrow: false, text: `<b>${title}</b>`, font: { size: 14, color: CHART_COLORS.blue900 } },
        { xref: "paper", yref: "paper", x: 0.5, y: 0.45, showarrow: false, text: "Ajuste os filtros ou selecione outro período.", font: { size: 11, color: "#5A7A80" } },
      ],
    }),
    { responsive: true }
  );
}

function renderChartPropostasSN(geralRows) {
  const perf = perfInfoGeral(geralRows, {
    rowFilter: (r) => normalizePropostaStatus(r["Proposta Enviada ?"]) === "S",
    valueFn: () => 1,
  });
  const grouped = groupCount(
    geralRows.map((r) => ({ ...r, _status: normalizePropostaStatus(r["Proposta Enviada ?"]) })),
    "_status"
  ).sort((a, b) => a.label.localeCompare(b.label, "pt-BR"));
  const labels = grouped.map((g) => g.label);
  const values = grouped.map((g) => g.value);
  if (!values.length) return plotEmpty("chartPropSN", "Sem dados para exibir.", perf.text);

  Plotly.newPlot(
    "chartPropSN",
    [{
      type: "pie",
      hole: 0.58,
      labels,
      values,
      textinfo: "percent",
      marker: { colors: [CHART_COLORS.blue500, CHART_COLORS.coral500, CHART_COLORS.blue200] },
      hovertemplate: "<b>Status %{label}</b><br>Leads: %{value}<br>%{percent}<extra></extra>",
    }],
    baseLayout({ margin: { l: 10, r: 10, t: 46, b: 8 }, annotations: perfAnnotation(perf.text) }),
    { responsive: true }
  );

  document.getElementById("chartPropSN").on("plotly_click", (event) => {
    const status = event.points?.[0]?.label;
    const detailRows = geralRows.filter((r) => normalizePropostaStatus(r["Proposta Enviada ?"]) === status);
    openDetails(`Propostas Enviadas (${status})`, `${detailRows.length} lead(s)`, detailRows, ["Data", "Nome", "Empresa", "Fonte", "Classificação", "Entrada", "Serviço", "Proposta Enviada ?", "Valor Proposta", "TIPO"]);
  });
}

function renderChartIcpLeads(geralRows) {
  if (!document.getElementById("chartLeadIcp")) return;
  const perf = perfInfoGeral(geralRows, { valueFn: () => 1 });
  const grouped = groupCount(geralRows, "Classificação").sort((a, b) => b.value - a.value);
  const labels = grouped.map((g) => g.label);
  const values = grouped.map((g) => g.value);
  if (!values.length) return plotEmpty("chartLeadIcp", "Sem dados para exibir.", perf.text);

  Plotly.newPlot(
    "chartLeadIcp",
    [{
      type: "pie",
      hole: 0.52,
      labels,
      values,
      textinfo: "percent",
      marker: { colors: [CHART_COLORS.blue700, CHART_COLORS.blue500, CHART_COLORS.blue200] },
      hovertemplate: "<b>ICP %{label}</b><br>Leads: %{value}<br>%{percent}<extra></extra>",
    }],
    baseLayout({ margin: { l: 16, r: 16, t: 50, b: 22 }, annotations: perfAnnotation(perf.text) }),
    { responsive: true }
  );

  document.getElementById("chartLeadIcp").on("plotly_click", (event) => {
    const icp = event.points?.[0]?.label;
    const detailRows = geralRows.filter((r) => valueOrDash(r["Classificação"]) === icp);
    openDetails(`ICP Leads: ${icp}`, `${detailRows.length} lead(s)`, detailRows, ["Data", "Nome", "Empresa", "Fonte", "Classificação", "Entrada", "Serviço", "Proposta Enviada ?", "Valor Proposta", "TIPO"]);
  });
}

function renderChartEntradasLeads(geralRows) {
  const perf = perfInfoGeral(geralRows, { valueFn: () => 1 });
  const normalizedRows = geralRows.map((row) => {
    const entradaLabel = formatLeadEntradaLabel(row["Canal"] || row["Fonte"] || row["Entrada"]);
    return { ...row, _entradaKey: normalizeLeadEntradaKey(entradaLabel), _entradaLabel: entradaLabel };
  });

  const groupedMap = new Map();
  for (const row of normalizedRows) {
    const key = row._entradaKey;
    if (!groupedMap.has(key)) groupedMap.set(key, { key, label: row._entradaLabel, value: 0 });
    groupedMap.get(key).value += 1;
  }
  const grouped = [...groupedMap.values()].sort((a, b) => b.value - a.value);
  const y = grouped.map((g) => g.label);
  const x = grouped.map((g) => g.value);
  const keys = grouped.map((g) => g.key);
  if (!x.length) return plotEmpty("chartLeadEntradas", "Sem dados para exibir.", perf.text);

  Plotly.newPlot(
    "chartLeadEntradas",
    [{ type: "bar", orientation: "h", y, x, text: x.map((v) => NUM.format(v)), textposition: y.length <= 8 ? "auto" : "none", marker: { color: CHART_COLORS.blue500 }, customdata: keys, hovertemplate: "<b>%{y}</b><br>Entradas: %{x}<extra></extra>" }],
    baseLayout({ margin: { l: 140, r: 20, t: 54, b: 40 }, annotations: perfAnnotation(perf.text) }),
    { responsive: true }
  );

  document.getElementById("chartLeadEntradas").on("plotly_click", (event) => {
    const entradaKey = event.points?.[0]?.customdata;
    const entradaLabel = event.points?.[0]?.y || "-";
    const detailRows = normalizedRows
      .filter((r) => r._entradaKey === entradaKey)
      .map((r) => ({ ...r, Canal: r._entradaLabel }));
    openDetails(`Entrada de Leads: ${entradaLabel}`, `${detailRows.length} lead(s)`, detailRows, ["Data", "Nome", "Empresa", "Canal", "Fonte", "Classificação", "Serviço", "Proposta Enviada ?", "Valor Proposta", "TIPO"]);
  });
}

function renderChartPropostasEnviadas(geralRows) {
  const onlySent = geralRows.filter((r) => normalizePropostaStatus(r["Proposta Enviada ?"]) === "S");
  const perf = perfInfoGeral(geralRows, { rowFilter: (r) => normalizePropostaStatus(r["Proposta Enviada ?"]) === "S", valueFn: () => 1 });
  const series = monthSeries(onlySent, () => 1, { dateField: "Data de envio", monthField: "Mês", yearField: "Ano" });
  const x = [...series.labels];
  const y = [...series.values];
  const periodKeys = [...series.keys];
  if (!y.length) return plotEmpty("chartPropEnviadas", "Sem propostas enviadas no filtro.", perf.text);

  Plotly.newPlot(
    "chartPropEnviadas",
    [{ type: "bar", x, y, text: y.map((v) => NUM.format(v)), textposition: y.length <= 6 ? "outside" : "none", cliponaxis: false, marker: { color: CHART_COLORS.coral700 }, customdata: periodKeys, hovertemplate: "<b>%{x}</b><br>Propostas: %{y}<extra></extra>" }],
    baseLayout({ annotations: perfAnnotation(perf.text) }),
    { responsive: true }
  );

  document.getElementById("chartPropEnviadas").on("plotly_click", (event) => {
    const periodKey = String(event.points?.[0]?.customdata || "");
    if (!periodKey) return;
    const detailRows = onlySent.filter((r) => periodKeyFrom(r, "Data de envio", "Mês", "Ano") === periodKey);
    openDetails(`Propostas Enviadas: ${periodLabel(periodKey)}`, `${detailRows.length} proposta(s)`, detailRows, ["Data", "Data de envio", "Nome", "Empresa", "Fonte", "Classificação", "Entrada", "Serviço", "Proposta Enviada ?", "Valor Proposta", "TIPO"]);
  });
}

function renderChartPropostasSemana(geralRows) {
  const onlySent = geralRows.filter((r) => normalizePropostaStatus(r["Proposta Enviada ?"]) === "S");
  const perf = perfInfoGeral(geralRows, { rowFilter: (r) => normalizePropostaStatus(r["Proposta Enviada ?"]) === "S", valueFn: () => 1 });
  const order = ["1ª semana", "2ª semana", "3ª semana", "4ª semana", "5ª semana"];
  const grouped = groupCount(onlySent, "Semana Envio").sort((a, b) => order.indexOf(a.label) - order.indexOf(b.label));
  const x = grouped.map((g) => g.label);
  const y = grouped.map((g) => g.value);
  if (!y.length) return plotEmpty("chartPropSemana", "Sem propostas enviadas no filtro.", perf.text);

  Plotly.newPlot(
    "chartPropSemana",
    [{ type: "bar", x, y, text: y.map((v) => NUM.format(v)), textposition: y.length <= 8 ? "auto" : "none", marker: { color: CHART_COLORS.blue200 }, customdata: x, hovertemplate: "<b>%{x}</b><br>Propostas: %{y}<extra></extra>" }],
    baseLayout({ annotations: perfAnnotation(perf.text) }),
    { responsive: true }
  );

  document.getElementById("chartPropSemana").on("plotly_click", (event) => {
    const semana = event.points?.[0]?.customdata;
    const detailRows = onlySent.filter((r) => valueOrDash(r["Semana Envio"]) === semana);
    openDetails(`Envio por Semana: ${semana}`, `${detailRows.length} proposta(s)`, detailRows, ["Data", "Data de envio", "Semana Envio", "Nome", "Empresa", "Fonte", "Serviço", "Valor Proposta"]);
  });
}

function renderChartVolumePropostas(geralRows) {
  const onlySent = geralRows.filter((r) => normalizePropostaStatus(r["Proposta Enviada ?"]) === "S");
  const perf = perfInfoGeral(geralRows, {
    rowFilter: (r) => normalizePropostaStatus(r["Proposta Enviada ?"]) === "S",
    valueFn: (r) => toNumber(r["Valor Proposta"]),
  });
  const series = monthSeries(onlySent, (r) => toNumber(r["Valor Proposta"]), { dateField: "Data de envio", monthField: "Mês", yearField: "Ano" });
  const x = [...series.labels];
  const y = [...series.values];
  const periodKeys = [...series.keys];
  if (!y.length) return plotEmpty("chartPropVolume", "Sem volume de propostas no filtro.", perf.text);

  Plotly.newPlot(
    "chartPropVolume",
    [{ type: "bar", x, y, text: y.map((v) => BRL.format(v)), textposition: y.length <= 6 ? "outside" : "none", cliponaxis: false, marker: { color: CHART_COLORS.coral500 }, customdata: periodKeys, hovertemplate: "<b>%{x}</b><br>Volume: %{text}<extra></extra>" }],
    baseLayout({ yaxis: { tickprefix: "R$ ", separatethousands: true }, annotations: perfAnnotation(perf.text) }),
    { responsive: true }
  );

  document.getElementById("chartPropVolume").on("plotly_click", (event) => {
    const periodKey = String(event.points?.[0]?.customdata || "");
    if (!periodKey) return;
    const detailRows = onlySent.filter((r) => periodKeyFrom(r, "Data de envio", "Mês", "Ano") === periodKey);
    const total = detailRows.reduce((sum, r) => sum + toNumber(r["Valor Proposta"]), 0);
    openDetails(`Volume de Propostas: ${periodLabel(periodKey)}`, `Volume total: ${BRL.format(total)}`, detailRows, ["Data", "Data de envio", "Nome", "Empresa", "Fonte", "Serviço", "Valor Proposta", "TIPO"]);
  });
}

function renderChartVendas(vendasRows) {
  const perf = perfInfoVendas(vendasRows, { valueFn: (r) => toNumber(r["Valor"]) });
  const countSeries = monthSeries(vendasRows, () => 1, { dateField: "Data", monthField: "MÊS VENDA", yearField: "Ano" });
  const valueSeries = monthSeries(vendasRows, (r) => toNumber(r["Valor"]), { dateField: "Data", monthField: "MÊS VENDA", yearField: "Ano" });
  const x = [...valueSeries.labels];
  const periodKeys = [...valueSeries.keys];
  const yValue = [...valueSeries.values];
  const yCountMap = new Map(countSeries.keys.map((k, idx) => [k, countSeries.values[idx]]));
  const yCount = periodKeys.map((k) => toNumber(yCountMap.get(k)));
  const anomalies = detectAnomalies(yValue);
  const nextForecast = linearForecast(yValue, 1);
  const lastPeriodDate = periodKeyToDate(periodKeys[periodKeys.length - 1]);
  const nextPeriodDate = lastPeriodDate ? addMonths(lastPeriodDate, 1) : null;
  const nextPeriodKey = nextPeriodDate ? dateToPeriodKey(nextPeriodDate) : "";
  const nextPeriodLabel = nextPeriodKey ? periodLabel(nextPeriodKey) : `${x[x.length - 1]}+1`;
  const forecastX = nextForecast !== null && x.length ? [...x, nextPeriodLabel] : x;
  const forecastY = nextForecast !== null && x.length ? [...yValue, nextForecast] : yValue;
  if (!x.length) return plotEmpty("chartVendas", "Sem vendas no filtro.", perf.text);

  Plotly.newPlot(
    "chartVendas",
    [
      { type: "bar", x, y: yCount, name: "Qtd Vendas", text: yCount.map((v) => NUM.format(v)), textposition: x.length <= 6 ? "auto" : "none", marker: { color: CHART_THEME.quantity }, customdata: periodKeys, hovertemplate: "<b>%{x}</b><br>Qtd: %{y}<extra></extra>" },
      { type: "scatter", mode: "lines+markers", x, y: yValue, yaxis: "y2", customdata: periodKeys, name: "Valor Vendas", line: { color: CHART_THEME.realized, width: 3 }, marker: { size: 8 }, hovertemplate: "<b>%{x}</b><br>Valor: R$ %{y:,.2f}<extra></extra>" },
      ...(nextForecast !== null && x.length ? [{
        type: "scatter",
        mode: "lines+markers",
        x: forecastX,
        y: forecastY,
        customdata: [...periodKeys, nextPeriodKey || "FORECAST"],
        yaxis: "y2",
        name: "Previsão (1 mês)",
        line: { color: CHART_THEME.forecast, width: 2, dash: "dash" },
        marker: { size: 7 },
        hovertemplate: "<b>%{x}</b><br>Previsão: R$ %{y:,.2f}<extra></extra>",
      }] : []),
      ...(anomalies.length ? [{
        type: "scatter",
        mode: "markers+text",
        x: anomalies.map((a) => x[a.index]),
        y: anomalies.map((a) => yValue[a.index]),
        customdata: anomalies.map((a) => periodKeys[a.index]),
        yaxis: "y2",
        name: "Anomalia",
        text: anomalies.map(() => "Alerta"),
        textposition: "top center",
        marker: { color: CHART_THEME.planned, size: 12, symbol: "diamond" },
        hovertemplate: "<b>%{x}</b><br>Valor atípico: R$ %{y:,.2f}<extra></extra>",
      }] : []),
    ],
    baseLayout({
      legend: { orientation: "h", y: -0.25 },
      xaxis: { tickangle: -22, automargin: true },
      yaxis: { title: "Qtd" },
      yaxis2: { title: "Valor (R$)", overlaying: "y", side: "right", tickprefix: "R$ ", separatethousands: true, tickfont: { size: 12 } },
      annotations: perfAnnotation(perf.text),
    }),
    { responsive: true }
  );

  document.getElementById("chartVendas").on("plotly_click", (event) => {
    const periodKey = String(event.points?.[0]?.customdata || "");
    if (!periodKey || periodKey === "FORECAST") return;
    const detailRows = vendasRows.filter((r) => periodKeyFrom(r, "Data", "MÊS VENDA", "Ano") === periodKey);
    const total = detailRows.reduce((sum, r) => sum + toNumber(r["Valor"]), 0);
    openDetails(`Vendas: ${periodLabel(periodKey)}`, `${detailRows.length} venda(s) | ${BRL.format(total)}`, detailRows, ["Data", "Cliente", "Serviço", "Entrada", "ICP", "Tipo", "Origem", "Valor"]);
  });
}

function renderChartVendasServico(vendasRows) {
  const perf = perfInfoVendas(vendasRows, { valueFn: (r) => toNumber(r["Valor"]) });
  const grouped = groupSum(vendasRows, "Serviço", "Valor").sort((a, b) => b.value - a.value);
  const x = grouped.map((g) => g.label);
  const y = grouped.map((g) => g.value);
  if (!x.length) return plotEmpty("chartVendasServico", "Sem vendas por serviço no filtro.", perf.text);

  Plotly.newPlot(
    "chartVendasServico",
    [{ type: "bar", x, y, text: y.map((v) => BRL.format(v)), textposition: x.length <= 6 ? "outside" : "none", cliponaxis: false, marker: { color: CHART_THEME.realized }, customdata: x, hovertemplate: "<b>%{x}</b><br>Valor: %{text}<extra></extra>" }],
    baseLayout({ xaxis: { tickangle: -18, automargin: true }, yaxis: { tickprefix: "R$ ", separatethousands: true }, annotations: perfAnnotation(perf.text) }),
    { responsive: true }
  );

  document.getElementById("chartVendasServico").on("plotly_click", (event) => {
    const servico = event.points?.[0]?.customdata;
    const detailRows = vendasRows.filter((r) => valueOrDash(r["Serviço"]) === servico);
    openDetails(`Vendas por Serviço: ${servico}`, `${detailRows.length} venda(s)`, detailRows, ["Data", "Cliente", "Serviço", "Entrada", "ICP", "Tipo", "Origem", "Valor"]);
  });
}

function normalizeGeoUnit(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

function resolveUnitCoords(unit) {
  const key = normalizeGeoUnit(unit);
  if (!key) return null;
  if (UNIT_GEO_COORDS[key]) return UNIT_GEO_COORDS[key];
  if (key.includes("SANTA ROSA")) return UNIT_GEO_COORDS["SANTA ROSA/RS"];
  if (key.includes("MOGI")) return UNIT_GEO_COORDS["MOGI DAS CRUZES"];
  if (key.includes("PINDAMONHANGABA")) return UNIT_GEO_COORDS.PINDAMONHANGABA;
  if (key.includes("SOROCABA")) return UNIT_GEO_COORDS.SOROCABA;
  if (key.includes("CAMPINAS")) return UNIT_GEO_COORDS.CAMPINAS;
  if (key.includes("SUZANO")) return UNIT_GEO_COORDS.SUZANO;
  if (key.includes("SANTOS")) return UNIT_GEO_COORDS.SANTOS;
  if (key.includes("HORTOLANDIA")) return UNIT_GEO_COORDS.HORTOLANDIA;
  if (key.includes("LONDRINA")) return { lat: -23.304, lon: -51.169 };
  return null;
}

function renderChartMapaVendas(vendasRows) {
  if (!document.getElementById("chartMapaVendas")) return;
  const perf = perfInfoVendas(vendasRows, { valueFn: (r) => toNumber(r["Valor"]) });
  const grouped = new Map();
  for (const row of vendasRows || []) {
    const unit = String(row["Unidade"] || row["Cliente"] || "").trim();
    if (!unit) continue;
    const coords = resolveUnitCoords(unit);
    if (!coords) continue;
    const key = `${coords.lat}|${coords.lon}|${unit}`;
    if (!grouped.has(key)) grouped.set(key, { unit, lat: coords.lat, lon: coords.lon, qtd: 0, valor: 0 });
    const g = grouped.get(key);
    g.qtd += 1;
    g.valor += toNumber(row["Valor"]);
  }
  const points = [...grouped.values()];
  if (!points.length) return plotEmpty("chartMapaVendas", "Sem coordenadas de local de venda para exibir no mapa.", perf.text);

  const maxQtd = Math.max(...points.map((p) => p.qtd), 1);
  Plotly.newPlot(
    "chartMapaVendas",
    [{
      type: "scattergeo",
      mode: "markers",
      lat: points.map((p) => p.lat),
      lon: points.map((p) => p.lon),
      text: points.map((p) => `${p.unit}<br>Vendas: ${p.qtd}<br>Valor: ${BRL.format(p.valor)}`),
      customdata: points.map((p) => p.unit),
      marker: {
        size: points.map((p) => 14 + (p.qtd / maxQtd) * 20),
        color: points.map((p) => p.valor),
        colorscale: [[0, "#BFE6E8"], [0.5, "#37A5AA"], [1, "#005E6D"]],
        showscale: true,
        colorbar: { title: "Valor (R$)" },
        opacity: 0.85,
        line: { color: "#ffffff", width: 1 },
      },
      hovertemplate: "%{text}<extra></extra>",
    }],
    baseLayout({
      margin: { l: 12, r: 12, t: 46, b: 12 },
      geo: {
        scope: "south america",
        center: { lat: -23.5, lon: -47.5 },
        projection: { type: "mercator" },
        showland: true,
        landcolor: "#edf6f7",
        showcountries: true,
        countrycolor: "#b8d5d8",
        showocean: true,
        oceancolor: "#f8fcfd",
      },
      annotations: perfAnnotation(perf.text),
    }),
    { responsive: true }
  );

  document.getElementById("chartMapaVendas").on("plotly_click", (event) => {
    const unit = String(event.points?.[0]?.customdata || "");
    if (!unit) return;
    const detailRows = vendasRows.filter((r) => String(r["Unidade"] || r["Cliente"] || "").trim() === unit);
    const total = detailRows.reduce((sum, row) => sum + toNumber(row["Valor"]), 0);
    openDetails(
      `Mapa de Vendas: ${unit}`,
      `${detailRows.length} venda(s) | ${BRL.format(total)}`,
      detailRows,
      ["Data", "Cliente", "Unidade", "Serviço", "Entrada", "Tipo", "Origem", "Valor"]
    );
  });
}

function renderChartVendasTipo(vendasRows) {
  const perf = perfInfoVendas(vendasRows, { valueFn: (r) => toNumber(r["Valor"]) });
  const grouped = groupSum(vendasRows, "Tipo", "Valor").sort((a, b) => b.value - a.value);
  const labels = grouped.map((g) => g.label);
  const values = grouped.map((g) => g.value);
  if (!labels.length) return plotEmpty("chartVendasTipo", "Sem mix de tipo no filtro.", perf.text);

  Plotly.newPlot(
    "chartVendasTipo",
    [{ type: "pie", hole: 0.58, labels, values, textinfo: "percent", marker: { colors: [CHART_COLORS.coral500, CHART_COLORS.blue500, CHART_COLORS.blue200] }, hovertemplate: "<b>Tipo %{label}</b><br>Valor: R$ %{value:,.2f}<br>%{percent}<extra></extra>" }],
    baseLayout({ margin: { l: 10, r: 10, t: 46, b: 8 }, annotations: perfAnnotation(perf.text) }),
    { responsive: true }
  );

  document.getElementById("chartVendasTipo").on("plotly_click", (event) => {
    const tipo = event.points?.[0]?.label;
    const detailRows = vendasRows.filter((r) => String(r["Tipo"] || "").toUpperCase() === String(tipo || "").toUpperCase());
    openDetails(`Vendas ${tipo}`, `${detailRows.length} venda(s)`, detailRows, ["Data", "Cliente", "Serviço", "Entrada", "ICP", "Tipo", "Origem", "Valor"]);
  });
}

function renderChartVendasSocio(vendasRows) {
  const enriched = vendasRows.map((r) => ({ ...r, _socio: normalizeSocioStatus(r) }));
  const perf = perfInfoVendas(enriched, {
    rowFilter: (r) => r._socio === "Sócio",
    valueFn: (r) => toNumber(r["Valor"]),
  });
  const grouped = groupSum(enriched, "_socio", "Valor").sort((a, b) => a.label.localeCompare(b.label, "pt-BR"));
  const labels = grouped.map((g) => g.label);
  const values = grouped.map((g) => g.value);
  if (!labels.length) return plotEmpty("chartVendasSocio", "Sem dados de origem no filtro.", perf.text);

  Plotly.newPlot(
    "chartVendasSocio",
    [{ type: "pie", hole: 0.58, labels, values, textinfo: "percent", marker: { colors: [CHART_COLORS.blue700, CHART_COLORS.coral700, CHART_COLORS.blue200] }, hovertemplate: "<b>Origem %{label}</b><br>Valor: R$ %{value:,.2f}<br>%{percent}<extra></extra>" }],
    baseLayout({ margin: { l: 10, r: 10, t: 46, b: 8 }, annotations: perfAnnotation(perf.text) }),
    { responsive: true }
  );

  document.getElementById("chartVendasSocio").on("plotly_click", (event) => {
    const socio = event.points?.[0]?.label;
    const detailRows = enriched.filter((r) => r._socio === socio);
    openDetails(`Vendas por Origem (${socio})`, `${detailRows.length} venda(s)`, detailRows, ["Data", "Cliente", "Serviço", "Entrada", "ICP", "Tipo", "Origem", "Valor"]);
  });
}

function sumMetaMonthsFromRow(row) {
  if (!row || typeof row !== "object") return 0;
  return MONTH_ORDER.reduce((sum, abbr) => sum + toNumber(row[MONTH_FULL[abbr]]), 0);
}

function getAcumuladoMetricRows(token) {
  const rows = state.raw?.records?.acumulado || [];
  const wanted = normalizeKey(token);
  return rows.filter((row) => {
    const metric = normalizeKey(row["Métrica"] || "");
    return metric === wanted || metric.includes(wanted);
  });
}

function getAcumuladoMetricRow(token) {
  return getAcumuladoMetricRows(token)[0] || null;
}

function rowMagnitudeScore(row) {
  const monthly = sumMetaMonthsFromRow(row);
  const acumulado = Math.abs(toNumber(row?.["ACUMULADO"]));
  return Math.max(Math.abs(monthly), acumulado);
}

function selectAcumuladoValueRow(...tokens) {
  const all = tokens.flatMap((token) => getAcumuladoMetricRows(token));
  if (!all.length) return null;
  return all
    .slice()
    .sort((a, b) => rowMagnitudeScore(b) - rowMagnitudeScore(a))[0] || null;
}

function selectAcumuladoPercentRow(...tokens) {
  const all = tokens.flatMap((token) => getAcumuladoMetricRows(token));
  if (!all.length) return null;
  const percentileCandidates = all.filter((row) => {
    const ac = Math.abs(toNumber(row?.["ACUMULADO"]));
    const monthly = Math.abs(sumMetaMonthsFromRow(row));
    return ac > 0 && ac <= 2 && monthly <= 20;
  });
  if (percentileCandidates.length) {
    return percentileCandidates
      .slice()
      .sort((a, b) => Math.abs(toNumber(b?.["ACUMULADO"])) - Math.abs(toNumber(a?.["ACUMULADO"])))[0] || null;
  }
  return all[0] || null;
}

function getScopedVendaPeriods(vendasRows) {
  return [...new Set(
    (vendasRows || [])
      .map((row) => periodKeyFrom(row, "Data", "MÊS VENDA", "Ano"))
      .filter(Boolean)
  )].sort((a, b) => String(a).localeCompare(String(b)));
}

function getGaugeValuesFromAcumulado(vendasRows) {
  // Não usa percentual bruto da aba ACUMULADO para o velocímetro porque
  // essa aba pode conter meses de anos distintos (ex.: DEZ/25 e 2026),
  // e o painel precisa respeitar o ano filtrado no front.
  return null;
}

function getAnnualMetaContext(vendasRows) {
  const recurringMetric = selectAcumuladoValueRow("META PREVISTA RECORRENTES", "META RECOR.", "META RECORRENTES");
  const punctualMetric = selectAcumuladoValueRow("META PREVISTA PONTUAIS", "META PONT.", "META PONTUAIS");
  const splitMeta = sumMetaMonthsFromRow(recurringMetric) + sumMetaMonthsFromRow(punctualMetric);
  if (splitMeta > 0) return { value: splitMeta, source: "ACUMULADO", estimated: false };

  const totalMetric = selectAcumuladoValueRow("META TOTAL");
  const totalFromMonths = sumMetaMonthsFromRow(totalMetric);
  if (totalFromMonths > 0) return { value: totalFromMonths, source: "ACUMULADO", estimated: false };

  const dashMeta = toNumber(state.raw?.records?.dash?.meta);
  if (dashMeta > 0) return { value: dashMeta, source: "DASH", estimated: false };

  const byMonth = monthSeries(vendasRows, (r) => toNumber(r["Valor"]));
  if (byMonth.values.length > 0) {
    const avg = byMonth.values.reduce((sum, value) => sum + toNumber(value), 0) / byMonth.values.length;
    const estimatedAnnual = avg * 12;
    if (estimatedAnnual > 0) return { value: estimatedAnnual, source: "Estimativa", estimated: true };
  }
  return { value: 0, source: "Sem meta", estimated: true };
}

function getMetaProgressContext(vendasRows) {
  const gaugeFromBase = getGaugeValuesFromAcumulado(vendasRows);
  const metaContext = getAnnualMetaContext(vendasRows);
  const annualMeta = metaContext.value;
  const totalAtingidoFallback = vendasRows.reduce((sum, r) => sum + toNumber(r["Valor"]), 0);

  const periodsFromScope = getScopedVendaPeriods(vendasRows).length;
  const periodsScope = Math.max(1, periodsFromScope || 1);
  const metaPeriodoFallback = annualMeta > 0 ? (annualMeta / 12) * periodsScope : 0;

  const metaPeriodo = gaugeFromBase ? gaugeFromBase.metaPeriodo : metaPeriodoFallback;
  const totalAtingido = gaugeFromBase ? gaugeFromBase.totalAtingido : totalAtingidoFallback;
  const atingimentoPeriodo = metaPeriodo > 0 ? (totalAtingido / metaPeriodo) * 100 : 0;
  const atingimentoMetaFinal = annualMeta > 0 ? (totalAtingido / annualMeta) * 100 : 0;
  const atingimento = gaugeFromBase ? gaugeFromBase.atingimento : atingimentoMetaFinal;
  const source = gaugeFromBase?.source || metaContext.source;
  const metaLabel = metaContext.estimated && !gaugeFromBase ? `Meta estimada (${source})` : `Meta (${source})`;
  const resumoAtingimento = metaPeriodo <= 0
    ? "Meta não definida"
    : `Meta final atingida: ${atingimentoMetaFinal.toFixed(1).replace(".", ",")}%`;

  return {
    metaPeriodo,
    totalAtingido,
    atingimento,
    atingimentoPeriodo,
    atingimentoMetaFinal,
    annualMeta,
    source,
    metaLabel,
    resumoAtingimento,
  };
}

function renderChartAcumuladoMeta(vendasRows) {
  const perf = perfInfoVendas(vendasRows, { valueFn: (r) => toNumber(r["Valor"]) });
  const series = monthSeries(vendasRows, (r) => toNumber(r["Valor"]), { dateField: "Data", monthField: "MÊS VENDA", yearField: "Ano" });
  if (!series.keys.length) return plotEmpty("chartAcumuladoMeta", "Sem dados acumulados no filtro.", perf.text);

  const metaContext = getAnnualMetaContext(vendasRows);
  const annualMeta = metaContext.value;
  const monthlyMeta = annualMeta > 0 ? annualMeta / 12 : 0;

  let runningReal = 0;
  let runningMeta = 0;
  let currentYear = "";
  const x = [...series.labels];
  const periodKeys = [...series.keys];
  const yReal = [];
  const yMeta = [];
  const yPct = [];
  for (let i = 0; i < series.values.length; i += 1) {
    const periodYear = String(periodKeys[i] || "").split("-")[0];
    if (currentYear && periodYear && periodYear !== currentYear) {
      runningReal = 0;
      runningMeta = 0;
    }
    if (periodYear) currentYear = periodYear;
    runningReal += toNumber(series.values[i]);
    runningMeta += monthlyMeta;
    yReal.push(runningReal);
    yMeta.push(runningMeta);
    yPct.push(runningMeta > 0 ? (runningReal / runningMeta) * 100 : 0);
  }

  const nextAccForecast = linearForecast(yReal, 1);
  const lastPeriodDate = periodKeyToDate(periodKeys[periodKeys.length - 1]);
  const nextPeriodDate = lastPeriodDate ? addMonths(lastPeriodDate, 1) : null;
  const nextPeriodKey = nextPeriodDate ? dateToPeriodKey(nextPeriodDate) : "";
  const nextPeriodLabel = nextPeriodKey ? periodLabel(nextPeriodKey) : `${x[x.length - 1]}+1`;
  const xForecast = nextAccForecast !== null ? [...x, nextPeriodLabel] : x;
  const yAccForecast = nextAccForecast !== null ? [...yReal, nextAccForecast] : yReal;

  Plotly.newPlot(
    "chartAcumuladoMeta",
    [
      {
        type: "bar",
        x,
        y: yReal,
        name: "Meta Atingida (Acumulado)",
        text: yReal.map((v) => BRL.format(v)),
        textposition: yReal.length <= 6 ? "outside" : "none",
        cliponaxis: false,
        marker: { color: CHART_THEME.realized },
        customdata: periodKeys,
        hovertemplate: "<b>%{x}</b><br>Atingida acumulada: %{text}<extra></extra>",
      },
      {
        type: "scatter",
        mode: "lines+markers",
        x,
        y: yMeta,
        customdata: periodKeys,
        name: "Meta Prevista (Acumulada)",
        line: { color: CHART_THEME.planned, width: 3 },
        marker: { size: 8 },
        hovertemplate: "<b>%{x}</b><br>Meta acumulada: R$ %{y:,.2f}<extra></extra>",
      },
      {
        type: "scatter",
        mode: "lines+markers",
        x,
        y: yPct,
        customdata: periodKeys,
        yaxis: "y2",
        name: "Atingimento Acumulado (%)",
        line: { color: CHART_THEME.conversion, width: 2, dash: "dot" },
        marker: { size: 6 },
        hovertemplate: "<b>%{x}</b><br>Atingimento acumulado: %{y:.1f}%<extra></extra>",
      },
      ...(nextAccForecast !== null ? [{
        type: "scatter",
        mode: "lines+markers",
        x: xForecast,
        y: yAccForecast,
        customdata: [...periodKeys, nextPeriodKey || "FORECAST"],
        name: "Previsão acumulada",
        line: { color: CHART_THEME.forecast, width: 2, dash: "dash" },
        marker: { size: 7 },
        hovertemplate: "<b>%{x}</b><br>Previsão acumulada: R$ %{y:,.2f}<extra></extra>",
      }] : []),
    ],
    baseLayout({
      legend: { orientation: "h", y: -0.25 },
      yaxis: { tickprefix: "R$ ", separatethousands: true },
      yaxis2: { overlaying: "y", side: "right", ticksuffix: "%", range: [0, Math.max(120, Math.max(...yPct) * 1.2)] },
      annotations: perfAnnotation(perf.text),
    }),
    { responsive: true }
  );

  document.getElementById("chartAcumuladoMeta").on("plotly_click", (event) => {
    const periodKey = String(event.points?.[0]?.customdata || "");
    if (!periodKey || periodKey === "FORECAST" || !periodKeys.includes(periodKey)) return;
    const idx = periodKeys.indexOf(periodKey);
    const periodText = x[idx] || periodLabel(periodKey);
    const targetYear = String(periodKey).split("-")[0];
    const detailRows = vendasRows.filter((r) => {
      const rowKey = periodKeyFrom(r, "Data", "MÊS VENDA", "Ano");
      if (!rowKey) return false;
      const rowYear = String(rowKey).split("-")[0];
      return rowYear === targetYear && rowKey <= periodKey;
    });
    openDetails(
      `Acumulado até ${periodText}`,
      `Meta prevista acumulada: ${BRL.format(yMeta[idx] || 0)} | Meta atingida acumulada: ${BRL.format(yReal[idx] || 0)} | Atingimento acumulado: ${Number(yPct[idx] || 0).toFixed(1)}%`,
      detailRows,
      ["Data", "Cliente", "Serviço", "Entrada", "ICP", "Tipo", "Origem", "Valor", "MÊS VENDA"]
    );
  });
}

function renderChartMetaGauge(vendasRows) {
  const perf = perfInfoVendas(vendasRows, { valueFn: (r) => toNumber(r["Valor"]) });
  const context = getMetaProgressContext(vendasRows);
  const {
    metaPeriodo,
    totalAtingido,
    annualMeta,
    atingimento,
    atingimentoPeriodo,
    atingimentoMetaFinal,
    metaLabel,
    resumoAtingimento,
  } = context;
  const gaugeMax = Math.max(120, Math.ceil(atingimento / 10) * 10, 200);

  Plotly.newPlot(
    "chartMetaGauge",
    [{
      type: "indicator",
      mode: "gauge+number",
      value: atingimento,
      number: { suffix: "%", valueformat: ".1f" },
      gauge: {
        axis: { range: [0, gaugeMax] },
        bar: { color: CHART_THEME.realized },
        steps: [
          { range: [0, 70], color: "#fde8e7" },
          { range: [70, 100], color: "#fff5dd" },
          { range: [100, gaugeMax], color: "#e6f7ec" },
        ],
        threshold: { line: { color: CHART_THEME.planned, width: 4 }, thickness: 0.85, value: 100 },
      },
      hovertemplate: "<b>Percentual alcançado</b><br>%{value:.1f}%<extra></extra>",
    }],
    baseLayout({
      margin: { l: 22, r: 22, t: 56, b: 26 },
      annotations: [
        ...perfAnnotation(perf.text),
        {
          xref: "paper",
          yref: "paper",
          x: 0.5,
          y: -0.2,
          showarrow: false,
          text: `Meta acumulada do período: ${BRL.format(metaPeriodo)}<br>Realizado no período: ${BRL.format(totalAtingido)} | Meta final anual: ${BRL.format(annualMeta)}<br>Atingimento período: ${atingimentoPeriodo.toFixed(1).replace(".", ",")}% | Atingimento meta final: ${atingimentoMetaFinal.toFixed(1).replace(".", ",")}%`,
          font: { size: 12, color: CHART_COLORS.blue900 },
        },
      ],
    }),
    { responsive: true }
  );

  document.getElementById("chartMetaGauge").on("plotly_click", () => {
    openDetails(
      "Meta (Velocímetro)",
      `${resumoAtingimento} | Meta acumulada do período: ${BRL.format(metaPeriodo)} | Realizado: ${BRL.format(totalAtingido)} | Meta final anual: ${BRL.format(annualMeta)} | Atingimento período: ${atingimentoPeriodo.toFixed(1).replace(".", ",")}%`,
      vendasRows,
      ["Data", "Cliente", "Serviço", "Entrada", "ICP", "Tipo", "Origem", "Valor", "MÊS VENDA"]
    );
  });
}

function renderChartMetaCompare(vendasRows) {
  const perf = perfInfoVendas(vendasRows, { valueFn: (r) => toNumber(r["Valor"]) });
  const context = getMetaProgressContext(vendasRows);
  const { metaPeriodo, totalAtingido, atingimento, metaLabel, resumoAtingimento } = context;
  if (metaPeriodo <= 0 && totalAtingido <= 0) {
    return plotEmpty("chartMetaCompare", "Sem dados de meta para comparação.", perf.text);
  }

  const x = [metaLabel, "Realizado"];
  const y = [metaPeriodo, totalAtingido];
  const colors = [CHART_THEME.planned, CHART_THEME.realized];

  Plotly.newPlot(
    "chartMetaCompare",
    [{
      type: "bar",
      x,
      y,
      marker: { color: colors },
      text: y.map((value) => BRL.format(value)),
      textposition: "outside",
      cliponaxis: false,
      hovertemplate: "<b>%{x}</b><br>Valor: R$ %{y:,.2f}<extra></extra>",
    }],
    baseLayout({
      margin: { l: 42, r: 20, t: 56, b: 56 },
      yaxis: { tickprefix: "R$ ", separatethousands: true },
      annotations: [
        ...perfAnnotation(perf.text),
        {
          xref: "paper",
          yref: "paper",
          x: 0.5,
          y: -0.22,
          showarrow: false,
          text: `${resumoAtingimento}<br>${metaLabel}: ${BRL.format(metaPeriodo)} | Realizado: ${BRL.format(totalAtingido)}`,
          font: { size: 12, color: CHART_COLORS.blue900 },
        },
      ],
    }),
    { responsive: true }
  );

  document.getElementById("chartMetaCompare").on("plotly_click", () => {
    openDetails(
      "Meta x Realizado (R$)",
      `${resumoAtingimento} | ${metaLabel}: ${BRL.format(metaPeriodo)} | Realizado: ${BRL.format(totalAtingido)}`,
      vendasRows,
      ["Data", "Cliente", "Serviço", "Entrada", "ICP", "Tipo", "Origem", "Valor", "MÊS VENDA"]
    );
  });
}

function normalizeResponsibleLabel(value) {
  const cleaned = String(value || "").trim().replace(/\s+/g, " ");
  if (!cleaned) return "-";
  return cleaned
    .toLowerCase()
    .split(" ")
    .map((part) => (part ? part[0].toUpperCase() + part.slice(1) : part))
    .join(" ");
}

function responsibleGroupKey(value) {
  const cleaned = String(value || "").trim().replace(/\s+/g, " ");
  if (!cleaned) return "-";
  return cleaned
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase();
}

function conversionGroupKey(value, keyName) {
  if (keyName === "Entrada") return responsibleGroupKey(value);
  return valueOrDash(value);
}

function conversionGroupLabel(value, keyName) {
  if (keyName === "Entrada") return normalizeResponsibleLabel(value);
  return valueOrDash(value);
}

function buildConversionGroup(rows, vendasRows, keyName) {
  const groups = new Map();
  for (const row of rows) {
    const matchKey = conversionGroupKey(row[keyName], keyName);
    const label = conversionGroupLabel(row[keyName], keyName);
    if (!groups.has(matchKey)) groups.set(matchKey, { matchKey, label, leads: 0, propostas: 0, vendas: 0 });
    const g = groups.get(matchKey);
    if (g.label === "-" && label !== "-") g.label = label;
    g.leads += 1;
    if (String(row["Proposta Enviada ?"] || "").toUpperCase() === "S") g.propostas += 1;
    if (getLeadStage(row, vendasRows) === "Venda" || inferLeadSoldFromVendas(row, vendasRows)) g.vendas += 1;
  }
  return [...groups.values()].map((g) => ({
    ...g,
    taxa: g.propostas > 0 ? (g.vendas / g.propostas) * 100 : 0,
  })).sort((a, b) => b.taxa - a.taxa);
}

function renderConversionChart(chartId, title, geralRows, vendasRows, keyName) {
  const perf = perfInfoVendas(vendasRows, { valueFn: (r) => toNumber(r["Valor"]) });
  const grouped = buildConversionGroup(geralRows, vendasRows, keyName);
  const x = grouped.map((g) => g.label);
  const propostas = grouped.map((g) => g.propostas);
  const vendas = grouped.map((g) => g.vendas);
  const taxa = grouped.map((g) => Number(g.taxa.toFixed(1)));
  const groupKeys = grouped.map((g) => g.matchKey);
  if (!x.length) return plotEmpty(chartId, `Sem dados de ${title.toLowerCase()}.`, perf.text);

  Plotly.newPlot(
    chartId,
    [
      { type: "bar", x, y: propostas, customdata: groupKeys, name: "Propostas", marker: { color: CHART_THEME.support }, hovertemplate: "<b>%{x}</b><br>Propostas: %{y}<extra></extra>" },
      { type: "bar", x, y: vendas, customdata: groupKeys, name: "Vendas", marker: { color: CHART_THEME.realized }, hovertemplate: "<b>%{x}</b><br>Vendas: %{y}<extra></extra>" },
      { type: "scatter", mode: "lines+markers", x, y: taxa, customdata: groupKeys, yaxis: "y2", name: "Conversão (%)", line: { color: CHART_THEME.conversion, width: 2 }, marker: { size: 7 }, hovertemplate: "<b>%{x}</b><br>Conversão: %{y:.1f}%<extra></extra>" },
    ],
    baseLayout({
      barmode: "group",
      legend: { orientation: "h", y: -0.25 },
      yaxis: { title: "Qtd" },
      yaxis2: { overlaying: "y", side: "right", ticksuffix: "%", range: [0, 100] },
      annotations: perfAnnotation(perf.text),
    }),
    { responsive: true }
  );

  document.getElementById(chartId).on("plotly_click", (event) => {
    const groupKey = String(event.points?.[0]?.customdata || "");
    if (!groupKey) return;
    const detailRows = geralRows.filter((r) => conversionGroupKey(r[keyName], keyName) === groupKey).map((r) => ({
      ...r,
      "Status Lead": getLeadStage(r, vendasRows),
      "Convertido?": inferLeadSoldFromVendas(r, vendasRows) ? "S" : "N",
    }));
    const groupTitle = grouped.find((g) => g.matchKey === groupKey)?.label || groupKey;
    openDetails(
      `${title}: ${groupTitle}`,
      `${detailRows.length} lead(s) | Conversão baseada em proposta x venda`,
      detailRows,
      ["Data", "Nome", "Empresa", "Canal", "Entrada", "Classificação", "Serviço", "Proposta Enviada ?", "Valor Proposta", "Status Lead", "Convertido?"]
    );
  });
}

function renderAll() {
  const vendasRows = state.filteredVendas;
  const geralRows = getComparableGeral();
  const editableLeads = getEditableLeads();

  updateKpis(vendasRows);
  renderChartFunilComercial(vendasRows);
  renderAdvancedAnalytics(vendasRows, geralRows);
  renderChartPropostasSN(geralRows);
  renderChartIcpLeads(geralRows);
  renderChartEntradasLeads(geralRows);
  renderChartPropostasEnviadas(geralRows);
  renderChartPropostasSemana(geralRows);
  renderChartVolumePropostas(geralRows);
  renderChartVendas(vendasRows);
  renderChartVendasServico(vendasRows);
  renderChartMapaVendas(vendasRows);
  renderChartVendasTipo(vendasRows);
  renderChartVendasSocio(vendasRows);
  renderChartAcumuladoMeta(vendasRows);
  renderChartMetaGauge(vendasRows);
  renderChartMetaCompare(vendasRows);
  renderConversionChart("chartConvResponsavel", "Conversão por Responsável", geralRows, vendasRows, "Entrada");
  renderConversionChart("chartConvCanal", "Conversão por Canal", geralRows, vendasRows, "Canal");

  updateDataModeInfo();
  renderCustomPreview();
  renderLeadTable();
  remindDueTasks(editableLeads);
}

function normalizeKey(key) {
  return String(key || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Za-z0-9]/g, "")
    .toUpperCase();
}

const KEY_ALIASES = {
  Data: ["DATA"],
  Cliente: ["CLIENTE", "EMPRESA", "NOME"],
  Unidade: ["UNIDADE", "FILIAL", "PLANTA", "SITE"],
  "Serviço": ["SERVICO", "SERVICO", "PRODUTO"],
  "Sócio?": ["SOCIO", "SOCIOS", "ASSOCIADO", "TIPOSOCIO"],
  Entrada: ["ENTRADA", "CANAL", "FONTE"],
  "Reativação?": ["REATIVACAO", "REATIVACAO"],
  ICP: ["ICP", "CLASSIFICACAO"],
  Tipo: ["TIPO"],
  "Valor Serviço": ["VALORSERVICO", "VALORMENSAL", "VALORUNITARIO", "VALORBASE"],
  "Meses Recorrência": ["MESESRECORRENCIA", "MESES", "QTDMESES", "PERIODORECORRENCIA"],
  Valor: ["VALOR", "VALORPROPOSTA"],
  "Status Venda": ["STATUSVENDA", "CANCELADA", "CANCELADA?"],
  "MÊS VENDA": ["MESVENDA", "MES", "MÊSVENDA"],
  Ano: ["ANO", "YEAR"],
  Quarter: ["QUARTER", "TRIMESTRE"],
  Origem: ["ORIGEM"],
};

function pickByAliases(row, aliases) {
  const map = {};
  for (const [k, v] of Object.entries(row)) map[normalizeKey(k)] = v;
  for (const key of aliases) {
    const norm = normalizeKey(key);
    if (norm in map && nonEmpty(map[norm])) return map[norm];
  }
  return "";
}

function normalizeVendaRecord(row) {
  const normalized = {};
  for (const [dest, aliases] of Object.entries(KEY_ALIASES)) normalized[dest] = pickByAliases(row, aliases);

  normalized.Data = String(normalized.Data || "").slice(0, 10);
  normalized.Cliente = String(normalized.Cliente || "").trim();
  normalized.Unidade = String(normalized.Unidade || "").trim();
  normalized["Serviço"] = normalizeServiceName(normalized["Serviço"]);
  normalized["Sócio?"] = String(normalized["Sócio?"] || "").trim();
  normalized.Entrada = String(normalized.Entrada || "").trim();
  normalized["Reativação?"] = String(normalized["Reativação?"] || "Não").trim() || "Não";
  normalized.ICP = String(normalized.ICP || "").trim().toUpperCase();
  normalized.Tipo = String(normalized.Tipo || "").trim().toUpperCase();
  normalized["Valor Serviço"] = toNumber(normalized["Valor Serviço"]);
  normalized["Meses Recorrência"] = Math.max(0, Math.floor(toNumber(normalized["Meses Recorrência"])));
  normalized.Valor = toNumber(normalized.Valor);
  normalized["Status Venda"] = normalizeSaleStatus(normalized["Status Venda"]);
  normalized["MÊS VENDA"] = normalizeMonth(normalized["MÊS VENDA"] || monthFromDate(normalized.Data));
  normalized.Ano = String(normalized.Ano || yearFromDate(normalized.Data) || new Date().getFullYear());
  normalized.Quarter = String(normalized.Quarter || quarterFromMonthAbbrev(normalized["MÊS VENDA"]));
  normalized.Canal = normalizeChannelName(normalized.Entrada || normalized.Canal || "");
  normalized.Origem = normalizeOrigemValue(normalized.Origem || row.ORIGEM || normalized["Sócio?"]);
  normalized["Sócio?"] = origemToSocioFlag(normalized.Origem);

  if (normalized.Tipo === "R" && normalized["Valor Serviço"] > 0 && normalized["Meses Recorrência"] > 0) {
    normalized.Valor = normalized["Valor Serviço"] * normalized["Meses Recorrência"];
    normalized["Equação"] = `${BRL.format(normalized["Valor Serviço"])} x ${normalized["Meses Recorrência"]} mês(es)`;
  } else {
    normalized["Equação"] = normalized.Tipo === "R" ? "Valor informado manualmente (sem base completa de recorrência)." : "-";
  }

  if (!normalized["MÊS VENDA"] && normalized.Data) normalized["MÊS VENDA"] = monthFromDate(normalized.Data);
  if (!normalized.Ano) normalized.Ano = String(new Date().getFullYear());
  if (!normalized.Quarter) normalized.Quarter = quarterFromMonthAbbrev(normalized["MÊS VENDA"]);
  if (!nonEmpty(normalized.Cliente) || !nonEmpty(normalized["Serviço"])) return null;
  if (normalized.Tipo === "R" && normalized.Valor <= 0) return null;
  return normalized;
}

function generateLeadId() {
  return `LD-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
}

function stableLeadIdFromFields(data, nome, empresa, servico) {
  const slug = (value) => String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "-")
    .replace(/(^-|-$)/g, "");
  const base = [data, nome, empresa, servico].map(slug).filter(Boolean).join("-");
  return base ? `LD-${base}` : generateLeadId();
}

function weekLabelFromDate(iso) {
  return weekOfMonth(iso);
}

function normalizeLeadStage(value) {
  const text = String(value || "").trim();
  if (KANBAN_STAGES.includes(text)) return text;
  const upper = text.toUpperCase();
  if (upper === "LEAD") return "Lead";
  if (upper === "PROPOSTA") return "Proposta";
  if (upper === "VENDA") return "Venda";
  return "";
}

function normalizeLeadRecord(row) {
  const data = String(row.Data || "").slice(0, 10);
  const nome = String(row.Nome || "").trim();
  const empresa = String(row.Empresa || row.Cliente || "").trim();
  const unidade = String(row.Unidade || row.Filial || row.Planta || "").trim();
  const email = String(row.Email || "").trim();
  const telefone = String(row.Telefone || row.WhatsApp || row.Whatsapp || "").trim();
  const cargo = String(row.Cargo || "").trim();
  const canal = normalizeChannelName(row.Canal || row.Fonte || row.Entrada || "");
  const entrada = normalizeResponsibleLabel(row.Entrada || row.Responsável || "");
  const campanhaRaw = String(row.Campanha || "").trim();
  const campanha = campanhaRaw ? formatLeadEntradaLabel(campanhaRaw) : "N/A";
  const classificacao = String(row["Classificação"] || row.ICP || "").trim().toUpperCase();
  const servico = normalizeServiceName(row["Serviço"]);
  const tipo = String(row.TIPO || row.Tipo || "P").trim().toUpperCase();
  const propostaFlag = String(row["Proposta Enviada ?"] || "N").trim().toUpperCase() === "S" ? "S" : "N";
  const valorProposta = toNumber(row["Valor Proposta"]);
  const mesesRecorrencia = Math.max(0, Math.floor(toNumber(row["Meses Recorrência"] || row["Meses Recorrencia"] || 0)));
  const valorMensalRecorrente = toNumber(row["Valor Mensal Recorrente"] || row["Valor Serviço"] || 0);
  const dataEnvio = String(row["Data de envio"] || "").slice(0, 10);
  const dataVenda = String(row["Data da venda"] || row["Data Venda"] || "").slice(0, 10);
  const mes = normalizeMonth(row["Mês"] || monthFromDate(data));
  const ano = String(row.Ano || yearFromDate(data) || new Date().getFullYear());
  const quarter = String(row.Quarter || quarterFromMonthAbbrev(mes));
  const semanaEntrada = String(row["Semana Entrada"] || weekLabelFromDate(data));
  const semanaEnvio = String(row["Semana Envio"] || weekLabelFromDate(dataEnvio || data));
  const leadId = String(row.LeadId || stableLeadIdFromFields(data, nome, empresa, servico));
  const deletedRaw = String(row["Excluído?"] ?? row.Excluido ?? row.ExcluidoFlag ?? "").trim().toUpperCase();
  const deleted = row["Excluído?"] === true || deletedRaw === "S" || deletedRaw === "SIM" || deletedRaw === "TRUE" || deletedRaw === "1";

  if (!data || !nome || !empresa || !servico) return null;

  const historico = Array.isArray(row["Histórico Proposta"]) ? row["Histórico Proposta"] : [];
  const statusLeadRaw = normalizeLeadStage(row["Status Lead"]);
  const statusManual = normalizeLeadStage(row["Status Manual"]);
  const statusLead = statusLeadRaw || (propostaFlag === "S" ? "Proposta" : "Lead");
  return {
    LeadId: leadId,
    Data: data,
    Nome: nome,
    Empresa: empresa,
    Unidade: unidade,
    Email: email,
    Telefone: telefone,
    Cargo: cargo,
    Canal: canal,
    Fonte: canal,
    Campanha: campanha,
    "Classificação": classificacao || "B",
    Entrada: entrada,
    "Serviço": servico,
    "Proposta Enviada ?": propostaFlag,
    "Valor Proposta": valorProposta,
    "Meses Recorrência": mesesRecorrencia,
    "Valor Mensal Recorrente": valorMensalRecorrente,
    TIPO: tipo,
    "Data de envio": dataEnvio || data,
    "Data da venda": dataVenda,
    "Mês": mes,
    Ano: ano,
    Quarter: quarter,
    "Semana Entrada": semanaEntrada,
    "Semana Envio": semanaEnvio,
    "Status Lead": statusLead,
    "Status Manual": statusManual,
    "Vendido?": String(row["Vendido?"] || (statusLead === "Venda" ? "S" : "N")).toUpperCase() === "S" ? "S" : "N",
    "Valor Venda": toNumber(row["Valor Venda"] || 0),
    "Histórico Proposta": historico,
    "Histórico Status": Array.isArray(row["Histórico Status"]) ? row["Histórico Status"] : [],
    Tarefas: Array.isArray(row.Tarefas) ? row.Tarefas : [],
    "Histórico Tarefas": Array.isArray(row["Histórico Tarefas"]) ? row["Histórico Tarefas"] : [],
    "Observação Proposta": String(row["Observação Proposta"] || "").trim(),
    "Venda Match Id": String(row["Venda Match Id"] || "").trim(),
    "Excluído?": deleted,
  };
}

function upsertCustomLead(lead) {
  const idx = state.customGeral.findIndex((row) => row.LeadId === lead.LeadId);
  if (idx >= 0) state.customGeral[idx] = normalizeLeadRecord(lead);
  else state.customGeral.push(normalizeLeadRecord(lead));
  state.customGeral = state.customGeral.filter((row) => row !== null);
}

function resetLeadFormEditState() {
  state.editingLeadId = "";
  state.serviceSearch = "";
  setSelectedLeadServices([]);
  if (refs.recurringMonths) refs.recurringMonths.value = "";
  if (refs.recurringMonthlyValue) refs.recurringMonthlyValue.value = "";
  if (refs.recurringTotal) refs.recurringTotal.value = "";
  toggleRecurringFields();
  if (refs.leadServiceSearch) refs.leadServiceSearch.value = "";
  if (refs.leadFormSubmit) refs.leadFormSubmit.textContent = "Salvar Lead/Proposta";
  if (refs.cancelLeadEdit) refs.cancelLeadEdit.hidden = true;
}

function startLeadEdit(leadId) {
  const lead = getEditableLeads().find((row) => row.LeadId === leadId);
  if (!lead || !refs.leadForm) return;
  const setField = (name, value) => {
    const el = refs.leadForm.elements.namedItem(name);
    if (!el) return;
    el.value = valueOrDash(value) === "-" ? "" : String(value ?? "");
  };
  setField("Data", lead.Data);
  setField("Nome", lead.Nome);
  setField("Empresa", lead.Empresa);
  setField("Email", lead.Email);
  setField("Telefone", lead.Telefone);
  setField("Cargo", lead.Cargo);
  setField("Unidade", lead.Unidade);
  setField("Canal", lead.Canal);
  setField("Entrada", lead.Entrada);
  setField("Campanha", lead.Campanha);
  setField("Classificação", lead["Classificação"]);
  setSelectedLeadServices([lead["Serviço"]]);
  setField("TIPO", lead.TIPO || "P");
  setField("Meses Recorrência", lead["Meses Recorrência"] || "");
  setField("Valor Mensal Recorrente", lead["Valor Mensal Recorrente"] || "");
  setField("Proposta Enviada ?", lead["Proposta Enviada ?"] || "N");
  setField("Valor Proposta", toNumber(lead["Valor Proposta"]) > 0 ? toNumber(lead["Valor Proposta"]) : "");
  setField("Data de envio", lead["Data de envio"]);
  toggleRecurringFields();
  calculateRecurringProposalValue();

  state.editingLeadId = leadId;
  if (refs.leadFormSubmit) refs.leadFormSubmit.textContent = "Atualizar Lead/Proposta";
  if (refs.cancelLeadEdit) refs.cancelLeadEdit.hidden = false;
  refs.tabInput?.scrollIntoView({ behavior: "smooth", block: "start" });
}

function deleteLead(leadId) {
  const lead = getEditableLeads().find((row) => row.LeadId === leadId);
  if (!lead) return;
  const ok = confirm(`Excluir lead ${lead.Empresa} | ${lead.Nome}?`);
  if (!ok) return;

  const baseLeadIds = new Set(
    (state.raw?.records?.geral || [])
      .map((row) => normalizeLeadRecord(row))
      .filter((row) => row !== null)
      .map((row) => row.LeadId)
  );

  state.customGeral = state.customGeral.filter((row) => row.LeadId !== leadId);
  if (baseLeadIds.has(leadId)) {
    upsertCustomLead({ ...lead, "Excluído?": true });
  }

  if (state.editingLeadId === leadId) {
    refs.leadForm?.reset();
    resetLeadFormEditState();
  }

  persistCustomGeral();
  setDataMode("custom");
  clearFilterSets();
  rebuildChipFilters();
  applyFilters();
  renderLeadTable();
}

function persistCustom() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.customVendas));
}

function persistCustomGeral() {
  localStorage.setItem(STORAGE_KEY_GERAL, JSON.stringify(state.customGeral));
}

function loadCustomFromStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const arr = JSON.parse(raw);
    if (Array.isArray(arr)) state.customVendas = prepareVendasRecords(arr);
  } catch (_) {
    state.customVendas = [];
  }
}

function loadCustomGeralFromStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_GERAL);
    if (!raw) return;
    const arr = JSON.parse(raw);
    if (Array.isArray(arr)) {
      state.customGeral = arr.map((row) => normalizeLeadRecord(row)).filter((row) => row !== null);
    }
  } catch (_) {
    state.customGeral = [];
  }
}

function persistCrmViewState() {
  localStorage.setItem(STORAGE_KEY_CRM_VIEW, JSON.stringify(state.crmView));
}

function setDataMode(mode) {
  const normalized = mode === "custom" ? "custom" : "original";
  state.mode = normalized;
  localStorage.setItem(STORAGE_KEY_MODE, normalized);
}

function loadDataMode() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_MODE);
    if (raw === "custom" || raw === "original") {
      state.mode = raw;
      return;
    }
  } catch (_) {
    // noop
  }
  state.mode = "original";
}

function loadCrmViewState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_CRM_VIEW);
    if (!raw) return;
    const parsed = JSON.parse(raw);
    if (typeof parsed?.kanbanHidden === "boolean") state.crmView.kanbanHidden = parsed.kanbanHidden;
    if (typeof parsed?.kanbanCompact === "boolean") state.crmView.kanbanCompact = parsed.kanbanCompact;
  } catch (_) {
    state.crmView = { kanbanHidden: false, kanbanCompact: true };
  }
}

function applyCrmViewState() {
  refs.crmWorkspace?.classList.toggle("kanban-hidden", state.crmView.kanbanHidden);
  refs.crmKanban?.classList.toggle("compact", state.crmView.kanbanCompact);
  if (refs.toggleKanbanVisibility) {
    refs.toggleKanbanVisibility.textContent = state.crmView.kanbanHidden ? "Mostrar Kanban" : "Ocultar Kanban";
  }
  if (refs.toggleKanbanCompact) {
    refs.toggleKanbanCompact.textContent = state.crmView.kanbanCompact ? "Modo Compacto" : "Modo Confortável";
  }
}

function renderCustomPreview() {
  if (!refs.customTable) return;
  const thead = refs.customTable.querySelector("thead");
  const tbody = refs.customTable.querySelector("tbody");
  const cols = ["Data", "Cliente", "Serviço", "Origem", "Status Venda", "Entrada", "Reativação?", "ICP", "Tipo", "Valor Serviço", "Meses Recorrência", "Equação", "Valor", "MÊS VENDA"];
  thead.innerHTML = "";
  tbody.innerHTML = "";

  const hr = document.createElement("tr");
  cols.forEach((c) => {
    const th = document.createElement("th");
    th.textContent = formatColumnLabel(c);
    hr.appendChild(th);
  });
  thead.appendChild(hr);

  for (const row of state.customVendas.slice(0, 200)) {
    const tr = document.createElement("tr");
    cols.forEach((c) => {
      const td = document.createElement("td");
      td.textContent = formatCellValue(c, row[c]);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  }
}

function setAdvancedAnalyticsVisible(visible) {
  if (!refs.advancedAnalyticsSection || !refs.toggleAdvancedAnalytics) return;
  refs.advancedAnalyticsSection.hidden = !visible;
  refs.advancedAnalyticsSection.classList.toggle("is-hidden", !visible);
  refs.toggleAdvancedAnalytics.textContent = visible ? "Ocultar" : "Mostrar";
  refs.toggleAdvancedAnalytics.setAttribute("aria-expanded", visible ? "true" : "false");
}

function leadSalesKey(empresa, servico, unidade = "") {
  return `${String(empresa || "").trim().toUpperCase()}|${String(servico || "").trim().toUpperCase()}|${String(unidade || "").trim().toUpperCase()}`;
}

function toDateOrNull(value) {
  const date = parseDateInput(String(value || "").slice(0, 10));
  return date || null;
}

function saleMatchId(row) {
  return [
    String(row?.Data || "").trim(),
    String(row?.Cliente || row?.Empresa || "").trim().toUpperCase(),
    String(row?.Unidade || "").trim().toUpperCase(),
    String(row?.["Serviço"] || "").trim().toUpperCase(),
    String(row?.Tipo || "").trim().toUpperCase(),
    toNumber(row?.Valor).toFixed(2),
  ].join("|");
}

function leadSalePairKey(empresa, servico) {
  return `${String(empresa || "").trim().toUpperCase()}|${String(servico || "").trim().toUpperCase()}`;
}

function isLeadLockedForManualDecision(lead) {
  const manualStage = normalizeLeadStage(lead?.["Status Manual"]);
  const soldFlag = String(lead?.["Vendido?"] || "").toUpperCase() === "S";
  const hasSaleDate = nonEmpty(lead?.["Data da venda"]);
  const hasSaleValue = toNumber(lead?.["Valor Venda"]) > 0;
  const hasMatchId = nonEmpty(lead?.["Venda Match Id"]);
  return Boolean(manualStage || soldFlag || hasSaleDate || hasSaleValue || hasMatchId);
}

function matchScoreLeadSale(lead, sale) {
  const leadValue = toNumber(lead["Valor Proposta"]);
  const saleValue = toNumber(sale["Valor"]);
  const baseValue = leadValue > 0 ? leadValue : saleValue;
  const valueDiff = Math.abs(baseValue - saleValue);

  const leadDate = toDateOrNull(lead.Data);
  const saleDate = toDateOrNull(sale.Data);
  const dateDiffDays = leadDate && saleDate ? Math.abs((saleDate.getTime() - leadDate.getTime()) / 86400000) : 999999;
  const saleTime = saleDate ? saleDate.getTime() : Number.MAX_SAFE_INTEGER;

  return [valueDiff, dateDiffDays, saleTime];
}

function buildLeadSaleAssignments(leadRows, vendasRows) {
  const sales = (vendasRows || []).map((row, idx) => ({
    row,
    idx,
    used: false,
    matchId: saleMatchId(row),
  })).sort((a, b) => String(a.row?.Data || "").localeCompare(String(b.row?.Data || ""), "pt-BR"));

  const normalizedLeads = (leadRows || [])
    .map((lead) => normalizeLeadRecord(lead))
    .filter((lead) => lead !== null);
  const assignedLeads = new Set();
  const assignments = new Map();

  const salesByMatchId = new Map(sales.map((sale) => [sale.matchId, sale]));

  const leadsByExact = new Map();
  const leadsByPair = new Map();
  for (const lead of normalizedLeads) {
    const empresa = String(lead.Empresa || lead.Cliente || "").trim();
    const servico = String(lead["Serviço"] || "").trim();
    const unidade = String(lead.Unidade || "").trim();
    if (!empresa || !servico) continue;
    const exactKey = leadSalesKey(empresa, servico, unidade);
    const pairKey = leadSalePairKey(empresa, servico);
    if (!leadsByExact.has(exactKey)) leadsByExact.set(exactKey, []);
    if (!leadsByPair.has(pairKey)) leadsByPair.set(pairKey, []);
    leadsByExact.get(exactKey).push(lead);
    leadsByPair.get(pairKey).push(lead);
  }

  // 1) Reserva vendas para leads já trabalhados manualmente
  // para evitar que um lead novo remapeie histórico existente.
  for (const lead of normalizedLeads) {
    if (!isLeadLockedForManualDecision(lead)) continue;
    const existingMatchId = String(lead["Venda Match Id"] || "").trim();
    if (existingMatchId && salesByMatchId.has(existingMatchId)) {
      const sale = salesByMatchId.get(existingMatchId);
      if (!sale.used) {
        sale.used = true;
        assignedLeads.add(lead.LeadId);
        assignments.set(lead.LeadId, sale);
        continue;
      }
    }
  }

  const lockedLeadIds = new Set(
    normalizedLeads
      .filter((lead) => isLeadLockedForManualDecision(lead))
      .map((lead) => lead.LeadId)
  );

  for (const sale of sales) {
    if (sale.used) continue;
    const empresa = String(sale.row.Cliente || sale.row.Empresa || "").trim();
    const servico = String(sale.row["Serviço"] || "").trim();
    const unidade = String(sale.row.Unidade || "").trim();
    if (!empresa || !servico) continue;

    const exactKey = leadSalesKey(empresa, servico, unidade);
    const pairKey = leadSalePairKey(empresa, servico);
    const baseCandidates = (unidade
      ? leadsByExact.get(exactKey)
      : (leadsByExact.get(exactKey)?.length ? leadsByExact.get(exactKey) : leadsByPair.get(pairKey))) || [];
    const saleDate = toDateOrNull(sale.row.Data);

    let bestLead = null;
    let bestScore = null;
    for (const lead of baseCandidates) {
      if (lockedLeadIds.has(lead.LeadId)) continue;
      if (assignedLeads.has(lead.LeadId)) continue;
      if (String(lead["Proposta Enviada ?"] || "").toUpperCase() !== "S") continue;
      const leadDate = toDateOrNull(lead.Data);
      if (leadDate && saleDate) {
        const leadMinusSaleDays = (leadDate.getTime() - saleDate.getTime()) / 86400000;
        if (leadMinusSaleDays > MAX_SALE_BEFORE_LEAD_DAYS) continue;
      }
      const score = matchScoreLeadSale(lead, sale.row);
      if (!bestScore || score[0] < bestScore[0]
        || (score[0] === bestScore[0] && score[1] < bestScore[1])
        || (score[0] === bestScore[0] && score[1] === bestScore[1] && score[2] < bestScore[2])) {
        bestLead = lead;
        bestScore = score;
      }
    }

    if (bestLead) {
      assignedLeads.add(bestLead.LeadId);
      sale.used = true;
      assignments.set(bestLead.LeadId, sale);
    }
  }

  return assignments;
}

function enrichLeadStatusFromSales(leadRows, vendasRows) {
  // Modo de contenção: evita remapeamento automático no front que poderia
  // alterar leads já trabalhados ao cadastrar novos leads.
  // O vínculo automático permanece apenas no processo de extração da base.
  if (!AUTO_CROSSMATCH_LEADS_FRONT) {
    return (leadRows || [])
      .map((lead) => normalizeLeadRecord(lead))
      .filter((row) => row !== null);
  }

  const assignments = buildLeadSaleAssignments(leadRows, vendasRows);
  return leadRows.map((lead) => {
    const normalized = normalizeLeadRecord(lead);
    if (!normalized) return null;
    const autoObs = String(normalized["Observação Proposta"] || "").toUpperCase().includes("VENDA CONFIRMADA AUTOMATICAMENTE POR CRUZAMENTO");
    const baseline = autoObs
      ? normalizeLeadRecord({
        ...normalized,
        "Vendido?": "N",
        "Valor Venda": 0,
        "Data da venda": "",
        "Status Lead": String(normalized["Proposta Enviada ?"] || "").toUpperCase() === "S" ? "Proposta" : "Lead",
      })
      : normalized;
    const assigned = assignments.get(baseline.LeadId);
    const manualStage = normalizeLeadStage(normalized["Status Manual"]);
    if (manualStage) {
      return normalizeLeadRecord({
        ...baseline,
        "Status Lead": manualStage,
      });
    }
    if (!assigned) return normalizeLeadRecord({ ...baseline, "Venda Match Id": "" });
    const saleRow = assigned.row;
    const saleValue = toNumber(saleRow["Valor"]);
    return normalizeLeadRecord({
      ...baseline,
      "Status Lead": "Venda",
      "Vendido?": "S",
      "Valor Venda": saleValue > 0 ? saleValue : toNumber(baseline["Valor Venda"]),
      "Data da venda": baseline["Data da venda"] || saleRow?.Data || "",
      "Proposta Enviada ?": "S",
      "Data de envio": baseline["Data de envio"] || saleRow?.Data || baseline.Data,
      "Observação Proposta": baseline["Observação Proposta"] || "Venda confirmada automaticamente por cruzamento da base (GERAL x VENDAS) com vínculo 1:1.",
      "Venda Match Id": saleMatchId(saleRow),
    });
  }).filter((row) => row !== null);
}

function getEditableLeads() {
  const leads = getActiveGeral().map((row) => normalizeLeadRecord(row)).filter((row) => row !== null);
  return enrichLeadStatusFromSales(leads, getActiveVendas()).filter((row) => row["Excluído?"] !== true);
}

function isSaleMatchForLead(leadRow, vendaRow) {
  const empresa = String(leadRow.Empresa || "").trim().toUpperCase();
  const servico = String(leadRow["Serviço"] || "").trim().toUpperCase();
  const unidade = String(leadRow.Unidade || "").trim().toUpperCase();
  if (!empresa || !servico) return false;
  const sameCompany = String(vendaRow?.Cliente || "").trim().toUpperCase() === empresa;
  const sameService = String(vendaRow?.["Serviço"] || "").trim().toUpperCase() === servico;
  if (!sameCompany || !sameService) return false;
  if (!unidade) return true;
  return String(vendaRow?.Unidade || "").trim().toUpperCase() === unidade;
}

function getRelatedSalesForLead(leadRow, vendasRows, options = {}) {
  const { mode = "assigned" } = options;
  const matchId = String(leadRow?.["Venda Match Id"] || "").trim();
  if (matchId) {
    return (vendasRows || []).filter((row) => saleMatchId(row) === matchId);
  }
  if (mode === "all_matches") {
    return (vendasRows || []).filter((row) => isSaleMatchForLead(leadRow, row));
  }
  const leadId = String(leadRow?.LeadId || "").trim();
  if (!leadId) return [];
  const leadPool = getActiveGeral().map((row) => normalizeLeadRecord(row)).filter((row) => row !== null);
  const activeSales = (vendasRows || []).filter((row) => !isSaleCancelled(row));
  const assignments = buildLeadSaleAssignments(leadPool, activeSales);
  const assigned = assignments.get(leadId);
  if (!assigned?.row) return [];
  const assignedMatchId = saleMatchId(assigned.row);
  return (vendasRows || []).filter((row) => saleMatchId(row) === assignedMatchId);
}

function leadHasVenda(leadRow, vendasRows) {
  const relatedSales = getRelatedSalesForLead(leadRow, vendasRows);
  if (relatedSales.length > 0) return true;
  return String(leadRow["Vendido?"] || "").toUpperCase() === "S";
}

function leadHasProposta(leadRow, vendasRows) {
  if (String(leadRow["Proposta Enviada ?"] || "").toUpperCase() === "S") return true;
  return leadHasVenda(leadRow, vendasRows);
}

function passesCrmSharedFilters(lead, vendasRows) {
  if (state.filters.ano.size && !state.filters.ano.has(lead["Ano"])) return false;
  if (state.filters.quarter.size && !state.filters.quarter.has(lead["Quarter"])) return false;
  if (state.filters.mes.size && !state.filters.mes.has(lead["Mês"])) return false;
  if (state.filters.entrada.size && !state.filters.entrada.has(String(lead["Canal"] || "").trim())) return false;
  if (state.filters.servico.size && !state.filters.servico.has(lead["Serviço"])) return false;
  if (state.filters.tipo.size && !state.filters.tipo.has(String(lead["TIPO"] || "").toUpperCase())) return false;

  if (state.filters.origem.size) {
    const directOrigin = String(lead["Origem"] || "").trim();
    if (directOrigin && state.filters.origem.has(directOrigin)) return true;
    const relatedSales = getRelatedSalesForLead(lead, vendasRows);
    if (!relatedSales.length) return false;
    return relatedSales.some((sale) => state.filters.origem.has(String(sale["Origem"] || "").trim()));
  }
  return true;
}

function getFilteredLeads(leads, vendasRows) {
  const search = String(state.leadView.search || "").trim().toLowerCase();
  const stageFilter = state.leadView.stage || "TODOS";
  return leads.filter((lead) => {
    if (!passesCrmSharedFilters(lead, vendasRows)) return false;
    const stage = getLeadStage(lead, vendasRows);
    if (stageFilter !== "TODOS" && stage !== stageFilter) return false;
    if (!search) return true;
    const text = [
      lead.Nome,
      lead.Empresa,
      lead.Email,
      lead.Telefone,
      lead.Cargo,
      lead.Unidade,
      lead.Canal,
      lead["Serviço"],
      lead.Entrada,
      lead["Classificação"],
      stage,
    ].map((v) => String(v || "").toLowerCase()).join(" ");
    return text.includes(search);
  });
}

function stageBadgeClass(stage) {
  if (stage === "Venda") return "stage-badge venda";
  if (stage === "Proposta") return "stage-badge proposta";
  return "stage-badge lead";
}

function renderCrmOverview(leads, vendasRows) {
  const totalLeads = leads.length;
  const totalPropostas = leads.filter((lead) => leadHasProposta(lead, vendasRows)).length;
  const totalVendas = leads.filter((lead) => leadHasVenda(lead, vendasRows)).length;
  const conversao = totalPropostas > 0 ? (totalVendas / totalPropostas) * 100 : 0;
  const valorPropostas = leads.reduce((sum, lead) => sum + toNumber(lead["Valor Proposta"]), 0);
  const valorVendas = (vendasRows || []).reduce((sum, row) => sum + toNumber(row["Valor"]), 0);

  refs.crmTotalLeads.textContent = NUM.format(totalLeads);
  refs.crmTotalPropostas.textContent = NUM.format(totalPropostas);
  refs.crmTotalVendas.textContent = NUM.format(totalVendas);
  refs.crmTaxaConversao.textContent = `${conversao.toFixed(1).replace(".", ",")}%`;
  refs.crmValorPropostas.textContent = BRL.format(valorPropostas);
  refs.crmValorVendas.textContent = BRL.format(valorVendas);

  const stageCount = {
    Lead: totalLeads,
    Proposta: totalPropostas,
    Venda: totalVendas,
  };
  const maxVal = Math.max(1, stageCount.Lead, stageCount.Proposta, stageCount.Venda);
  refs.crmStageBars.innerHTML = "";
  ["Lead", "Proposta", "Venda"].forEach((stage) => {
    const row = document.createElement("div");
    row.className = "crm-stage-row";
    const label = document.createElement("span");
    label.className = "crm-stage-label";
    label.textContent = stage;
    const track = document.createElement("div");
    track.className = "crm-stage-track";
    const fill = document.createElement("div");
    fill.className = "crm-stage-fill";
    fill.style.width = `${(stageCount[stage] / maxVal) * 100}%`;
    track.appendChild(fill);
    const value = document.createElement("span");
    value.className = "crm-stage-value";
    value.textContent = `${NUM.format(stageCount[stage])} lead(s)`;
    row.appendChild(label);
    row.appendChild(track);
    row.appendChild(value);
    refs.crmStageBars.appendChild(row);
  });
}

function recordLeadStageHistory(lead, fromStage, toStage, source = "kanban") {
  const history = Array.isArray(lead["Histórico Status"]) ? [...lead["Histórico Status"]] : [];
  history.push({
    when: new Date().toISOString(),
    from: fromStage || "",
    to: toStage,
    source,
  });
  lead["Histórico Status"] = history;
}

function generateTaskId() {
  return `TK-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
}

function ensureLeadTaskContainers(lead) {
  if (!Array.isArray(lead.Tarefas)) lead.Tarefas = [];
  if (!Array.isArray(lead["Histórico Tarefas"])) lead["Histórico Tarefas"] = [];
}

function pushTaskHistory(lead, action, task, extra = {}) {
  ensureLeadTaskContainers(lead);
  lead["Histórico Tarefas"].push({
    when: new Date().toISOString(),
    action,
    taskId: task?.id || "",
    type: task?.tipo || "",
    owner: task?.responsavel || "",
    dueDate: task?.prazo || "",
    status: task?.status || "",
    ...extra,
  });
}

function nextTaskSuggestionByType(type) {
  const map = {
    "E-mail": "Agendar ligação de acompanhamento em até 2 dias.",
    "Ligação": "Registrar resumo e enviar e-mail com próximos passos.",
    "Reunião": "Enviar proposta ou resumo da reunião no mesmo dia.",
    "WhatsApp": "Agendar ligação para aprofundar o diagnóstico.",
  };
  return map[type] || "Defina a próxima ação para manter o lead em andamento.";
}

function normalizeTaskStatus(value) {
  const status = String(value || "").trim();
  if (["Pendente", "Em andamento", "Concluída", "Cancelada"].includes(status)) return status;
  return "Pendente";
}

function renderTaskHistory(lead) {
  if (!refs.leadTaskHistory) return;
  ensureLeadTaskContainers(lead);
  const items = [...lead["Histórico Tarefas"]].sort((a, b) => String(b.when).localeCompare(String(a.when), "pt-BR"));
  refs.leadTaskHistory.innerHTML = "";
  if (!items.length) {
    refs.leadTaskHistory.innerHTML = '<div class="task-history-item">Sem histórico de tarefas.</div>';
    return;
  }
  for (const item of items.slice(0, 120)) {
    const div = document.createElement("div");
    div.className = "task-history-item";
    const when = new Date(item.when).toLocaleString("pt-BR");
    div.textContent = `${when} | ${item.action} | ${item.type || "-"} | ${item.owner || "-"} | ${item.status || "-"}`;
    refs.leadTaskHistory.appendChild(div);
  }
}

function renderLeadTasks(lead) {
  if (!refs.leadTaskList) return;
  ensureLeadTaskContainers(lead);
  refs.leadTaskList.innerHTML = "";
  const tasks = [...lead.Tarefas].sort((a, b) => String(a.prazo || "").localeCompare(String(b.prazo || ""), "pt-BR"));
  if (!tasks.length) {
    refs.leadTaskList.innerHTML = '<div class="hint">Sem tarefas cadastradas.</div>';
    renderTaskHistory(lead);
    return;
  }

  for (const task of tasks) {
    const card = document.createElement("article");
    card.className = "task-card";
    const top = document.createElement("div");
    top.className = "task-card-top";
    const title = document.createElement("strong");
    title.textContent = `${task.tipo} • ${task.responsavel}`;
    const status = document.createElement("span");
    status.className = stageBadgeClass(task.status === "Concluída" ? "Venda" : task.status === "Em andamento" ? "Proposta" : "Lead");
    status.textContent = task.status;
    top.appendChild(title);
    top.appendChild(status);

    const p1 = document.createElement("p");
    p1.textContent = `Prazo: ${formatDateBr(task.prazo)}${task.descricao ? ` | ${task.descricao}` : ""}`;
    const actions = document.createElement("div");
    actions.className = "task-actions";

    const btnDone = document.createElement("button");
    btnDone.type = "button";
    btnDone.className = "ghost-btn";
    btnDone.textContent = "Concluir";
    btnDone.disabled = task.status === "Concluída";
    btnDone.addEventListener("click", () => updateTaskStatus(lead.LeadId, task.id, "Concluída"));

    const btnDoing = document.createElement("button");
    btnDoing.type = "button";
    btnDoing.className = "ghost-btn";
    btnDoing.textContent = "Em andamento";
    btnDoing.disabled = task.status === "Em andamento";
    btnDoing.addEventListener("click", () => updateTaskStatus(lead.LeadId, task.id, "Em andamento"));

    const btnCancel = document.createElement("button");
    btnCancel.type = "button";
    btnCancel.className = "ghost-btn";
    btnCancel.textContent = "Cancelar";
    btnCancel.disabled = task.status === "Cancelada";
    btnCancel.addEventListener("click", () => updateTaskStatus(lead.LeadId, task.id, "Cancelada"));

    actions.appendChild(btnDone);
    actions.appendChild(btnDoing);
    actions.appendChild(btnCancel);

    card.appendChild(top);
    card.appendChild(p1);
    card.appendChild(actions);
    refs.leadTaskList.appendChild(card);
  }
  renderTaskHistory(lead);
}

function buildLeadDataSnapshot(lead) {
  if (!refs.leadDataSnapshot) return;
  const rows = [
    ["Data", formatDateBr(lead.Data)],
    ["Nome", lead.Nome],
    ["Empresa", lead.Empresa],
    ["E-mail", lead.Email],
    ["Telefone", lead.Telefone],
    ["Cargo", lead.Cargo],
    ["Unidade", lead.Unidade],
    ["Canal", lead.Canal],
    ["Responsável", lead.Entrada],
    ["Serviço", lead["Serviço"]],
    ["Classificação", lead["Classificação"]],
    ["Tipo", lead.TIPO],
    ["Data de envio", formatDateBr(lead["Data de envio"])],
    ["Data da venda", formatDateBr(lead["Data da venda"])],
    ["Proposta Enviada?", lead["Proposta Enviada ?"]],
    ["Valor Proposta", BRL.format(toNumber(lead["Valor Proposta"]))],
    ["Valor Venda", BRL.format(toNumber(lead["Valor Venda"]))],
    ["Status Lead", lead["Status Lead"]],
  ];
  refs.leadDataSnapshot.innerHTML = "";
  for (const [label, value] of rows) {
    const item = document.createElement("div");
    item.className = "lead-data-item";
    const safeValue = valueOrDash(value);
    if (normalizeKey(label).includes("TELEFONE")) {
      const url = buildWhatsAppLink(value);
      if (url) {
        item.innerHTML = `<small>${label}</small><strong><a href="${url}" target="_blank" rel="noopener noreferrer">${safeValue}</a></strong>`;
      } else {
        item.innerHTML = `<small>${label}</small><strong>${safeValue}</strong>`;
      }
    } else {
      item.innerHTML = `<small>${label}</small><strong>${safeValue}</strong>`;
    }
    refs.leadDataSnapshot.appendChild(item);
  }
}

function renderLeadChangeHistory(lead) {
  if (!refs.leadChangeHistory) return;
  const statusHistory = Array.isArray(lead["Histórico Status"]) ? lead["Histórico Status"] : [];
  const propostaHistory = Array.isArray(lead["Histórico Proposta"]) ? lead["Histórico Proposta"] : [];
  const merged = [];
  for (const h of statusHistory) {
    merged.push({
      when: h.when,
      text: `Status: ${h.from || "-"} → ${h.to || "-"} (${h.source || "sistema"})`,
    });
  }
  for (const h of propostaHistory) {
    merged.push({
      when: h.when,
      text: `Venda/Proposta: ${BRL.format(toNumber(h.oldValue))} → ${BRL.format(toNumber(h.newValue))} | ${Number(h.discountPct || 0).toFixed(1)}% desconto`,
    });
  }
  merged.sort((a, b) => String(b.when).localeCompare(String(a.when), "pt-BR"));
  refs.leadChangeHistory.innerHTML = "";
  if (!merged.length) {
    refs.leadChangeHistory.innerHTML = '<div class="task-history-item">Sem histórico de alterações.</div>';
    return;
  }
  for (const item of merged.slice(0, 120)) {
    const line = document.createElement("div");
    line.className = "task-history-item";
    const when = item.when ? new Date(item.when).toLocaleString("pt-BR") : "-";
    line.textContent = `${when} | ${item.text}`;
    refs.leadChangeHistory.appendChild(line);
  }
}

function addLeadTask(leadId, payload) {
  const lead = getEditableLeads().find((row) => row.LeadId === leadId);
  if (!lead) return { ok: false, message: "Lead não encontrado." };
  ensureLeadTaskContainers(lead);
  if (!payload.responsavel || !payload.prazo) {
    return { ok: false, message: "Responsável e prazo são obrigatórios para criar tarefa." };
  }
  const task = {
    id: generateTaskId(),
    tipo: payload.tipo,
    responsavel: payload.responsavel,
    prazo: payload.prazo,
    status: normalizeTaskStatus(payload.status),
    descricao: String(payload.descricao || "").trim(),
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
  lead.Tarefas.push(task);
  pushTaskHistory(lead, "Tarefa criada", task);
  upsertCustomLead(lead);
  persistCustomGeral();
  setDataMode("custom");
  notifyTask("Nova tarefa criada", `${lead.Empresa} | ${task.tipo} | ${task.responsavel} | ${formatDateBr(task.prazo)}`);
  return { ok: true, task };
}

function updateTaskStatus(leadId, taskId, status) {
  const lead = getEditableLeads().find((row) => row.LeadId === leadId);
  if (!lead) return;
  ensureLeadTaskContainers(lead);
  const task = lead.Tarefas.find((t) => t.id === taskId);
  if (!task) return;
  const prev = task.status;
  task.status = normalizeTaskStatus(status);
  task.updatedAt = new Date().toISOString();
  if (task.status === "Concluída") task.completedAt = task.updatedAt;
  pushTaskHistory(lead, "Status de tarefa alterado", task, { fromStatus: prev, toStatus: task.status });
  upsertCustomLead(lead);
  persistCustomGeral();
  setDataMode("custom");
  if (task.status === "Concluída") {
    notifyTask("Tarefa concluída", `${lead.Empresa} | ${task.tipo} concluída por ${task.responsavel} (${formatDateBr(task.prazo)})`);
  }
  const suggestion = task.status === "Concluída" ? nextTaskSuggestionByType(task.tipo) : "";
  refs.taskSuggestion.textContent = suggestion ? `Próxima tarefa sugerida: ${suggestion}` : "";
  openLeadConversionModal(leadId);
}

function hasSaleForLead(lead) {
  return getRelatedSalesForLead(lead, getActiveVendas()).length > 0;
}

function ensureSaleFromLead(lead) {
  if (hasSaleForLead(lead)) return;
  const saleDate = String(lead["Data da venda"] || lead["Data de envio"] || new Date().toISOString().slice(0, 10));
  const finalValue = toNumber(lead["Valor Venda"]) > 0 ? toNumber(lead["Valor Venda"]) : toNumber(lead["Valor Proposta"]);
  if (finalValue <= 0) return;
  const vendaRow = normalizeVendaRecord({
    Data: saleDate,
    Cliente: lead.Empresa,
    Unidade: lead.Unidade,
    Serviço: lead["Serviço"],
    Entrada: lead.Canal || lead.Fonte || lead.Entrada,
    Origem: "Não Sócio",
    "Status Venda": "Ativa",
    ICP: lead["Classificação"],
    Tipo: lead.TIPO,
    Valor: finalValue,
    "MÊS VENDA": normalizeMonth(lead["Mês"] || monthFromDate(saleDate)),
  });
  if (!vendaRow) return;
  lead["Venda Match Id"] = saleMatchId(vendaRow);
  state.customVendas.push(vendaRow);
  state.customVendas = prepareVendasRecords(state.customVendas);
  persistCustom();
}

function markLeadSalesCancelled(lead, cancelDate, reason = "") {
  const mergedAll = getMergedVendasRaw();
  const relatedSales = getRelatedSalesForLead(lead, mergedAll).filter((sale) => !isSaleCancelled(sale));
  if (!relatedSales.length) return 0;

  let changed = 0;
  for (const sale of relatedSales) {
    const cancelledRow = normalizeVendaRecord({
      ...sale,
      "Status Venda": "Cancelada",
      "Data Cancelamento": String(cancelDate || new Date().toISOString().slice(0, 10)).slice(0, 10),
      "Motivo Cancelamento": String(reason || "").trim(),
    });
    if (!cancelledRow) continue;
    state.customVendas.push(cancelledRow);
    changed += 1;
  }
  if (changed > 0) {
    state.customVendas = prepareVendasRecords(state.customVendas);
    persistCustom();
  }
  return changed;
}

function reconcileLeadsFromSalesBase() {
  const allLeads = getEditableLeads();
  const rawSales = getMergedVendasRaw();
  let changed = 0;

  for (const lead of allLeads) {
    let related = getRelatedSalesForLead(lead, rawSales);
    if (!related.length) {
      // Fallback legado: busca correspondência direta para casos antigos sem Venda Match Id.
      related = getRelatedSalesForLead(lead, rawSales, { mode: "all_matches" });
    }
    if (!related.length) continue;
    const hasActive = related.some((sale) => !isSaleCancelled(sale));
    const hasCancelled = related.some((sale) => isSaleCancelled(sale));
    const currentStage = normalizeLeadStage(lead["Status Manual"] || lead["Status Lead"]);
    const soldFlag = String(lead["Vendido?"] || "").toUpperCase() === "S";

    if (hasActive) {
      if (currentStage === "Venda" && soldFlag) continue;
      const activeSale = related.find((sale) => !isSaleCancelled(sale));
      const updatedVenda = normalizeLeadRecord({
        ...lead,
        "Status Lead": "Venda",
        "Status Manual": "Venda",
        "Proposta Enviada ?": "S",
        "Vendido?": "S",
        "Data da venda": lead["Data da venda"] || activeSale?.Data || "",
        "Valor Venda": toNumber(lead["Valor Venda"]) > 0 ? toNumber(lead["Valor Venda"]) : toNumber(activeSale?.Valor || lead["Valor Proposta"]),
        "Venda Match Id": activeSale ? saleMatchId(activeSale) : String(lead["Venda Match Id"] || ""),
      });
      if (updatedVenda) {
        upsertCustomLead(updatedVenda);
        changed += 1;
      }
      continue;
    }

    if (!hasCancelled) continue;
    if (currentStage === "Proposta" && !soldFlag && toNumber(lead["Valor Venda"]) === 0) continue;

    const updated = normalizeLeadRecord({
      ...lead,
      "Status Lead": "Proposta",
      "Status Manual": "Proposta",
      "Vendido?": "N",
      "Data da venda": "",
      "Valor Venda": 0,
      "Venda Match Id": "",
      "Observação Proposta": `Reconciliação automática: venda cancelada detectada para ${lead.Empresa}.`,
    });
    if (updated) {
      upsertCustomLead(updated);
      changed += 1;
    }
  }

  if (changed > 0) {
    persistCustomGeral();
  }
  return changed;
}

function updateLeadStage(leadId, targetStage, source = "kanban") {
  const lead = getEditableLeads().find((row) => row.LeadId === leadId);
  if (!lead) return;
  const toStage = normalizeLeadStage(targetStage);
  if (!toStage) return;
  const fromStage = getLeadStage(lead, getActiveVendas());
  if (fromStage === toStage) return;

  lead["Status Lead"] = toStage;
  lead["Status Manual"] = toStage;
  if (toStage === "Lead") {
    lead["Proposta Enviada ?"] = "N";
    lead["Vendido?"] = "N";
    lead["Valor Venda"] = 0;
  } else if (toStage === "Proposta") {
    lead["Proposta Enviada ?"] = "S";
    lead["Vendido?"] = "N";
    lead["Data de envio"] = lead["Data de envio"] || new Date().toISOString().slice(0, 10);
    lead["Valor Venda"] = 0;
  } else if (toStage === "Venda") {
    lead["Proposta Enviada ?"] = "S";
    lead["Vendido?"] = "S";
    lead["Data de envio"] = lead["Data de envio"] || new Date().toISOString().slice(0, 10);
    lead["Data da venda"] = lead["Data da venda"] || new Date().toISOString().slice(0, 10);
    if (toNumber(lead["Valor Venda"]) <= 0) lead["Valor Venda"] = toNumber(lead["Valor Proposta"]);
    ensureSaleFromLead(lead);
  }

  recordLeadStageHistory(lead, fromStage, toStage, source);
  upsertCustomLead(lead);
  persistCustomGeral();
  setDataMode("custom");
  clearFilterSets();
  rebuildChipFilters();
  applyFilters();
  renderLeadTable();
}

function createKanbanCard(lead, stage) {
  const card = document.createElement("article");
  card.className = "kanban-card";
  card.draggable = true;
  card.dataset.leadId = lead.LeadId;
  card.dataset.stage = stage;

  const top = document.createElement("div");
  top.className = "kanban-card-top";
  const company = document.createElement("strong");
  company.textContent = valueOrDash(lead.Empresa);
  const type = document.createElement("span");
  type.className = stageBadgeClass(stage);
  type.textContent = stage;
  top.appendChild(company);
  top.appendChild(type);

  const name = document.createElement("p");
  name.className = "kanban-card-name";
  name.textContent = valueOrDash(lead.Nome);
  const meta = document.createElement("p");
  meta.className = "kanban-card-meta";
  const unidade = valueOrDash(lead.Unidade);
  meta.textContent = `${valueOrDash(lead["Serviço"])} | ${unidade} | ${valueOrDash(lead.Canal)}`;
  const value = document.createElement("p");
  value.className = "kanban-card-value";
  if (stage === "Venda") value.textContent = `Venda: ${BRL.format(toNumber(lead["Valor Venda"]))}`;
  else value.textContent = `Proposta: ${BRL.format(toNumber(lead["Valor Proposta"]))}`;
  const pendingTasks = (Array.isArray(lead.Tarefas) ? lead.Tarefas : [])
    .filter((task) => !["Concluída", "Cancelada"].includes(normalizeTaskStatus(task.status)));
  const pendingChip = document.createElement("span");
  pendingChip.className = "task-pending-chip";
  pendingChip.textContent = pendingTasks.length ? `Tarefas pendentes: ${pendingTasks.length}` : "Sem tarefas pendentes";

  card.appendChild(top);
  card.appendChild(name);
  card.appendChild(meta);
  card.appendChild(value);
  card.appendChild(pendingChip);

  card.addEventListener("dragstart", (event) => {
    event.dataTransfer?.setData("text/plain", lead.LeadId);
    event.dataTransfer.effectAllowed = "move";
    card.classList.add("dragging");
  });
  card.addEventListener("dragend", () => {
    card.classList.remove("dragging");
    refs.crmKanban.querySelectorAll(".kanban-dropzone.over").forEach((el) => el.classList.remove("over"));
  });
  card.addEventListener("click", () => openLeadConversionModal(lead.LeadId));
  return card;
}

function renderCrmKanban(leads, vendasRows) {
  if (!refs.crmKanban) return;
  applyCrmViewState();
  refs.crmKanban.innerHTML = "";
  const grouped = new Map(KANBAN_STAGES.map((stage) => [stage, []]));
  for (const lead of leads) {
    const stage = normalizeLeadStage(getLeadStage(lead, vendasRows)) || "Lead";
    grouped.get(stage).push(lead);
  }

  for (const stage of KANBAN_STAGES) {
    const column = document.createElement("section");
    column.className = "kanban-column";
    column.dataset.stage = stage;
    const head = document.createElement("div");
    head.className = "kanban-column-head";
    const title = document.createElement("h3");
    title.textContent = stage;
    const count = document.createElement("span");
    count.className = "kanban-count";
    count.textContent = `${grouped.get(stage).length}`;
    head.appendChild(title);
    head.appendChild(count);

    const dropzone = document.createElement("div");
    dropzone.className = "kanban-dropzone";
    dropzone.dataset.stage = stage;
    dropzone.addEventListener("dragover", (event) => {
      event.preventDefault();
      dropzone.classList.add("over");
    });
    dropzone.addEventListener("dragleave", () => {
      dropzone.classList.remove("over");
    });
    dropzone.addEventListener("drop", (event) => {
      event.preventDefault();
      dropzone.classList.remove("over");
      const leadId = event.dataTransfer?.getData("text/plain");
      if (!leadId) return;
      updateLeadStage(leadId, stage, "kanban_drag");
    });

    for (const lead of grouped.get(stage)) {
      dropzone.appendChild(createKanbanCard(lead, stage));
    }

    column.appendChild(head);
    column.appendChild(dropzone);
    refs.crmKanban.appendChild(column);
  }
}

function compareLeadValues(a, b) {
  const aNum = parseSortNumber(a);
  const bNum = parseSortNumber(b);
  if (aNum !== null && bNum !== null) return aNum - bNum;
  return String(a ?? "").localeCompare(String(b ?? ""), "pt-BR", { sensitivity: "base" });
}

function getSortedLeads(leads) {
  const sorted = [...leads];
  const { sortColumn, sortDirection } = state.leadView;
  if (!sortColumn) return sorted;
  const direction = sortDirection === "desc" ? -1 : 1;
  sorted.sort((a, b) => direction * compareLeadValues(a[sortColumn], b[sortColumn]));
  return sorted;
}

function renderLeadTable() {
  const thead = refs.leadTable.querySelector("thead");
  const tbody = refs.leadTable.querySelector("tbody");
  const cols = ["Data", "Nome", "Empresa", "Telefone", "Cargo", "Unidade", "Canal", "Serviço", "Proposta Enviada ?", "Valor Proposta", "Jornada", "Obs. Proposta", "TIPO", "Status Lead", "Situação Venda", "Ações"];
  const vendasAtivas = state.filteredVendas;
  thead.innerHTML = "";
  tbody.innerHTML = "";
  const filteredLeads = getFilteredLeads(getEditableLeads(), vendasAtivas);
  renderCrmOverview(filteredLeads, vendasAtivas);
  renderCrmKanban(filteredLeads, vendasAtivas);
  const leads = getSortedLeads(filteredLeads);
  const totalPages = Math.max(1, Math.ceil(leads.length / state.leadView.pageSize));
  state.leadView.page = Math.min(Math.max(1, state.leadView.page), totalPages);
  const start = (state.leadView.page - 1) * state.leadView.pageSize;
  const pageRows = leads.slice(start, start + state.leadView.pageSize);

  const hr = document.createElement("tr");
  cols.forEach((c) => {
    const th = document.createElement("th");
    th.textContent = formatColumnLabel(c);
    if (c !== "Ações" && state.leadView.sortColumn === c) th.classList.add(state.leadView.sortDirection === "asc" ? "sort-asc" : "sort-desc");
    if (c !== "Ações") th.addEventListener("click", () => {
      if (state.leadView.sortColumn === c) {
        state.leadView.sortDirection = state.leadView.sortDirection === "asc" ? "desc" : "asc";
      } else {
        state.leadView.sortColumn = c;
        state.leadView.sortDirection = "asc";
      }
      state.leadView.page = 1;
      renderLeadTable();
    });
    hr.appendChild(th);
  });
  thead.appendChild(hr);

  for (const row of pageRows) {
    const tr = document.createElement("tr");
    tr.className = "lead-row";
    tr.addEventListener("click", () => openLeadConversionModal(row.LeadId));
    cols.forEach((c) => {
      const td = document.createElement("td");
      if (c === "Ações") {
        const actions = document.createElement("div");
        actions.className = "lead-actions";
        const btnEdit = document.createElement("button");
        btnEdit.type = "button";
        btnEdit.className = "ghost-btn";
        btnEdit.textContent = "Editar";
        btnEdit.addEventListener("click", (event) => {
          event.stopPropagation();
          startLeadEdit(row.LeadId);
        });
        const btnDelete = document.createElement("button");
        btnDelete.type = "button";
        btnDelete.className = "ghost-btn";
        btnDelete.textContent = "Excluir";
        btnDelete.addEventListener("click", (event) => {
          event.stopPropagation();
          deleteLead(row.LeadId);
        });
        actions.appendChild(btnEdit);
        actions.appendChild(btnDelete);
        td.appendChild(actions);
      } else if (c === "Jornada") {
        td.appendChild(createLeadStageTrail(row, vendasAtivas));
      } else if (c === "Valor Proposta") {
        td.textContent = BRL.format(toNumber(row[c]));
      } else if (c === "Obs. Proposta") {
        td.textContent = valueOrDash(row["Observação Proposta"]);
      } else if (c === "Status Lead") {
        const badge = document.createElement("span");
        const stage = getLeadStage(row, vendasAtivas);
        badge.className = stageBadgeClass(stage);
        badge.textContent = stage;
        td.appendChild(badge);
      } else if (c === "Situação Venda") {
        const relatedAll = getRelatedSalesForLead(row, getMergedVendasRaw());
        const hasCancelled = relatedAll.some((sale) => isSaleCancelled(sale));
        const hasActive = relatedAll.some((sale) => !isSaleCancelled(sale));
        const badge = document.createElement("span");
        badge.className = "stage-badge";
        if (hasActive) {
          badge.classList.add("venda");
          badge.textContent = "Ativa";
        } else if (hasCancelled) {
          badge.classList.add("cancelada");
          badge.textContent = "Cancelada";
        } else {
          badge.classList.add("lead");
          badge.textContent = "-";
        }
        td.appendChild(badge);
      } else {
        if (normalizeKey(c).includes("TELEFONE")) {
          const url = buildWhatsAppLink(row[c]);
          if (url) {
            const link = document.createElement("a");
            link.href = url;
            link.target = "_blank";
            link.rel = "noopener noreferrer";
            link.textContent = formatCellValue(c, row[c]);
            td.appendChild(link);
          } else {
            td.textContent = formatCellValue(c, row[c]);
          }
        } else {
          td.textContent = formatCellValue(c, row[c]);
        }
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  }

  refs.leadPageInfo.textContent = `Página ${state.leadView.page} de ${totalPages}`;
  refs.leadPrevPage.disabled = state.leadView.page <= 1;
  refs.leadNextPage.disabled = state.leadView.page >= totalPages;
}

function inferLeadSoldFromVendas(leadRow, vendasRows) {
  return leadHasVenda(leadRow, vendasRows);
}

function getLeadStage(leadRow, vendasRows) {
  const explicit = normalizeLeadStage(leadRow["Status Manual"] || leadRow["Status Lead"]);
  if (explicit) return explicit;
  const proposta = String(leadRow["Proposta Enviada ?"] || "").toUpperCase() === "S";
  if (inferLeadSoldFromVendas(leadRow, vendasRows)) return "Venda";
  if (proposta) return "Proposta";
  return "Lead";
}

function createLeadStageTrail(leadRow, vendasRows) {
  const stage = getLeadStage(leadRow, vendasRows);
  const stageOrder = ["Lead", "Proposta", "Venda"];
  const idx = stageOrder.indexOf(stage);
  const wrap = document.createElement("div");
  wrap.className = "stage-track";

  stageOrder.forEach((name, i) => {
    const dot = document.createElement("span");
    dot.className = "stage-node";
    if (i < idx) dot.classList.add("done");
    if (i === idx) dot.classList.add(i === stageOrder.length - 1 ? "done" : "current");
    wrap.appendChild(dot);

    if (i < stageOrder.length - 1) {
      const line = document.createElement("span");
      line.className = "stage-line";
      if (i < idx) line.classList.add("done");
      wrap.appendChild(line);
    }
  });

  const label = document.createElement("span");
  label.className = "stage-label";
  label.textContent = stage;
  wrap.appendChild(label);
  return wrap;
}

function closeLeadConversionModal() {
  refs.leadModal.classList.remove("open");
  refs.leadModal.setAttribute("aria-hidden", "true");
  refs.leadModalOverlay.hidden = true;
  state.selectedLeadId = "";
  state.selectedLeadTaskId = "";
  if (refs.leadTaskForm) refs.leadTaskForm.reset();
  if (refs.taskSuggestion) refs.taskSuggestion.textContent = "";
}

function openLeadConversionModal(leadId) {
  const lead = getEditableLeads().find((r) => r.LeadId === leadId);
  if (!lead) return;
  state.selectedLeadId = leadId;
  state.selectedLeadTaskId = "";
  const historyCount = Array.isArray(lead["Histórico Status"]) ? lead["Histórico Status"].length : 0;
  refs.leadModalMeta.textContent = `${lead.Empresa} | ${valueOrDash(lead.Unidade)} | ${lead.Nome} | Proposta atual: ${BRL.format(toNumber(lead["Valor Proposta"]))} | Movimentações: ${historyCount}`;
  buildLeadDataSnapshot(lead);
  renderLeadChangeHistory(lead);
  renderLeadTasks(lead);
  refs.leadSoldFlag.value = String(lead["Vendido?"] || "N").toUpperCase() === "S" ? "S" : "N";
  refs.leadDiscountFlag.value = "N";
  refs.leadNewValueWrap.hidden = true;
  refs.leadNewValue.value = "";
  refs.leadSaleDate.value = String(lead["Data da venda"] || new Date().toISOString().slice(0, 10)).slice(0, 10);
  const relatedAll = getRelatedSalesForLead(lead, getMergedVendasRaw());
  const hasCancelled = relatedAll.some((sale) => isSaleCancelled(sale));
  refs.leadSaleStatus.value = hasCancelled && String(lead["Vendido?"] || "").toUpperCase() !== "S" ? "CANCELADA" : "ATIVA";
  refs.leadSaleSocio.value = "N";
  refs.leadSaleObs.value = "";
  toggleSaleFieldsByFlag();
  if (refs.taskSuggestion) refs.taskSuggestion.textContent = "";
  refs.leadModal.classList.add("open");
  refs.leadModal.setAttribute("aria-hidden", "false");
  refs.leadModalOverlay.hidden = false;
}

function toggleSaleFieldsByFlag() {
  const sold = refs.leadSoldFlag?.value === "S";
  const cancelled = refs.leadSaleStatus?.value === "CANCELADA";
  const showDiscount = sold && !cancelled;
  if (!showDiscount) {
    refs.leadDiscountFlag.value = "N";
    refs.leadNewValue.value = "";
  }
  refs.leadDiscountFlag.closest("label").hidden = !showDiscount;
  refs.leadNewValueWrap.hidden = !showDiscount || refs.leadDiscountFlag.value !== "S";
  refs.leadSaleDate.closest("label").hidden = !sold;
  refs.leadSaleStatus.closest("label").hidden = !sold;
  refs.leadSaleSocio.closest("label").hidden = !sold || cancelled;
  refs.leadSaleObs.closest("label").hidden = !sold;
}

function processLeadConversion() {
  const lead = getEditableLeads().find((r) => r.LeadId === state.selectedLeadId);
  if (!lead) return;
  const fromStageBefore = getLeadStage(lead, getActiveVendas());
  const sold = refs.leadSoldFlag.value === "S";
  const saleDate = String(refs.leadSaleDate.value || new Date().toISOString().slice(0, 10));
  const saleStatus = refs.leadSaleStatus.value === "CANCELADA" ? "CANCELADA" : "ATIVA";
  const obsRaw = String(refs.leadSaleObs.value || "").trim();
  const persistAndRefresh = () => {
    persistCustomGeral();
    setDataMode("custom");
    clearFilterSets();
    rebuildChipFilters();
    applyFilters();
    renderLeadTable();
    closeLeadConversionModal();
  };

  if (!sold && fromStageBefore !== "Venda") {
    closeLeadConversionModal();
    return;
  }

  if (!sold || saleStatus === "CANCELADA") {
    const cancelReason = obsRaw || "Venda cancelada manualmente no CRM.";
    const cancelledCount = markLeadSalesCancelled(lead, saleDate, cancelReason);
    const oldSaleValue = toNumber(lead["Valor Venda"] || lead["Valor Proposta"]);
    const historico = Array.isArray(lead["Histórico Proposta"]) ? [...lead["Histórico Proposta"]] : [];
    historico.push({
      when: new Date().toISOString(),
      oldValue: oldSaleValue,
      newValue: 0,
      discountPct: 0,
      mode: "cancel",
      note: `Venda cancelada${cancelledCount > 0 ? "" : " (sem venda vinculada encontrada na base ativa)"}. ${cancelReason}`,
    });

    lead["Proposta Enviada ?"] = "S";
    lead["Data de envio"] = lead["Data de envio"] || saleDate;
    lead["Observação Proposta"] = `Venda cancelada em ${formatDateBr(saleDate)}. ${cancelReason}`;
    lead["Status Lead"] = "Proposta";
    lead["Status Manual"] = "Proposta";
    lead["Vendido?"] = "N";
    lead["Data da venda"] = "";
    lead["Valor Venda"] = 0;
    lead["Venda Match Id"] = "";
    lead["Histórico Proposta"] = historico;
    recordLeadStageHistory(lead, fromStageBefore, "Proposta", "modal_cancelamento");
    upsertCustomLead(lead);
    persistAndRefresh();
    return;
  }

  const hasDiscount = refs.leadDiscountFlag.value === "S";
  const oldValue = toNumber(lead["Valor Proposta"]);
  let finalValue = oldValue;
  let discountPct = 0;
  if (hasDiscount || oldValue <= 0) {
    finalValue = toNumber(refs.leadNewValue.value);
    if (finalValue <= 0) {
      alert("Informe um valor válido para a venda.");
      return;
    }
    if (hasDiscount && oldValue > 0 && finalValue >= oldValue) {
      alert("Com desconto = Sim, o novo valor deve ser menor que o valor original da proposta.");
      return;
    }
  }
  if (oldValue > 0 && finalValue < oldValue) {
    discountPct = ((oldValue - finalValue) / oldValue) * 100;
  }

  const socioFlag = refs.leadSaleSocio.value === "S" ? "S" : "N";
  const origemVenda = socioFlag === "S" ? "Sócio" : "Não Sócio";
  const discountObs = hasDiscount && oldValue > 0
    ? `Atualização de proposta: ${BRL.format(oldValue)} -> ${BRL.format(finalValue)} (${discountPct.toFixed(1)}% desconto)`
    : "Venda com manutenção do valor original da proposta.";
  const obs = obsRaw ? `${discountObs} | Obs: ${obsRaw}` : discountObs;

  const historico = Array.isArray(lead["Histórico Proposta"]) ? [...lead["Histórico Proposta"]] : [];
  historico.push({
    when: new Date().toISOString(),
    oldValue,
    newValue: finalValue,
    discountPct: Number(discountPct.toFixed(1)),
    mode: hasDiscount ? "new" : "keep",
    note: obs,
  });

  lead["Proposta Enviada ?"] = "S";
  lead["Data de envio"] = lead["Data de envio"] || saleDate;
  lead["Valor Proposta"] = finalValue;
  lead["Observação Proposta"] = obs;
  lead["Status Lead"] = "Venda";
  lead["Status Manual"] = "Venda";
  lead["Vendido?"] = "S";
  lead["Data da venda"] = saleDate;
  lead["Valor Venda"] = finalValue;
  lead["Histórico Proposta"] = historico;
  recordLeadStageHistory(lead, fromStageBefore, "Venda", "modal_conversao");
  upsertCustomLead(lead);

  const vendaRow = normalizeVendaRecord({
    Data: saleDate,
    Cliente: lead.Empresa,
    Unidade: lead.Unidade,
    Serviço: lead["Serviço"],
    Entrada: lead.Canal || lead.Fonte || lead.Entrada,
    Origem: origemVenda,
    "Status Venda": "Ativa",
    ICP: lead["Classificação"],
    Tipo: lead.TIPO,
    Valor: finalValue,
    "MÊS VENDA": normalizeMonth(lead["Mês"] || monthFromDate(saleDate)),
  });
  if (vendaRow) {
    lead["Venda Match Id"] = saleMatchId(vendaRow);
    upsertCustomLead(lead);
    state.customVendas.push(vendaRow);
    state.customVendas = prepareVendasRecords(state.customVendas);
    persistCustom();
  }

  persistAndRefresh();
}

function updateDataModeInfo() {
  const customCount = state.customVendas.length;
  const customLeadCount = state.customGeral.length;
  const active = getActiveVendas().length;
  const modeLabel = state.mode === "custom" ? "Base combinada (Excel + editada)" : "Base original";
  refs.dataModeInfo.textContent = `${modeLabel} ativa | Vendas ativas: ${NUM.format(active)} | Vendas editadas: ${NUM.format(customCount)} | Leads editados: ${NUM.format(customLeadCount)}`;
}

function updateEntradaFilterButton() {
  const selected = [...state.filters.entrada];
  refs.entradaDropdownBtn.textContent = selected.length ? `Entrada(s): ${selected.join(", ")}` : "Selecionar entrada(s)";
}

function updateServicoFilterButton() {
  if (!refs.servicoDropdownBtn) return;
  const selected = [...state.filters.servico];
  refs.servicoDropdownBtn.textContent = selected.length ? `Serviço(s): ${selected.join(", ")}` : "Selecionar serviço(s)";
}

function renderEntradaDropdown(values, relatedSet = new Set()) {
  refs.entradaDropdown.innerHTML = "";
  const search = document.createElement("input");
  search.type = "search";
  search.className = "filter-dropdown-search";
  search.placeholder = "Buscar canal...";
  search.value = state.entradaSearch;
  search.addEventListener("input", () => {
    state.entradaSearch = search.value;
    renderEntradaDropdown(values, relatedSet);
  });
  refs.entradaDropdown.appendChild(search);

  const searchNorm = String(state.entradaSearch || "").trim().toLowerCase();
  const visibleValues = values.filter((v) => String(v || "").toLowerCase().includes(searchNorm));
  const hasAnyFilter = Object.values(state.filters).some((set) => set.size > 0);
  for (const value of visibleValues) {
    const label = document.createElement("label");
    label.className = "filter-option";
    if (!state.filters.entrada.has(value) && hasAnyFilter) {
      label.classList.add(relatedSet.has(value) ? "related" : "unrelated");
    }
    const input = document.createElement("input");
    input.type = "checkbox";
    input.value = value;
    input.checked = state.filters.entrada.has(value);
    input.addEventListener("change", () => {
      if (input.checked) state.filters.entrada.add(value);
      else state.filters.entrada.delete(value);
      applyFilters();
    });
    const text = document.createElement("span");
    text.textContent = value;
    label.appendChild(input);
    label.appendChild(text);
    refs.entradaDropdown.appendChild(label);
  }
  if (visibleValues.length === 0) {
    const empty = document.createElement("div");
    empty.className = "hint";
    empty.textContent = "Nenhum canal encontrado.";
    refs.entradaDropdown.appendChild(empty);
  }
  updateEntradaFilterButton();
}

function renderServicoDropdown(values, relatedSet = new Set()) {
  if (!refs.servicoDropdown) return;
  refs.servicoDropdown.innerHTML = "";
  const search = document.createElement("input");
  search.type = "search";
  search.className = "filter-dropdown-search";
  search.placeholder = "Buscar serviço...";
  search.value = state.servicoSearch;
  search.addEventListener("input", () => {
    state.servicoSearch = search.value;
    renderServicoDropdown(values, relatedSet);
  });
  refs.servicoDropdown.appendChild(search);

  const searchNorm = String(state.servicoSearch || "").trim().toLowerCase();
  const visibleValues = values.filter((v) => String(v || "").toLowerCase().includes(searchNorm));
  const hasAnyFilter = Object.values(state.filters).some((set) => set.size > 0);
  for (const value of visibleValues) {
    const label = document.createElement("label");
    label.className = "filter-option";
    if (!state.filters.servico.has(value) && hasAnyFilter) {
      label.classList.add(relatedSet.has(value) ? "related" : "unrelated");
    }
    const input = document.createElement("input");
    input.type = "checkbox";
    input.value = value;
    input.checked = state.filters.servico.has(value);
    input.addEventListener("change", () => {
      if (input.checked) state.filters.servico.add(value);
      else state.filters.servico.delete(value);
      applyFilters();
    });
    const text = document.createElement("span");
    text.textContent = value;
    label.appendChild(input);
    label.appendChild(text);
    refs.servicoDropdown.appendChild(label);
  }
  if (visibleValues.length === 0) {
    const empty = document.createElement("div");
    empty.className = "hint";
    empty.textContent = "Nenhum serviço encontrado.";
    refs.servicoDropdown.appendChild(empty);
  }
  updateServicoFilterButton();
}

async function importFromFile(file) {
  const name = file.name.toLowerCase();
  let records = [];

  if (name.endsWith(".json")) {
    const text = await file.text();
    const parsed = JSON.parse(text);
    if (Array.isArray(parsed)) records = parsed;
    else if (Array.isArray(parsed?.records?.vendas)) records = parsed.records.vendas;
  } else if (name.endsWith(".csv") || name.endsWith(".xlsx") || name.endsWith(".xls")) {
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array" });
    const sheetName = wb.SheetNames.includes("VENDAS") ? "VENDAS" : wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    records = XLSX.utils.sheet_to_json(ws, { defval: "" });
  }

  const normalized = records
    .map((row) => normalizeVendaRecord(row))
    .filter((row) => row !== null);

  if (normalized.length === 0) {
    alert("Não encontrei linhas válidas no arquivo. Verifique colunas como Cliente, Serviço e Valor.");
    return;
  }

  state.customVendas = prepareVendasRecords(normalized);
  persistCustom();
  setDataMode("custom");
  clearFilterSets();
  rebuildChipFilters();
  applyFilters();
}

function switchTab(tab) {
  refs.tabButtons.forEach((btn) => btn.classList.toggle("active", btn.dataset.tab === tab));
  refs.tabPainel.classList.toggle("active", tab === "painel");
  refs.tabInput.classList.toggle("active", tab === "input");
}

function exportExcel() {
  const rows = state.filteredVendas;
  const wb = XLSX.utils.book_new();

  const detailsRows = rows.map((r) => ({
    Data: r.Data,
    Cliente: r.Cliente,
    Serviço: r["Serviço"],
    Origem: r.Origem || "",
    "Status Venda": r["Status Venda"] || "Ativa",
    Entrada: r.Canal || r.Entrada,
    "Reativação?": r["Reativação?"],
    ICP: r.ICP,
    Tipo: r.Tipo,
    "Valor Serviço": toNumber(r["Valor Serviço"]),
    "Meses Recorrência": r["Meses Recorrência"] || "",
    Equação: r["Equação"] || "",
    Valor: toNumber(r.Valor),
    "MÊS VENDA": r["MÊS VENDA"],
  }));

  const summary = [
    { Métrica: "Vendas Filtradas", Valor: rows.length },
    { Métrica: "Valor Total", Valor: rows.reduce((s, r) => s + toNumber(r.Valor), 0) },
    { Métrica: "Ticket Médio", Valor: rows.length ? rows.reduce((s, r) => s + toNumber(r.Valor), 0) / rows.length : 0 },
    { Métrica: "% Recorrente", Valor: rows.length ? rows.filter((r) => String(r.Tipo).toUpperCase() === "R").length / rows.length : 0 },
    { Métrica: "Filtros Ano", Valor: [...state.filters.ano].join(", ") || "Todos" },
    { Métrica: "Filtros Quarter", Valor: [...state.filters.quarter].join(", ") || "Todos" },
    { Métrica: "Filtros Mês", Valor: [...state.filters.mes].join(", ") || "Todos" },
    { Métrica: "Filtros Origem", Valor: [...state.filters.origem].join(", ") || "Todos" },
    { Métrica: "Filtros Entrada", Valor: [...state.filters.entrada].join(", ") || "Todos" },
    { Métrica: "Filtros Serviço", Valor: [...state.filters.servico].join(", ") || "Todos" },
    { Métrica: "Filtros Tipo", Valor: [...state.filters.tipo].join(", ") || "Todos" },
  ];

  const byMonth = groupSum(rows, "MÊS VENDA", "Valor").map((x) => ({ Mês: x.label, Valor: x.value }));
  const byService = groupSum(rows, "Serviço", "Valor").map((x) => ({ Serviço: x.label, Valor: x.value }));

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "Resumo");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(detailsRows), "Detalhamento");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(byMonth), "Vendas_Mes");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(byService), "Vendas_Servico");

  const stamp = new Date().toISOString().replace(/[:T]/g, "-").slice(0, 16);
  XLSX.writeFile(wb, `Relatorio_KPI_Vendas_${stamp}.xlsx`);
}

async function exportPdf() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF("p", "mm", "a4");
  const pageW = doc.internal.pageSize.getWidth();

  doc.setFont("helvetica", "bold");
  doc.setFontSize(14);
  doc.text("Relatório KPI Comercial - ElevaLife", 14, 14);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.text(`Gerado em: ${new Date().toLocaleString("pt-BR")}`, 14, 20);
  doc.text(`Filtros: Ano[${[...state.filters.ano].join(", ") || "Todos"}] | Quarter[${[...state.filters.quarter].join(", ") || "Todos"}] | Mês[${[...state.filters.mes].join(", ") || "Todos"}]`, 14, 25);

  let y = 32;
  const rows = state.filteredVendas;
  const total = rows.reduce((s, r) => s + toNumber(r.Valor), 0);
  const ticket = rows.length ? total / rows.length : 0;
  const rec = rows.length ? rows.filter((r) => String(r.Tipo).toUpperCase() === "R").length / rows.length : 0;
  doc.text(`Vendas: ${NUM.format(rows.length)} | Valor: ${BRL.format(total)} | Ticket: ${BRL.format(ticket)} | %R: ${PCT.format(rec)}`, 14, y);
  y += 5;

  for (const chartId of chartIds) {
    const el = document.getElementById(chartId);
    if (!el) continue;
    try {
      const dataUrl = await Plotly.toImage(el, { format: "png", width: 1100, height: 620, scale: 1 });
      if (y > 250) {
        doc.addPage();
        y = 14;
      }
      doc.addImage(dataUrl, "PNG", 12, y, pageW - 24, 48);
      y += 54;
    } catch (err) {
      console.error("Erro exportando gráfico", chartId, err);
    }
  }

  const tableRows = rows.slice(0, 80).map((r) => [r.Data, r.Cliente, r["Serviço"], r.Origem || "-", r["Status Venda"] || "Ativa", r.Tipo, r["Equação"] || "-", BRL.format(toNumber(r.Valor))]);
  if (typeof doc.autoTable === "function") {
    doc.addPage();
    doc.setFontSize(11);
    doc.text("Detalhamento (primeiras 80 linhas)", 14, 14);
    doc.autoTable({
      startY: 18,
      head: [["Data", "Cliente", "Serviço", "Origem", "Status Venda", "Tipo", "Equação", "Valor"]],
      body: tableRows,
      styles: { fontSize: 8 },
      headStyles: { fillColor: [0, 123, 130] },
    });
  }

  const stamp = new Date().toISOString().replace(/[:T]/g, "-").slice(0, 16);
  doc.save(`Relatorio_KPI_Vendas_${stamp}.pdf`);
}

async function refreshSourceData() {
  const response = await fetch("/api/refresh", {
    method: "POST",
    cache: "no-store",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ reason: "manual_refresh" }),
  });

  if (!response.ok) {
    let message = "Falha ao atualizar a base.";
    try {
      const payload = await response.json();
      message = payload?.message || message;
    } catch (_) {
      if (response.status === 404) {
        message = "Endpoint /api/refresh não encontrado. Feche o servidor atual e abra pelo arquivo hostlocal.bat para habilitar atualização da planilha.";
      } else {
        message = `Erro HTTP ${response.status} ao atualizar a base.`;
      }
    }
    throw new Error(message);
  }

  const payload = await response.json();
  if (!payload?.ok) {
    throw new Error(payload?.message || "Não foi possível atualizar a base.");
  }
}

function wireEvents() {
  if (refs.btnRefreshData) {
    if (isLocalRuntime()) {
      refs.btnRefreshData.textContent = "Atualizar Tudo";
      refs.btnRefreshData.title = "Atualiza planilha (Excel) e recarrega a base.";
    } else {
      refs.btnRefreshData.textContent = "Atualizar Base Online";
      refs.btnRefreshData.title = "Recarrega a base publicada no ambiente online.";
    }
  }

  refs.tabButtons.forEach((btn) => {
    btn.addEventListener("click", () => switchTab(btn.dataset.tab));
  });

  refs.clearFilters.addEventListener("click", () => {
    const hadActiveFilters = Object.values(state.filters).some((set) => set.size > 0);
    clearFilterSets();
    rebuildChipFilters();
    applyFilters();
    const yearPinned = [...state.filters.ano][0] || getLatestVendasYear() || "";
    alert(hadActiveFilters ? `Filtros limpos. Ano padrão aplicado: ${yearPinned}.` : `Nenhum filtro ativo para limpar. Ano padrão aplicado: ${yearPinned}.`);
  });

  refs.toggleAdvancedAnalytics?.addEventListener("click", () => {
    const currentlyVisible = !refs.advancedAnalyticsSection?.hidden && !refs.advancedAnalyticsSection?.classList.contains("is-hidden");
    setAdvancedAnalyticsVisible(!currentlyVisible);
  });

  const runRefresh = async () => {
    const prevTextMain = refs.btnRefreshData.textContent;
    refs.btnRefreshData.disabled = true;
    refs.btnRefreshData.textContent = "Atualizando...";
    try {
      if (isLocalRuntime()) {
        await refreshSourceData();
        await loadData({ forceRefresh: true, showSuccessMessage: true });
        if (state.mode === "custom") {
          alert("Dados atualizados com sucesso. Base do Excel atualizada e dados manuais preservados na visualização combinada.");
        } else {
          alert("Dados atualizados com sucesso.");
        }
      } else {
        await loadData({ forceRefresh: true, showSuccessMessage: true });
        alert("Base online atualizada com sucesso.");
      }
    } catch (error) {
      refs.updatedAt.textContent = `Erro ao atualizar dados: ${error.message}`;
      console.error(error);
      alert(`Erro ao atualizar dados: ${error.message}`);
    } finally {
      refs.btnRefreshData.disabled = false;
      refs.btnRefreshData.textContent = prevTextMain;
    }
  };

  refs.btnRefreshData.addEventListener("click", async () => runRefresh());

  refs.toggleFilters.addEventListener("click", () => {
    const isOpen = refs.painelLayout.classList.toggle("filters-open");
    refs.toggleFilters.textContent = isOpen ? "-" : "+";
    setTimeout(() => window.dispatchEvent(new Event("resize")), 180);
  });

  refs.closeDetails.addEventListener("click", closeDetails);
  refs.overlay.addEventListener("click", closeDetails);
  refs.detailsPageSize.addEventListener("change", () => {
    state.detailsView.pageSize = Math.max(1, Number(refs.detailsPageSize.value) || 20);
    state.detailsView.page = 1;
    renderDetailsTable();
  });
  refs.detailsPrevPage.addEventListener("click", () => {
    state.detailsView.page = Math.max(1, state.detailsView.page - 1);
    renderDetailsTable();
  });
  refs.detailsNextPage.addEventListener("click", () => {
    state.detailsView.page = state.detailsView.page + 1;
    renderDetailsTable();
  });
  refs.btnExportExcel.addEventListener("click", exportExcel);
  refs.btnExportPdf.addEventListener("click", () => exportPdf());

  refs.fileInput.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      await importFromFile(file);
      alert(`Base editada carregada com ${state.customVendas.length} registros.`);
    } catch (err) {
      console.error(err);
      alert(`Erro ao importar: ${err.message}`);
    } finally {
      refs.fileInput.value = "";
    }
  });

  refs.btnUseOriginal.addEventListener("click", () => {
    setDataMode("original");
    clearFilterSets();
    rebuildChipFilters();
    applyFilters();
  });

  refs.btnUseCustom.addEventListener("click", () => {
    if (!state.customVendas.length && !state.customGeral.length) {
      alert("A base editada está vazia. Faça upload ou cadastre ao menos um lead/proposta no CRM.");
      return;
    }
    setDataMode("custom");
    clearFilterSets();
    rebuildChipFilters();
    applyFilters();
  });

  refs.btnClearCustom.addEventListener("click", () => {
    state.customVendas = [];
    state.customGeral = [];
    localStorage.removeItem(STORAGE_KEY);
    localStorage.removeItem(STORAGE_KEY_GERAL);
    refs.leadForm?.reset();
    resetLeadFormEditState();
    setDataMode("original");
    clearFilterSets();
    rebuildChipFilters();
    applyFilters();
    renderLeadTable();
  });

  refs.leadForm?.addEventListener("submit", (event) => {
    event.preventDefault();
    const selectedServices = getSelectedLeadServices();
    if (!selectedServices.length) {
      alert("Selecione um serviço na lista.");
      refs.leadServiceInput.reportValidity();
      return;
    }
    const formData = new FormData(refs.leadForm);
    const row = {};
    for (const [k, v] of formData.entries()) row[k] = v;

    if (String(row.TIPO || "").toUpperCase() === "R") {
      const recurringMonths = Math.max(1, Math.floor(toNumber(row["Meses Recorrência"])));
      const recurringMonthly = toNumber(row["Valor Mensal Recorrente"]);
      if (recurringMonths > 0 && recurringMonthly > 0) {
        row["Valor Proposta"] = recurringMonths * recurringMonthly;
      }
    }

    const servicesToPersist = state.editingLeadId ? [selectedServices[0]] : selectedServices;
    if (state.editingLeadId && selectedServices.length > 1) {
      alert("Na edição, apenas o primeiro serviço selecionado será aplicado. Para múltiplos serviços, cadastre um novo lead.");
    }

    const saved = [];
    for (let i = 0; i < servicesToPersist.length; i += 1) {
      const service = servicesToPersist[i];
      const payload = {
        ...row,
        "Serviço": service,
        LeadId: state.editingLeadId || generateLeadId(),
      };
      const normalized = normalizeLeadRecord(payload);
      if (!normalized) {
        continue;
      }
      upsertCustomLead(normalized);
      saved.push(normalized);
    }

    if (!saved.length) {
      alert("Preencha os campos obrigatórios do lead/proposta.");
      return;
    }

    persistCustomGeral();
    setDataMode("custom");
    refs.leadForm.reset();
    resetLeadFormEditState();
    toggleRecurringFields();
    clearFilterSets();
    rebuildChipFilters();
    applyFilters();
    renderLeadTable();
    updateLeadSuggestionLists();
    const companyLabel = saved[0]?.Empresa || "";
    alert(`Lead cadastrado com sucesso para ${companyLabel}. Serviços: ${saved.length}.`);
  });

  refs.cancelLeadEdit?.addEventListener("click", () => {
    refs.leadForm?.reset();
    resetLeadFormEditState();
  });

  refs.leadForm?.elements?.namedItem("TIPO")?.addEventListener("change", () => {
    toggleRecurringFields();
  });
  refs.recurringMonths?.addEventListener("input", () => {
    calculateRecurringProposalValue();
  });
  refs.recurringMonthlyValue?.addEventListener("input", () => {
    calculateRecurringProposalValue();
  });
  refs.leadEmpresaInput?.addEventListener("input", () => refreshCreateNewHintsFromDatalist());
  refs.leadNomeInput?.addEventListener("input", () => refreshCreateNewHintsFromDatalist());

  refs.leadServiceBtn?.addEventListener("click", (event) => {
    event.stopPropagation();
    renderLeadServiceOptions();
    refs.leadServiceDropdown.hidden = !refs.leadServiceDropdown.hidden;
  });
  refs.leadServiceDropdown?.addEventListener("click", (event) => event.stopPropagation());
  refs.leadServiceSearch?.addEventListener("input", () => {
    state.serviceSearch = refs.leadServiceSearch.value;
    renderLeadServiceOptions();
  });
  refs.leadServiceAdd?.addEventListener("click", () => {
    const candidate = normalizeServiceName(refs.leadServiceNew.value);
    if (!candidate) return;
    const existing = getServiceCatalog().find((service) => serviceKey(service) === serviceKey(candidate));
    const selected = getSelectedLeadServices();
    setSelectedLeadServices([...selected, existing || candidate]);
    refs.leadServiceNew.value = "";
    state.serviceSearch = "";
    if (refs.leadServiceSearch) refs.leadServiceSearch.value = "";
    renderLeadServiceOptions();
    refs.leadServiceDropdown.hidden = false;
  });

  refs.enableTaskNotifications?.addEventListener("click", async () => {
    if (!("Notification" in window)) {
      updateTaskNotificationStatus();
      return;
    }
    if (Notification.permission === "default") {
      await Notification.requestPermission();
    }
    updateTaskNotificationStatus();
    const leads = getEditableLeads();
    remindDueTasks(leads);
  });

  refs.toggleKanbanVisibility?.addEventListener("click", () => {
    state.crmView.kanbanHidden = !state.crmView.kanbanHidden;
    persistCrmViewState();
    applyCrmViewState();
  });

  refs.toggleKanbanCompact?.addEventListener("click", () => {
    state.crmView.kanbanCompact = !state.crmView.kanbanCompact;
    persistCrmViewState();
    applyCrmViewState();
  });

  refs.closeLeadModal?.addEventListener("click", closeLeadConversionModal);
  refs.leadModalOverlay?.addEventListener("click", closeLeadConversionModal);
  refs.leadSoldFlag?.addEventListener("change", () => {
    toggleSaleFieldsByFlag();
  });
  refs.leadSaleStatus?.addEventListener("change", () => {
    toggleSaleFieldsByFlag();
  });
  refs.leadDiscountFlag?.addEventListener("change", () => {
    toggleSaleFieldsByFlag();
  });
  refs.leadModalForm?.addEventListener("submit", (event) => {
    event.preventDefault();
    processLeadConversion();
  });
  refs.leadTaskForm?.addEventListener("submit", (event) => {
    event.preventDefault();
    if (!state.selectedLeadId) return;
    const payload = {
      tipo: refs.taskType.value,
      responsavel: String(refs.taskOwner.value || "").trim(),
      prazo: String(refs.taskDueDate.value || ""),
      status: refs.taskStatus.value,
      descricao: String(refs.taskDescription.value || "").trim(),
    };
    const result = addLeadTask(state.selectedLeadId, payload);
    if (!result.ok) {
      alert(result.message);
      return;
    }
    refs.leadTaskForm.reset();
    refs.taskStatus.value = "Pendente";
    refs.taskSuggestion.textContent = `Próxima tarefa sugerida: ${nextTaskSuggestionByType(payload.tipo)}`;
    openLeadConversionModal(state.selectedLeadId);
  });
  refs.entradaDropdownBtn?.addEventListener("click", (event) => {
    event.stopPropagation();
    refs.entradaDropdown.hidden = !refs.entradaDropdown.hidden;
  });
  refs.servicoDropdownBtn?.addEventListener("click", (event) => {
    event.stopPropagation();
    refs.servicoDropdown.hidden = !refs.servicoDropdown.hidden;
  });

  refs.leadPageSize?.addEventListener("change", () => {
    state.leadView.pageSize = Math.max(1, Number(refs.leadPageSize.value) || 20);
    state.leadView.page = 1;
    renderLeadTable();
  });
  refs.leadSearch?.addEventListener("input", () => {
    state.leadView.search = refs.leadSearch.value;
    state.leadView.page = 1;
    renderLeadTable();
  });
  refs.leadStageFilter?.addEventListener("change", () => {
    state.leadView.stage = refs.leadStageFilter.value || "TODOS";
    state.leadView.page = 1;
    renderLeadTable();
  });
  refs.leadPrevPage?.addEventListener("click", () => {
    state.leadView.page = Math.max(1, state.leadView.page - 1);
    renderLeadTable();
  });
  refs.leadNextPage?.addEventListener("click", () => {
    state.leadView.page = state.leadView.page + 1;
    renderLeadTable();
  });

  refs.entradaDropdown?.addEventListener("click", (event) => event.stopPropagation());
  refs.servicoDropdown?.addEventListener("click", (event) => event.stopPropagation());
  document.addEventListener("click", () => {
    if (refs.entradaDropdown) refs.entradaDropdown.hidden = true;
    if (refs.servicoDropdown) refs.servicoDropdown.hidden = true;
    if (refs.leadServiceDropdown) refs.leadServiceDropdown.hidden = true;
  });
}

async function loadData(options = {}) {
  const { forceRefresh = false, showSuccessMessage = false } = options;
  const cacheBuster = forceRefresh ? `?t=${Date.now()}` : "";
  const response = await fetch(`./data/kpi_data.json${cacheBuster}`, { cache: "no-store" });
  if (!response.ok) throw new Error("Não foi possível carregar data/kpi_data.json");

  state.raw = await response.json();
  state.raw.records.vendas = prepareVendasRecords(state.raw.records.vendas || []);
  state.raw.records.geral = prepareGeralRecords(state.raw.records.geral || []);
  const sourceWhen = new Date(state.raw.generatedAt);
  const refreshWhen = new Date();

  loadCustomFromStorage();
  loadCustomGeralFromStorage();
  loadDataMode();
  loadCrmViewState();
  // Mantém o modo escolhido pelo usuário entre recargas.
  // Se estiver em custom sem dados customizados, retorna ao original.
  if (state.mode === "custom" && !state.customVendas.length && !state.customGeral.length) {
    setDataMode("original");
  }
  const modeLabel = state.mode === "custom" ? "Base combinada ativa" : "Base original ativa";
  refs.updatedAt.textContent = showSuccessMessage
    ? `Dados atualizados em ${refreshWhen.toLocaleString("pt-BR")} | Fonte: ${sourceWhen.toLocaleString("pt-BR")} | ${modeLabel}`
    : `Base atualizada em: ${sourceWhen.toLocaleString("pt-BR")} | ${modeLabel}`;
  const reconciliationChanges = reconcileLeadsFromSalesBase();
  if (reconciliationChanges > 0) {
    refs.updatedAt.textContent += ` | Reconciliação automática: ${reconciliationChanges} ajuste(s)`;
  }
  // Evita filtros "presos" entre atualizações de base.
  clearFilterSets();
  updateTaskNotificationStatus();
  renderLeadServiceOptions();
  setSelectedLeadServices([]);
  updateLeadSuggestionLists();
  toggleRecurringFields();
  applyCrmViewState();
  rebuildChipFilters();
  applyFilters();
}

wireEvents();
setAdvancedAnalyticsVisible(false);
loadData({ forceRefresh: true }).catch((error) => {
  refs.updatedAt.textContent = `Erro ao carregar dados: ${error.message}`;
  console.error(error);
});
