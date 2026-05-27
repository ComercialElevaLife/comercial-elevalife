const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwVCO5q3eiD-kpo81vZkK9N61GG8v7Ug5GvY2QqiNG3OfbJcEkvijNoljW8HHZyhwRKOg/exec";
const EBOOK_URL = "https://elevalife-my.sharepoint.com/:b:/g/personal/leonardo_elevalife_com_br/IQBzo0-PPIipRZn9TXQItNchAfV2Ph436NAuZmqrDDTrtOk";
const ORIGEM = "Evento presencial - QR Code e-book NR-1";

const form = document.getElementById("lead-form");
const leadCard = document.getElementById("lead-card");
const successCard = document.getElementById("success-card");
const submitButton = document.getElementById("submit-button");
const feedback = document.getElementById("feedback");
const phoneInput = document.getElementById("telefone");
const downloadButton = document.getElementById("download-button");

let isSubmitting = false;
let lastSubmitTs = 0;

if (EBOOK_URL && EBOOK_URL !== "INSERIR_LINK_DO_EBOOK_AQUI") {
  downloadButton.href = EBOOK_URL;
} else {
  downloadButton.href = "#";
  downloadButton.addEventListener("click", (event) => {
    event.preventDefault();
    showFeedback("Configure o link do e-book na constante EBOOK_URL antes de publicar.");
  });
}

phoneInput.addEventListener("input", (event) => {
  event.target.value = formatBrazilPhone(event.target.value);
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  if (isSubmitting) return;

  const now = Date.now();
  if (now - lastSubmitTs < 2500) {
    showFeedback("Aguarde um instante antes de enviar novamente.");
    return;
  }

  const validationError = validateForm();
  if (validationError) {
    showFeedback(validationError);
    return;
  }

  if (SCRIPT_URL === "INSERIR_URL_DO_GOOGLE_APPS_SCRIPT_AQUI") {
    showFeedback("Configure a URL do Google Apps Script antes de publicar.");
    return;
  }

  const nome = form.nome.value.trim();
  const sobrenome = form.sobrenome.value.trim();
  const nomeCompleto = `${nome} ${sobrenome}`.trim();

  const payload = {
    nome: nomeCompleto,
    telefone: form.telefone.value.trim(),
    email: form.email.value.trim(),
    empresa: form.empresa.value.trim(),
    cargo: form.cargo.value.trim(),
    colaboradores: form.colaboradores.value,
    consentimento: "Sim",
    origem: ORIGEM
  };

  setSubmittingState(true);
  showFeedback("Enviando seus dados...", true);

  try {
    const params = new URLSearchParams(payload);
    const response = await fetch(SCRIPT_URL, {
      method: "POST",
      body: params
    });

    const result = await safeJson(response);

    if (!response.ok) {
      throw new Error(result.message || "Não foi possível concluir o envio.");
    }

    if (result.status && result.status !== "success") {
      throw new Error(result.message || "Não foi possível concluir o envio.");
    }

    showSuccess();
    form.reset();
    phoneInput.value = "";
    lastSubmitTs = Date.now();
  } catch (error) {
    const message = error instanceof Error ? error.message : "Falha de conexão. Tente novamente.";
    showFeedback(`Não foi possível enviar agora. ${message}`);
  } finally {
    setSubmittingState(false);
  }
});

function validateForm() {
  const nome = form.nome.value.trim();
  const sobrenome = form.sobrenome.value.trim();
  const telefone = form.telefone.value.trim();
  const email = form.email.value.trim();
  const empresa = form.empresa.value.trim();
  const cargo = form.cargo.value.trim();
  const colaboradores = form.colaboradores.value;
  const consentimento = form.consentimento.checked;

  if (!nome || !sobrenome || !telefone || !email || !empresa || !cargo || !colaboradores) {
    return "Preencha todos os campos obrigatórios.";
  }

  if (!isValidEmail(email)) {
    return "Digite um e-mail corporativo válido.";
  }

  if (!isValidPhone(telefone)) {
    return "Digite um telefone/WhatsApp completo no formato (00) 00000-0000.";
  }

  if (!consentimento) {
    return "Para continuar, marque a autorização de contato.";
  }

  return null;
}

function setSubmittingState(value) {
  isSubmitting = value;
  submitButton.disabled = value;
  submitButton.textContent = value ? "Enviando..." : "Liberar e-book e brinde";
}

function showSuccess() {
  feedback.textContent = "";
  leadCard.classList.add("hidden");
  successCard.classList.remove("hidden");
  successCard.scrollIntoView({ behavior: "smooth", block: "start" });
}

function showFeedback(message, isSuccess = false) {
  feedback.textContent = message;
  feedback.classList.toggle("success", isSuccess);
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/i.test(email);
}

function isValidPhone(phone) {
  const digits = phone.replace(/\D/g, "");
  return digits.length === 11;
}

function formatBrazilPhone(value) {
  const digits = value.replace(/\D/g, "").slice(0, 11);
  if (digits.length <= 2) return digits ? `(${digits}` : "";
  if (digits.length <= 7) return `(${digits.slice(0, 2)}) ${digits.slice(2)}`;
  return `(${digits.slice(0, 2)}) ${digits.slice(2, 7)}-${digits.slice(7)}`;
}

async function safeJson(response) {
  const text = await response.text();
  try {
    return JSON.parse(text);
  } catch {
    return { status: response.ok ? "success" : "error", message: text };
  }
}
