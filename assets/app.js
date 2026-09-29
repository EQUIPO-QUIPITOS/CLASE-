// Pestañas: alternan entre el formulario de emprendimientos y el de creadores
document.querySelectorAll("[data-tab]").forEach((btn) => {
  btn.addEventListener("click", () => {
    const id = btn.dataset.tab;
    document.querySelectorAll("[data-tab]").forEach((b) => {
      const on = b === btn;
      b.classList.toggle("active", on);
      b.setAttribute("aria-selected", on);
    });
    document.querySelectorAll(".panel").forEach((p) => (p.hidden = p.id !== id));
  });
});

// Envío: cada formulario declara su acción en data-action
document.querySelectorAll("form[data-action]").forEach((form) => {
  const status = form.querySelector(".status");
  const button = form.querySelector("button[type=submit]");

  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    status.className = "status";
    status.textContent = "";

    const data = Object.fromEntries(new FormData(form));
    const hp = data.website || ""; // campo trampa anti-spam
    delete data.website;

    if (!/^\S+@\S+\.\S+$/.test(data.correo || "")) {
      return showError("Escribe un correo válido, por ejemplo nombre@correo.com.");
    }
    if (String(data.whatsapp || "").replace(/\D/g, "").length < 8) {
      return showError("Escribe tu WhatsApp con código de país, por ejemplo +57 300 000 0000.");
    }

    button.disabled = true;
    button.textContent = "Enviando…";

    try {
      // text/plain evita el preflight de CORS con Apps Script
      const res = await fetch(CONFIG.API_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ action: form.dataset.action, data, hp }),
      });
      const json = await res.json();
      if (!json.ok) throw new Error(json.error || "Error desconocido");

      form.reset();
      status.className = "status ok";
      status.textContent = "Recibimos tu registro. Te escribiremos por WhatsApp en menos de 48 horas.";
    } catch (err) {
      showError("No pudimos enviar el formulario. Revisa tu conexión e inténtalo de nuevo.");
    } finally {
      button.disabled = false;
      button.textContent = form.dataset.cta;
    }
  });

  function showError(msg) {
    status.className = "status err";
    status.textContent = msg;
  }
});
