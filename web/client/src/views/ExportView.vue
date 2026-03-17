<script setup>
import { computed, onMounted, ref } from "vue";
import { useRouter } from "vue-router";

const router = useRouter();

const user = ref(null);
const orderInput = ref("");
const afterDate = ref("2024-01-01");
const loading = ref(false);
const error = ref("");
const result = ref(null);

onMounted(async () => {
  const res = await fetch("/api/me");
  if (!res.ok) return router.push("/login");
  user.value = await res.json();
});

const orderIds = computed(() =>
  orderInput.value
    .split(/[\n,\s]+/)
    .map((s) => s.trim())
    .filter((s) => /^\d+$/.test(s))
);

const orderCount = computed(() => orderIds.value.length);

async function runExport() {
  if (!orderIds.value.length) return;

  error.value = "";
  result.value = null;
  loading.value = true;

  try {
    const after = afterDate.value.replaceAll("-", "/");
    const res = await fetch("/api/export", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ orderIds: orderIds.value, afterDate: after }),
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: "Export failed" }));
      throw new Error(err.error || "Export failed");
    }

    const missingHeader = res.headers.get("X-Missing-Orders") || "";
    const found = parseInt(res.headers.get("X-Found-Count") || "0");
    const missing = missingHeader ? missingHeader.split(",").filter(Boolean) : [];

    const blob = await res.blob();
    const downloadUrl = URL.createObjectURL(blob);

    result.value = { downloadUrl, found, missing };
  } catch (err) {
    error.value = err.message || "Something went wrong. Please try again.";
  } finally {
    loading.value = false;
  }
}

async function logout() {
  await fetch("/auth/logout", { method: "POST" });
  router.push("/login");
}

const today = new Date().toISOString().slice(0, 10);
</script>

<template>
  <div class="layout">
    <!-- Header -->
    <header class="header">
      <div class="header-inner">
        <div class="brand">
          <span class="brand-icon">🧾</span>
          <span class="brand-name">Invoice Exporter</span>
        </div>
        <div v-if="user" class="user-menu">
          <img v-if="user.picture" :src="user.picture" class="avatar" :alt="user.name" />
          <span class="user-name">{{ user.name }}</span>
          <button class="btn-signout" @click="logout">Sign out</button>
        </div>
      </div>
    </header>

    <!-- Main -->
    <main class="main">
      <div class="card">
        <div class="card-header">
          <h2>Export Order Invoices</h2>
          <p>Paste WooCommerce order IDs below — one per line — to generate a combined PDF of all matching invoices.</p>
        </div>

        <div class="form">
          <!-- Order IDs textarea -->
          <div class="field">
            <label class="field-label">
              Order IDs
              <span class="badge" :class="{ active: orderCount > 0 }">{{ orderCount }}</span>
            </label>
            <textarea
              v-model="orderInput"
              class="textarea"
              placeholder="3555263&#10;3568403&#10;3574185&#10;..."
              rows="10"
              :disabled="loading"
            />
          </div>

          <!-- After date -->
          <div class="field field-row">
            <label class="field-label" for="afterDate">Search emails after</label>
            <input
              id="afterDate"
              v-model="afterDate"
              type="date"
              class="date-input"
              :disabled="loading"
            />
          </div>

          <!-- Submit -->
          <button
            class="btn-export"
            :disabled="loading || orderCount === 0"
            @click="runExport"
          >
            <template v-if="loading">
              <span class="spinner" />
              Processing {{ orderCount }} order{{ orderCount !== 1 ? "s" : "" }}&hellip;
            </template>
            <template v-else>
              Export PDF
            </template>
          </button>
        </div>

        <!-- Error -->
        <div v-if="error" class="error-box">
          <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
          {{ error }}
        </div>

        <!-- Result -->
        <div v-if="result" class="result-box">
          <div class="result-stats">
            <div class="stat stat-found">
              <span class="stat-num">{{ result.found }}</span>
              <span class="stat-label">invoice{{ result.found !== 1 ? "s" : "" }} found</span>
            </div>
            <div v-if="result.missing.length" class="stat stat-missing">
              <span class="stat-num">{{ result.missing.length }}</span>
              <span class="stat-label">not found</span>
            </div>
          </div>

          <a
            :href="result.downloadUrl"
            :download="`orders-export-${today}.pdf`"
            class="btn-download"
          >
            <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            Download PDF
          </a>

          <details v-if="result.missing.length" class="missing-details">
            <summary>Missing orders ({{ result.missing.length }})</summary>
            <ul class="missing-list">
              <li v-for="id in result.missing" :key="id">{{ id }}</li>
            </ul>
          </details>
        </div>
      </div>
    </main>
  </div>
</template>

<style scoped>
/* Layout */
.layout {
  min-height: 100vh;
  display: flex;
  flex-direction: column;
}

/* Header */
.header {
  background: #0f172a;
  color: #fff;
  padding: 0 24px;
  height: 60px;
  display: flex;
  align-items: center;
  flex-shrink: 0;
}

.header-inner {
  max-width: 780px;
  width: 100%;
  margin: 0 auto;
  display: flex;
  align-items: center;
  justify-content: space-between;
}

.brand {
  display: flex;
  align-items: center;
  gap: 10px;
}

.brand-icon {
  font-size: 20px;
}

.brand-name {
  font-size: 16px;
  font-weight: 600;
  letter-spacing: -0.01em;
}

.user-menu {
  display: flex;
  align-items: center;
  gap: 12px;
}

.avatar {
  width: 30px;
  height: 30px;
  border-radius: 50%;
  border: 2px solid rgba(255, 255, 255, 0.2);
}

.user-name {
  font-size: 14px;
  color: #cbd5e1;
}

.btn-signout {
  background: transparent;
  color: #94a3b8;
  border: 1px solid rgba(255, 255, 255, 0.15);
  border-radius: 6px;
  padding: 5px 12px;
  font-size: 13px;
  cursor: pointer;
  transition: color 0.15s, border-color 0.15s;
}

.btn-signout:hover {
  color: #fff;
  border-color: rgba(255, 255, 255, 0.4);
}

/* Main */
.main {
  flex: 1;
  display: flex;
  align-items: flex-start;
  justify-content: center;
  padding: 40px 24px;
}

.card {
  background: #fff;
  border-radius: 16px;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.07), 0 8px 24px rgba(0, 0, 0, 0.06);
  max-width: 680px;
  width: 100%;
  overflow: hidden;
}

.card-header {
  padding: 28px 32px 0;
}

.card-header h2 {
  font-size: 20px;
  font-weight: 700;
  color: #0f172a;
  margin-bottom: 8px;
}

.card-header p {
  font-size: 14px;
  color: #64748b;
  line-height: 1.6;
}

/* Form */
.form {
  padding: 24px 32px 32px;
  display: flex;
  flex-direction: column;
  gap: 20px;
}

.field {
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.field-row {
  flex-direction: row;
  align-items: center;
  gap: 12px;
}

.field-label {
  font-size: 13px;
  font-weight: 600;
  color: #374151;
  display: flex;
  align-items: center;
  gap: 8px;
}

.badge {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-width: 22px;
  height: 22px;
  padding: 0 6px;
  border-radius: 99px;
  font-size: 11px;
  font-weight: 700;
  background: #e2e8f0;
  color: #64748b;
  transition: background 0.2s, color 0.2s;
}

.badge.active {
  background: #dbeafe;
  color: #1d4ed8;
}

.textarea {
  width: 100%;
  border: 1.5px solid #e2e8f0;
  border-radius: 10px;
  padding: 12px 14px;
  font-size: 14px;
  font-family: "SF Mono", "Fira Code", "Fira Mono", monospace;
  line-height: 1.6;
  color: #1e293b;
  resize: vertical;
  transition: border-color 0.15s;
  outline: none;
}

.textarea:focus {
  border-color: #3b82f6;
  box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1);
}

.textarea:disabled {
  background: #f8fafc;
  color: #94a3b8;
}

.date-input {
  border: 1.5px solid #e2e8f0;
  border-radius: 8px;
  padding: 8px 12px;
  font-size: 14px;
  color: #1e293b;
  outline: none;
  transition: border-color 0.15s;
}

.date-input:focus {
  border-color: #3b82f6;
  box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1);
}

.btn-export {
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  background: #1d4ed8;
  color: #fff;
  border: none;
  border-radius: 10px;
  padding: 14px 24px;
  font-size: 15px;
  font-weight: 600;
  cursor: pointer;
  transition: background 0.15s, transform 0.1s;
}

.btn-export:hover:not(:disabled) {
  background: #1e40af;
}

.btn-export:active:not(:disabled) {
  transform: scale(0.99);
}

.btn-export:disabled {
  background: #cbd5e1;
  cursor: not-allowed;
}

/* Spinner */
.spinner {
  display: inline-block;
  width: 16px;
  height: 16px;
  border: 2.5px solid rgba(255, 255, 255, 0.4);
  border-top-color: #fff;
  border-radius: 50%;
  animation: spin 0.7s linear infinite;
  flex-shrink: 0;
}

@keyframes spin {
  to { transform: rotate(360deg); }
}

/* Error */
.error-box {
  display: flex;
  align-items: flex-start;
  gap: 10px;
  margin: 0 32px 28px;
  padding: 14px 16px;
  background: #fef2f2;
  color: #b91c1c;
  border: 1px solid #fecaca;
  border-radius: 10px;
  font-size: 14px;
  line-height: 1.5;
}

.error-box svg {
  flex-shrink: 0;
  margin-top: 1px;
}

/* Result */
.result-box {
  margin: 0 32px 32px;
  padding: 24px;
  background: #f0fdf4;
  border: 1px solid #bbf7d0;
  border-radius: 12px;
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.result-stats {
  display: flex;
  gap: 16px;
}

.stat {
  display: flex;
  align-items: baseline;
  gap: 6px;
}

.stat-num {
  font-size: 26px;
  font-weight: 800;
  line-height: 1;
}

.stat-label {
  font-size: 13px;
  font-weight: 500;
}

.stat-found .stat-num { color: #16a34a; }
.stat-found .stat-label { color: #15803d; }
.stat-missing .stat-num { color: #d97706; }
.stat-missing .stat-label { color: #b45309; }

.btn-download {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  background: #16a34a;
  color: #fff;
  border-radius: 8px;
  padding: 11px 20px;
  font-size: 14px;
  font-weight: 600;
  text-decoration: none;
  align-self: flex-start;
  transition: background 0.15s;
}

.btn-download:hover {
  background: #15803d;
}

/* Missing details */
.missing-details {
  font-size: 13px;
  color: #92400e;
}

.missing-details summary {
  cursor: pointer;
  font-weight: 600;
  color: #b45309;
  user-select: none;
}

.missing-list {
  margin-top: 8px;
  padding-left: 20px;
  list-style: disc;
  display: flex;
  flex-direction: column;
  gap: 4px;
  font-family: "SF Mono", "Fira Code", monospace;
}
</style>
