<script setup>
import { useRoute } from "vue-router";

const route = useRoute();
const isUnauthorized = route.query.error === "unauthorized";
</script>

<template>
  <div class="login-page">
    <div class="card">
      <div class="logo">
        <span>🧾</span>
      </div>

      <h1>Invoice Exporter</h1>
      <p class="subtitle">
        Sign in with your Particle for Men Google account to export WooCommerce
        order invoices as PDF.
      </p>

      <div v-if="isUnauthorized" class="alert">
        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
        Access is restricted to <strong>@particleformen.com</strong> accounts.
      </div>

      <a href="/auth/google" class="google-btn">
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="20" height="20">
          <path d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" fill="#4285F4"/>
          <path d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" fill="#34A853"/>
          <path d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l3.66-2.84z" fill="#FBBC05"/>
          <path d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z" fill="#EA4335"/>
        </svg>
        Sign in with Google
      </a>
    </div>
  </div>
</template>

<style scoped>
.login-page {
  min-height: 100vh;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 24px;
  background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 100%);
}

.card {
  background: #fff;
  border-radius: 16px;
  padding: 48px 40px;
  max-width: 420px;
  width: 100%;
  text-align: center;
  box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
}

.logo {
  font-size: 48px;
  line-height: 1;
  margin-bottom: 20px;
}

h1 {
  font-size: 24px;
  font-weight: 700;
  color: #0f172a;
  margin-bottom: 12px;
}

.subtitle {
  font-size: 14px;
  color: #64748b;
  line-height: 1.6;
  margin-bottom: 28px;
}

.alert {
  display: flex;
  align-items: center;
  gap: 8px;
  background: #fef2f2;
  color: #b91c1c;
  border: 1px solid #fecaca;
  border-radius: 8px;
  padding: 12px 16px;
  font-size: 13px;
  text-align: left;
  margin-bottom: 20px;
}

.google-btn {
  display: inline-flex;
  align-items: center;
  gap: 12px;
  background: #fff;
  color: #3c4043;
  border: 1px solid #dadce0;
  border-radius: 8px;
  padding: 12px 24px;
  font-size: 15px;
  font-weight: 500;
  text-decoration: none;
  transition: box-shadow 0.15s, background 0.15s;
  width: 100%;
  justify-content: center;
}

.google-btn:hover {
  background: #f8f9fa;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.12);
}
</style>
