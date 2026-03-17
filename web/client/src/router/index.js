import { createRouter, createWebHistory } from "vue-router";
import ExportView from "../views/ExportView.vue";
import LoginView from "../views/LoginView.vue";

const router = createRouter({
  history: createWebHistory(),
  routes: [
    { path: "/login", component: LoginView },
    { path: "/", component: ExportView, meta: { requiresAuth: true } },
  ],
});

router.beforeEach(async (to) => {
  if (to.meta.requiresAuth) {
    try {
      const res = await fetch("/api/me");
      if (!res.ok) return "/login";
    } catch {
      return "/login";
    }
  }
  if (to.path === "/login") {
    try {
      const res = await fetch("/api/me");
      if (res.ok) return "/";
    } catch {
      // not authenticated, stay on login
    }
  }
});

export default router;
