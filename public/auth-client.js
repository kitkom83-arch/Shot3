async function checkAuth() {
  const r = await fetch("/auth/me");
  const data = await r.json();
  if (!data.loggedIn) {
    window.location.href = "/login";
  }
}

document.getElementById("logoutBtn")?.addEventListener("click", async () => {
  await fetch("/auth/logout", { method: "POST" });
  window.location.href = "/login";
});

checkAuth();
