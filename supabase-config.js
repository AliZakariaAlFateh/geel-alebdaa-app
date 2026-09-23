// supabase-config.js
const SUPABASE_URL = 'https://jwuovphlsdevbsamwcjv.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_FCNkO48Px-sbdeBYYBf5oA_EOMLfXTM';

const _supabase = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// عميل منفصل لإنشاء الحسابات من لوحة الأدمن (بدون التأثير على جلسة الأدمن)
const _signupClient = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: {
    persistSession: false,
    autoRefreshToken: false,
    detectSessionInUrl: false,
    storageKey: 'sb-signup-temp'   
  }
});

async function getCurrentUser() {
  const { data: { session } } = await _supabase.auth.getSession();
  return session ? session.user : null;
}

async function getCurrentProfile() {
  const user = await getCurrentUser();
  if (!user) return null;
  const { data, error } = await _supabase
    .from('profiles').select('*').eq('id', user.id).single();
  if (error) return null;
  return data;
}

async function logout() {
  await _supabase.auth.signOut();
  window.location.href = 'login.html';
}

async function requireAuth() {
  const u = await getCurrentUser();
  if (!u) { window.location.href = 'login.html'; return null; }
  return u;
}

async function requireRole(roles) {
  const p = await getCurrentProfile();
  if (!p) { window.location.href = 'login.html'; return null; }
  if (!roles.includes(p.role)) {
    window.location.href = (p.role === 'teacher') ? 'dashboard.html' : 'admin.html';
    return null;
  }
  return p;
}