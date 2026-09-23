// supabase-config.js
const SUPABASE_URL = 'https://jwuovphlsdevbsamwcjv.supabase.co';//'https://jwuovphlsdevbsamwcjv.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_FCNkO48Px-sbdeBYYBf5oA_EOMLfXTM';

// Initialize Supabase Client
const _supabase = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// Helper function to check session
async function getCurrentUser() {
  const { data: { session } } = await _supabase.auth.getSession();
  return session ? session.user : null;
}