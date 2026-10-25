import { createClient } from '@supabase/supabase-js';

const rawUrl = import.meta.env.VITE_SUPABASE_URL || '';
const url = rawUrl.trim().replace(/^['"]|['"]$/g, '').replace(/\/(rest\/v1|auth\/v1)(\/.*)?$/i, '').replace(/\/+$/, '');
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

export const supabaseConfigured = Boolean(url && anonKey);
export const supabase = supabaseConfigured ? createClient(url, anonKey) : null;
