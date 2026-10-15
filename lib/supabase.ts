import { createClient, SupabaseClient } from '@supabase/supabase-js';

let cachedSupabase: SupabaseClient | null = null;

export const getSupabase = () => {
  if (!cachedSupabase) {
    const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL || 'https://placeholder.supabase.co';
    const supabaseAnonKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY || 'placeholder-key';
    cachedSupabase = createClient(supabaseUrl, supabaseAnonKey);
  }
  return cachedSupabase;
};

// Mantém exportação de compatibilidade chamando a função de lazy load
export const supabase = {
  from: (table: string) => getSupabase().from(table),
};