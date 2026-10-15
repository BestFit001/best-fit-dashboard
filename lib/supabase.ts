import { createClient, SupabaseClient } from '@supabase/supabase-js';

let cachedSupabase: SupabaseClient | null = null;

export const getSupabase = () => {
  if (!cachedSupabase) {
    const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL || 'https://placeholder.supabase.co';
    const supabaseAnonKey = 
      process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY || 
      process.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY || 
      'placeholder-key';
    
    cachedSupabase = createClient(supabaseUrl, supabaseAnonKey);
  }
  return cachedSupabase;
};

// Objeto proxy para manter compatibilidade com supabase.from(...)
export const supabase = {
  from: (table: string) => getSupabase().from(table),
};